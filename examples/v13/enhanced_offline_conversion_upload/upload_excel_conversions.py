import argparse
import json
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime, timedelta, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from auth_helper import *
from openapi_client.models.campaign import *


DEFAULT_WORKBOOK = (
    r"Enhanced_Offline_Conversion_Template.xlsx"
)
REQUIRED_COLUMNS = {"Conversion Name", "Conversion Time"}
IDENTIFIER_COLUMNS = {
    "Microsoft Click ID",
    "Hashed Email Address",
    "Hashed Phone Number",
}
HASH_PATTERN = re.compile(r"^[0-9a-fA-F]{64}$")
TIME_ZONE_PATTERN = re.compile(r"TimeZone=([+-])(\d{2})(\d{2})")


def parse_arguments():
    parser = argparse.ArgumentParser(
        description="Validate and upload enhanced offline conversions from Excel."
    )
    parser.add_argument("--file", default=DEFAULT_WORKBOOK, help="Path to the Excel file.")
    parser.add_argument(
        "--sheet", default="Offline Conversions", help="Worksheet containing conversions."
    )
    parser.add_argument(
        "--account-id",
        type=int,
        help="Microsoft Advertising account ID. The first accessible account is used if omitted.",
    )
    parser.add_argument(
        "--conversion-name",
        help="Override the Conversion Name in every Excel row with an existing goal name.",
    )
    parser.add_argument(
        "--list-goals",
        action="store_true",
        help="List offline conversion goals for the selected account without uploading.",
    )
    parser.add_argument(
        "--batch-size",
        type=int,
        default=1000,
        choices=range(1, 1001),
        metavar="1-1000",
        help="Conversions per API request (default: 1000).",
    )
    parser.add_argument(
        "--upload",
        action="store_true",
        help="Upload after validation. Without this flag, the script only previews the data.",
    )
    return parser.parse_args()


def read_excel_rows(workbook_path, sheet_name):
    workbook = Path(workbook_path).expanduser().resolve()
    if not workbook.is_file():
        raise FileNotFoundError(f"Workbook not found: {workbook}")

    powershell = shutil.which("pwsh") or shutil.which("powershell")
    if not powershell:
        raise RuntimeError("PowerShell is required to read this protected Excel workbook.")

    script = r"""
param([string]$Path, [string]$SheetName)
$ErrorActionPreference = "Stop"
$excel = $null
$book = $null
$sheet = $null
$used = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Open($Path, 0, $true)
    $sheet = $book.Worksheets.Item($SheetName)
    $used = $sheet.UsedRange
    $rows = @()
    for ($row = 1; $row -le $used.Rows.Count; $row++) {
        $values = @()
        for ($column = 1; $column -le $used.Columns.Count; $column++) {
            $values += [string]$sheet.Cells.Item($row, $column).Text
        }
        $rows += ,$values
    }
    ConvertTo-Json -InputObject $rows -Depth 4 -Compress
}
finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit() }
    if ($used) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($used) }
    if ($sheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) }
    if ($book) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
    if ($excel) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
"""
    script_path = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".ps1", encoding="utf-8", delete=False
        ) as script_file:
            script_file.write(script)
            script_path = Path(script_file.name)

        result = subprocess.run(
            [
                powershell,
                "-NoProfile",
                "-NonInteractive",
                "-File",
                str(script_path),
                "-Path",
                str(workbook),
                "-SheetName",
                sheet_name,
            ],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8-sig",
        )
    finally:
        if script_path is not None:
            script_path.unlink(missing_ok=True)

    if result.returncode:
        message = result.stderr.strip() or result.stdout.strip()
        raise RuntimeError(f"Excel could not read the workbook: {message}")
    return json.loads(result.stdout)


def parse_time_zone(parameter_row):
    parameter_text = " ".join(parameter_row)
    match = TIME_ZONE_PATTERN.search(parameter_text)
    if not match:
        raise ValueError("Row 1 must contain a time zone such as 'TimeZone=+0800'.")

    sign = 1 if match.group(1) == "+" else -1
    hours = int(match.group(2))
    minutes = int(match.group(3))
    if hours > 23 or minutes > 59:
        raise ValueError(f"Invalid time zone in row 1: {match.group(0)}")
    return timezone(sign * timedelta(hours=hours, minutes=minutes))


def parse_conversion_time(value, source_time_zone):
    formats = (
        "%m/%d/%Y %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%dT%H:%M:%S",
    )
    for date_format in formats:
        try:
            local_time = datetime.strptime(value, date_format)
            return local_time.replace(tzinfo=source_time_zone).astimezone(timezone.utc)
        except ValueError:
            continue
    raise ValueError(f"Unsupported Conversion Time: {value!r}")


def optional_value(row, headers, column_name):
    value = row[headers[column_name]].strip()
    return value or None


def validate_hash(value, column_name):
    if value is None:
        return None
    if not HASH_PATTERN.fullmatch(value):
        raise ValueError(f"{column_name} must be a 64-character SHA-256 hex value.")
    return value.lower()


def load_conversions(workbook_path, sheet_name):
    rows = read_excel_rows(workbook_path, sheet_name)
    if len(rows) < 3:
        raise ValueError("The worksheet does not contain any conversion rows.")

    source_time_zone = parse_time_zone(rows[0])
    headers = {name.strip(): index for index, name in enumerate(rows[1]) if name.strip()}
    missing = REQUIRED_COLUMNS - headers.keys()
    if missing:
        raise ValueError(f"Missing required column(s): {', '.join(sorted(missing))}")
    if not IDENTIFIER_COLUMNS.intersection(headers):
        raise ValueError("At least one conversion identifier column is required.")

    conversions = []
    errors = []
    for excel_row_number, row in enumerate(rows[2:], start=3):
        row += [""] * (len(rows[1]) - len(row))
        if not any(cell.strip() for cell in row):
            continue

        try:
            conversion_name = optional_value(row, headers, "Conversion Name")
            conversion_time_text = optional_value(row, headers, "Conversion Time")
            if not conversion_name or not conversion_time_text:
                raise ValueError("Conversion Name and Conversion Time are required.")

            microsoft_click_id = (
                optional_value(row, headers, "Microsoft Click ID")
                if "Microsoft Click ID" in headers
                else None
            )
            hashed_email = validate_hash(
                optional_value(row, headers, "Hashed Email Address")
                if "Hashed Email Address" in headers
                else None,
                "Hashed Email Address",
            )
            hashed_phone = validate_hash(
                optional_value(row, headers, "Hashed Phone Number")
                if "Hashed Phone Number" in headers
                else None,
                "Hashed Phone Number",
            )
            if not any((microsoft_click_id, hashed_email, hashed_phone)):
                raise ValueError(
                    "Microsoft Click ID, Hashed Email Address, or Hashed Phone Number is required."
                )

            conversion_time = parse_conversion_time(
                conversion_time_text, source_time_zone
            )
            if conversion_time > datetime.now(timezone.utc):
                raise ValueError("Conversion Time cannot be in the future.")

            kwargs = {
                "conversion_name": conversion_name,
                "conversion_time": conversion_time,
                "microsoft_click_id": microsoft_click_id,
                "hashed_email_address": hashed_email,
                "hashed_phone_number": hashed_phone,
            }
            if "Conversion Value" in headers:
                value = optional_value(row, headers, "Conversion Value")
                if value is not None:
                    kwargs["conversion_value"] = float(value.replace(",", ""))
            if "Conversion Currency" in headers:
                currency = optional_value(row, headers, "Conversion Currency")
                if currency is not None:
                    if not re.fullmatch(r"[A-Za-z]{3}", currency):
                        raise ValueError("Conversion Currency must be a 3-letter code.")
                    kwargs["conversion_currency_code"] = currency.upper()

            conversions.append((excel_row_number, OfflineConversion(**kwargs)))
        except (ValueError, IndexError) as error:
            errors.append(f"Excel row {excel_row_number}: {error}")

    if errors:
        raise ValueError("Workbook validation failed:\n  " + "\n  ".join(errors))
    if not conversions:
        raise ValueError("No conversion rows were found.")
    return conversions


def select_account(authorization_data, requested_account_id):
    if requested_account_id is None or authorization_data.account_id == requested_account_id:
        return

    customer_service = ServiceClient(
        "CustomerManagementService",
        13,
        authorization_data=authorization_data,
        environment=ENVIRONMENT,
    )
    user = customer_service.GetUser(GetUserRequest(UserId=None)).User
    accounts = search_accounts_by_user_id(customer_service, user.Id)
    account = next(
        (item for item in accounts if int(item.Id) == requested_account_id), None
    )
    if account is None:
        raise ValueError(
            f"Account {requested_account_id} is not accessible to the authenticated user."
        )
    authorization_data.account_id = account.Id
    authorization_data.customer_id = account.ParentCustomerId


def upload_conversions(authorization_data, conversions, batch_size):
    campaign_service = ServiceClient(
        service="CampaignManagementService",
        version=13,
        authorization_data=authorization_data,
        environment=ENVIRONMENT,
    )
    validate_conversion_goals(campaign_service, authorization_data.account_id, conversions)
    failed_rows = []

    for start in range(0, len(conversions), batch_size):
        batch = conversions[start : start + batch_size]
        request = ApplyOfflineConversionsRequest(
            offline_conversions=[conversion for _, conversion in batch]
        )
        response = campaign_service.apply_offline_conversions(
            apply_offline_conversions_request=request
        )
        partial_errors = response.PartialErrors or []
        for error in partial_errors:
            index = error.Index
            excel_row = batch[index][0] if index is not None and index < len(batch) else "unknown"
            failed_rows.append(excel_row)
            print(
                f"Excel row {excel_row} failed: code={error.Code}, "
                f"message={error.Message}"
            )
        print(
            f"Uploaded batch {start // batch_size + 1}: "
            f"{len(batch) - len(partial_errors)}/{len(batch)} succeeded."
        )

    if failed_rows:
        raise RuntimeError(
            f"{len(failed_rows)} conversion(s) failed. Excel rows: "
            + ", ".join(map(str, failed_rows))
        )


def get_offline_conversion_goals(campaign_service):
    request = GetConversionGoalsByIdsRequest(
        conversion_goal_ids=None,
        conversion_goal_types=ConversionGoalType.OFFLINECONVERSION,
    )
    response = campaign_service.get_conversion_goals_by_ids(
        get_conversion_goals_by_ids_request=request
    )
    if response.PartialErrors:
        details = "; ".join(
            f"{error.Code}: {error.Message}" for error in response.PartialErrors
        )
        raise RuntimeError(f"Could not retrieve offline conversion goals: {details}")
    return [goal for goal in (response.ConversionGoals or []) if goal is not None]


def validate_conversion_goals(campaign_service, account_id, conversions):
    goals = get_offline_conversion_goals(campaign_service)
    existing_names = {goal.Name for goal in goals}
    requested_names = {conversion.conversion_name for _, conversion in conversions}
    missing_names = requested_names - existing_names
    if not missing_names:
        print(
            "Matched offline conversion goal(s): "
            + ", ".join(sorted(requested_names))
        )
        return

    available = "\n".join(
        f"  - {goal.Name} (ID: {goal.Id}, status: {goal.Status})" for goal in goals
    )
    if not available:
        available = "  (No offline conversion goals found in this account.)"
    raise ValueError(
        "The following Conversion Name value(s) do not exactly match an offline "
        f"conversion goal in account {account_id}: "
        + ", ".join(sorted(missing_names))
        + "\nAvailable offline conversion goals:\n"
        + available
        + "\nUpdate the Excel file or pass --conversion-name with one of the names above."
    )


def main():
    args = parse_arguments()
    if args.list_goals:
        authorization_data = AuthorizationData(
            account_id=None,
            customer_id=None,
            developer_token=DEVELOPER_TOKEN,
            authentication=None,
        )
        authenticate(authorization_data)
        select_account(authorization_data, args.account_id)
        campaign_service = ServiceClient(
            service="CampaignManagementService",
            version=13,
            authorization_data=authorization_data,
            environment=ENVIRONMENT,
        )
        goals = get_offline_conversion_goals(campaign_service)
        print(f"Offline conversion goals in account {authorization_data.account_id}:")
        if not goals:
            print("  (None)")
        for goal in goals:
            print(f"  - {goal.Name} (ID: {goal.Id}, status: {goal.Status})")
        return

    conversions = load_conversions(args.file, args.sheet)
    if args.conversion_name:
        for _, conversion in conversions:
            conversion.conversion_name = args.conversion_name

    print(f"Validated {len(conversions)} conversion(s) from {args.file}:")
    for excel_row, conversion in conversions:
        identifiers = []
        if conversion.microsoft_click_id:
            identifiers.append("Microsoft Click ID")
        if conversion.hashed_email_address:
            identifiers.append("hashed email")
        if conversion.hashed_phone_number:
            identifiers.append("hashed phone")
        print(
            f"  Row {excel_row}: {conversion.conversion_name}, "
            f"{conversion.conversion_time.isoformat()}, "
            f"{conversion.conversion_value} {conversion.conversion_currency_code}, "
            f"{' + '.join(identifiers)}"
        )

    if not args.upload:
        print("\nDry run only. Re-run with --upload to send these conversions.")
        return

    authorization_data = AuthorizationData(
        account_id=None,
        customer_id=None,
        developer_token=DEVELOPER_TOKEN,
        authentication=None,
    )
    authenticate(authorization_data)
    select_account(authorization_data, args.account_id)
    print(f"Uploading to account {authorization_data.account_id}...")
    upload_conversions(authorization_data, conversions, args.batch_size)
    print(f"Successfully uploaded all {len(conversions)} conversion(s).")


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print(f"Error: {error}")
        raise
