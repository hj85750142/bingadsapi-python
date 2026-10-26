from auth_helper import *
from openapi_client.models.customer import *
from suds import WebFault
import logging

logging.basicConfig(level=logging.INFO)
logging.getLogger('suds.client').setLevel(logging.DEBUG)
logging.getLogger('suds.transport.http').setLevel(logging.DEBUG)

def output_status_message(message):
    print(message)

def output_bing_ads_webfault_error(error):
    for field in ('ErrorCode', 'Code', 'Details', 'FieldPath', 'Message'):
        if hasattr(error, field):
            output_status_message(f"{field}: {getattr(error, field)}")
    output_status_message('')

def output_error_detail(error_detail, error_attribute_set):
    api_errors = error_detail
    for field in error_attribute_set:
        api_errors = getattr(api_errors, field, None)
    if api_errors is None:
        return False
    if isinstance(api_errors, list):
        for api_error in api_errors:
            output_bing_ads_webfault_error(api_error)
    else:
        output_bing_ads_webfault_error(api_errors)
    return True

def output_webfault_errors(ex):
    if not hasattr(ex.fault, 'detail'):
        raise Exception("Unknown WebFault") from ex

    error_attribute_sets = (
        ('ApiFault', 'OperationErrors', 'OperationError'),
        ('AdApiFaultDetail', 'Errors', 'AdApiError'),
        ('ApiFaultDetail', 'BatchErrors', 'BatchError'),
        ('ApiFaultDetail', 'OperationErrors', 'OperationError'),
        ('EditorialApiFaultDetail', 'BatchErrors', 'BatchError'),
        ('EditorialApiFaultDetail', 'EditorialErrors', 'EditorialError'),
        ('EditorialApiFaultDetail', 'OperationErrors', 'OperationError'),
    )

    for error_attribute_set in error_attribute_sets:
        if output_error_detail(ex.fault.detail, error_attribute_set):
            return

    if hasattr(ex.fault.detail, 'ExceptionDetail'):
        api_errors = ex.fault.detail.ExceptionDetail
        if isinstance(api_errors, list):
            for api_error in api_errors:
                output_status_message(api_error.Message)
        else:
            output_status_message(api_errors.Message)
        return

    raise Exception("Unknown WebFault") from ex

def main(authorization_data):
    try:
        # 1. Search for user accounts
        print("-----\nSearchUserAccounts:")

        predicate = Predicate(
            Field='UserId',
            Operator='Equals',
            Value='YOUR_USER_ID'  # Or use a valid UserId
        )

        paging = Paging(
            Index=0,
            Size=100
        )

        search_request = SearchAccountsRequest(
            Predicates=[predicate,predicate],
            PageInfo=paging
        )

        search_response = customer_service.search_accounts(search_request)

        assert search_response is not None

        print("UserAccounts:")
        user_accounts = search_response.Accounts

        if user_accounts:
            for account in user_accounts:
                print(f"\nAccount:")
                print(f"  ID: {account.Id}")
                print(f"  Name: {account.Name}")
                print(f"  Number: {account.Number}")
                print(f"  Parent Customer ID: {account.ParentCustomerId}")
                print(f"  Account Life Cycle Status: {account.AccountLifeCycleStatus}")

        assert user_accounts is not None and len(user_accounts) > 0

        # 2. Get customer pilot features for each unique parent customer ID
        print("\n-----\nGetCustomerPilotFeatures:")

        customer_ids = [account.ParentCustomerId for account in user_accounts if account.ParentCustomerId]
        distinct_customer_ids = list(set(customer_ids))

        assert len(distinct_customer_ids) > 0

        for customer_id in distinct_customer_ids:
            pilot_request = GetCustomerPilotFeaturesRequest(
                CustomerId=customer_id
            )

            pilot_response = customer_service.get_customer_pilot_features(pilot_request)
            feature_pilot_flags = pilot_response.FeaturePilotFlags

            assert feature_pilot_flags is not None and len(feature_pilot_flags) > 0

            print(f"\nCustomerId: {customer_id}")
            print("FeaturePilotFlags:")
            print('; '.join(map(str, feature_pilot_flags)))

    except WebFault as ex:
        output_webfault_errors(ex)
    except Exception as ex:
        output_status_message(ex)

if __name__ == '__main__':
    print("Loading the web service client...")

    authorization_data = AuthorizationData(
        account_id=None,
        customer_id=None,
        developer_token=DEVELOPER_TOKEN,
        authentication=None,
    )

    authenticate(authorization_data)

    customer_service = ServiceClient(
        service='CustomerManagementService',
        version=13,
        authorization_data=authorization_data,
        environment=ENVIRONMENT,
    )

    main(authorization_data)
