## REST API 错误排查（Python SDK）

1. 请优先使用官方示例中的 **`ApiException` 异常处理方式**，输出 HTTP 状态码和原始错误响应正文，根据错误提示检查参数和权限。

   ```python
   from openapi_client.exceptions import ApiException

   try:
       # 在这里调用 API
       ...
   except ApiException as ex:
       print(f"HTTP status {ex.status}: {ex.body}")
   ```

   该方式只输出状态码和错误响应正文，不会自动打印请求、响应头或异常堆栈。即使只有一行，也可能已包含业务错误码、错误原因和 TrackingId。例如，本次测试输出了 HTTP `400`、业务错误码 `3030` 和 `Predicate passed in search is invalid.`。

   **注意：**仅打印 `ex` 或 traceback，可能出现 `HTTP response body:` 为空的情况，并不代表服务端没有返回错误详情；请显式打印 `ex.body`。网络连接等异常不一定属于 `ApiException`，也不一定有 HTTP 响应或 TrackingId。

2. 如果无法根据**异常信息**解决问题，请在 API 调用前添加以下配置，临时打印请求和响应的调试信息，结合 API 文档检查实际请求是否符合参数要求。

   ```python
   import logging
   import http.client as http_client

   logging.basicConfig(level=logging.INFO)
   logging.getLogger("urllib3").setLevel(logging.DEBUG)

   http_client.HTTPConnection.debuglevel = 1
   ```

   该配置可帮助查看请求路径、请求头、请求参数、响应状态和响应头；**不会保证打印响应 body**，请同时保留第 1 步中的异常输出。

   如需进一步定位调用链或网络错误，可将第 1 步中的异常处理扩展为以下形式，额外输出响应头和完整异常堆栈。通用异常打印后重新抛出，避免将失败当作成功处理。

   ```python
   from openapi_client.exceptions import ApiException
   import traceback

   try:
       # 在这里调用 API
       ...
   except ApiException as ex:
       print(f"HTTP status {ex.status}: {ex.body}")
       print(f"Response headers: {ex.headers}")
       traceback.print_exc()
       raise
   except Exception:
       traceback.print_exc()
       raise
   ```

   本节日志配置适用于当前基于 `urllib3` 的 REST SDK。`suds.client`、`suds.transport.http` 日志配置和 `WebFault` 异常处理针对 SOAP，不应替代上述 REST 排错方式。

   **安全提醒：**日志可能包含 access token、refresh token、Developer Token、Cookie 等敏感信息，分享前必须脱敏。排查结束后请将 `http_client.HTTPConnection.debuglevel` 恢复为 `0`，并将 `urllib3` 日志级别恢复为 `logging.WARNING`。仅调整 logging 级别不会关闭底层 HTTP 调试输出。

3. 如果仍然无法解决，请将**完整错误信息、发生时间及其时区、SDK 版本、脱敏后的相关请求和日志**包含在工单中。如果错误响应或响应头中有 **TrackingId**，请一并提供，以便进一步定位。
