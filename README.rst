===================
Bing Ads Python SDK
===================

Local troubleshooting examples
------------------------------

This repository is a sanitized working snapshot of
`BingAds/BingAds-Python-SDK <https://github.com/BingAds/BingAds-Python-SDK>`_,
with local Python examples and a
`REST troubleshooting guide <API错误排查.md>`_. The upstream license is
preserved in LICENSE. This is not the official SDK repository.

The examples under ``examples/v13`` include:

* ``search_user_accounts_official_exception_test.py``: official
  ``ApiException`` status and response-body output.
* ``search_user_accounts_test.py``: extended REST HTTP diagnostics,
  response headers, exception stack traces, and nonzero exit on failure.
* ``search_user_accounts_suds_logging_test.py``: comparison of SOAP-oriented
  exception handling and logging against a REST call.
* ``enhanced_offline_conversion_upload/upload_excel_conversions.py``:
  local Excel conversion validation and optional upload.

Before running, copy ``examples/v13/auth_helper.example.py`` to
``examples/v13/auth_helper.py`` locally, then configure ``YOUR_DEVELOPER_TOKEN``
and ``YOUR_CLIENT_ID`` in that local file. The real ``auth_helper.py`` and
refresh-token files are ignored by Git and must never be committed.
Replace ``YOUR_USER_ID`` in the three account-search scripts with your own
user ID. Run the examples from
``examples/v13`` using an environment with the project dependencies installed.
The suds comparison also requires ``suds-community``.

The search examples deliberately duplicate the UserId predicate to reproduce
HTTP 400 / business error 3030. The extended diagnostic script re-raises the
exception after printing it. Its HTTP logs can expose credentials; do not
publish raw output. No refresh tokens, customer workbooks, or local logs are
included. The Excel uploader requires your own workbook and can write
conversion data when explicitly invoked with ``--upload``.

.. image:: https://img.shields.io/pypi/v/bingads.svg
        :target: https://pypi.python.org/pypi/bingads


The Bing Ads Python Software Development Kit (SDK) Version 13.0 enhances the experience of developing Bing Ads applications
with the Python programming language.
The SDK includes classes for all Bing Ads REST API web services and abstracts the low level details of authentication with OAuth.
You can also read and write bulk files with the SDK BulkFileReader and BulkFileWriter,
and use the high level BulkServiceManager interface to abstract and execute operations in the low level Bulk API.
For more information, see `Bing Ads Client Libraries <https://docs.microsoft.com/en-us/bingads/guides/client-libraries>`_.

Getting Started
---------------

To get started developing Bing Ads applications with Python, 
install the SDK and either start with the `examples <https://github.com/BingAds/BingAds-Python-SDK/tree/master/examples>`_ or follow one of the application walkthroughs.
For more information, see `Getting Started Using Python with Bing Ads Services <https://docs.microsoft.com/en-us/bingads/guides/get-started-python>`_.

External Dependencies
---------------------
    
    - `requests <http://pypi.python.org/pypi/requests>`_
    
    - `enum34 <http://pypi.python.org/pypi/enum34>`_