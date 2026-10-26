from auth_helper import *
from openapi_client.models.customer import *
from openapi_client.exceptions import ApiException
import logging
import http.client as http_client
import traceback

logging.basicConfig(level=logging.INFO)
logging.getLogger("urllib3").setLevel(logging.DEBUG)
http_client.HTTPConnection.debuglevel = 1

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
        
        #search_request = SearchAccountsRequest(
        #    Predicates=[predicate],
        #    PageInfo=paging
        #)
        #
        #search_response = customer_service.search_accounts(search_request)
        
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
        
    except ApiException as ex:
        print(f"HTTP status {ex.status}: {ex.body}")
        print(f"Response headers: {ex.headers}")
        traceback.print_exc()
        raise
    except Exception:
        traceback.print_exc()
        raise

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