import requests
import urllib.parse
from datetime import datetime
import datetime
import json
from amazon_cred import credentials
import csv 
import os


def amazon_inventory():

    # Getting the LWA access token using the Seller Central App credentials. The token will be valid for 1 hour until it expires.
    token_response = requests.post(
        "https://api.amazon.com/auth/o2/token",
        data={
            "grant_type": "refresh_token",
            "refresh_token": credentials["refresh_token"],
            "client_id": credentials["lwa_app_id"],
            "client_secret": credentials["lwa_client_secret"],
        },
    )
    access_token = token_response.json()["access_token"]

    # North America SP API endpoint (from https://developer-docs.amazon.com/sp-api/docs/sp-api-endpoints)
    endpoint = "https://sellingpartnerapi-na.amazon.com"

    # US Amazon Marketplace ID (from https://developer-docs.amazon.com/sp-api/docs/marketplace-ids)
    marketplace_id = "ATVPDKIKX0DER"

    # Downloading orders (from https://developer-docs.amazon.com/sp-api/docs/orders-api-v0-reference#getorders)
    # the getOrders operation is a HTTP GET request with query parameters
    request_params = {
        "MarketplaceIds": marketplace_id,  # required parameter
        "CreatedAfter": (
            datetime.datetime.now() - datetime.timedelta(days=30)
        ).isoformat(),  # orders created since 30 days ago, the date needs to be in the ISO format
    }

    url = f"https://sellingpartnerapi-na.amazon.com/fba/inventory/v1/summaries?details=true&granularityType=Marketplace&granularityId=ATVPDKIKX0DER&marketplaceIds=ATVPDKIKX0DER&startDateTime=2025-06-15T10%3A30%3A00Z"



    headers = {"x-amz-access-token": access_token}

    response = requests.get(url, headers=headers)

    data = response.json()

    result = []

    for item in data["payload"]["inventorySummaries"]:
        # Convert ISO timestamp to normal datetime
        iso_time = item["lastUpdatedTime"]
        normal_time = iso_time.replace('T', ' ').replace('Z', '')
        
        result.append({
            "sellerSku": item["sellerSku"],
            "fulfillableQuantity": item["inventoryDetails"]["fulfillableQuantity"],
            "totalReservedQuantity": item["inventoryDetails"]["reservedQuantity"]["totalReservedQuantity"],
            "lastUpdatedTime": normal_time
        })

    return result

# if __name__ == "__main__":

#     result = amazon_inventory()

#     if result:
#         with open('amazon_inventory.csv', 'w', newline='', encoding='utf-8') as csvfile:
#             fieldnames = ['sellerSku', 'fulfillableQuantity', 'totalReservedQuantity', 'lastUpdatedTime']
#             writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
#             writer.writeheader()
#             writer.writerows(result)
        
#         print(f"CSV file created with {len(result)} records")