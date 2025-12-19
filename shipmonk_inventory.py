import requests
from datetime import datetime
import json
from shipmonk_cred import credentials
import csv


def filter_by_time(data :json, start_date: datetime, end_date: datetime):
    filtered_orders = []

    for order in data["data"]["orders"]:
        ordered_at_str = order.get("ordered_at")
        order_status = order.get("order_status", "")

        print("Ordered At:", ordered_at_str)
        print("Order Status:", order_status)

        if ordered_at_str and order_status.strip().lower() != "cancelled":
            ordered_at = datetime.fromisoformat(ordered_at_str.replace("Z", "+00:00"))
            ordered_at_naive = ordered_at.replace(tzinfo=None)

            if start_date <= ordered_at_naive <= end_date:
                print("Included:", ordered_at_naive)
                filtered_orders.append(order)

    return filtered_orders


def extract_useful_info(product: dict):
    inventory = product.get("inventory", {})
    locations = inventory.get("locations", [])
    
    return {
        "sku": product.get("sku"),
        "created_at": product.get("created_at"),
        "updated_at": product.get("updated_at"),
        "total_available": inventory.get("quantity_total_available"),
 
        
    }
   
def shipmonk_inventory():

    url = "https://api.shipmonk.com/v1/products?page=1&pageSize=100&sortBy=id&sortOrder=ASC"



    headers = {
        "accept": "application/json",
        "Api-Key": credentials["shipmonk_inventory_api"]
    }



    response = requests.get(url, headers=headers)
    responsee_json = json.loads(response.text)

    # Filter orders within the date range



    result_list = []
    for product in responsee_json["data"]:
        result_list.append(extract_useful_info(product))



    return result_list
# if __name__ == "__main__":

#     result = shopmonk_inventory()

#     if result:
#         with open('shopmonk_inventory.csv', 'w', newline='', encoding='utf-8') as csvfile:
#             fieldnames = ['sku', 'created_at', 'updated_at', 'total_available']
#             writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
#             writer.writeheader()
#             writer.writerows(result)
        
#         print(f"CSV file created with {len(result)} records")
