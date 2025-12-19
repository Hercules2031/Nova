import requests
from datetime import datetime
import json
import csv
from shipmonk_cred import credentials
import os
from pathlib import Path
from dateutil.relativedelta import relativedelta

def filter_by_time(data :json, start_date: datetime, end_date: datetime):
    filtered_orders = []

    for order in data["data"]["orders"]:
        ordered_at_str = order.get("ordered_at")
        order_status = order.get("order_status", "")




        if ordered_at_str and order_status.strip().lower() != "cancelled":
            ordered_at = datetime.fromisoformat(ordered_at_str.replace("Z", "+00:00"))
            ordered_at_naive = ordered_at.replace(tzinfo=None)

            if start_date <= ordered_at_naive <= end_date:
                filtered_orders.append(order)

    return filtered_orders


def extract_useful_info(order :json):
        
        result= []
        

        for item in order:
        
            items = item["items"]
            sku = items[0]["sku"] 
            quantity = items[0]["quantity"] 
            result.append({
            "order_number": item["order_number"],      
            "customer_email": item["customer_email"],   
            "ordered_at": item["ordered_at"],          
            "sku": sku ,
            "quantity": quantity
        })
            

        return(result)
   

def shipmonk_order():
    url = "https://api.shipmonk.com/v1/integrations/orders-list?page=1&pageSize=100&sortOrder=DESC"

    
    headers = {
        "accept": "application/json",
        "Api-Key": credentials["shipmonk_order_api"]
    }

    # Define your date range
    end_date = datetime.now()

    # 计算四个月前
    start_date = end_date - relativedelta(months=4)



    response = requests.get(url, headers=headers)
    responsee_json = json.loads(response.text)

    # Filter orders within the date range
    filtered_orders = filter_by_time(responsee_json, start_date, end_date)


    list = extract_useful_info(filtered_orders)



    return list

if __name__ == "__main__":

    result = shipmonk_order()

    if result:
        with open('shipmonk_order.csv', 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['order_number', 'customer_email', 'ordered_at', 'sku','quantity']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(result)
        
        print(f"CSV file created with {len(result)} records")
