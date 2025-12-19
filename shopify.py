import requests
import json
from datetime import datetime
from dateutil.relativedelta import relativedelta
import pandas as pd
import csv
from shopify_cred import credentials
import os
from pathlib import Path
from urllib.parse import urlencode

def filter_by_time(data: json, start_date: datetime, end_date: datetime):
    filtered_orders = []

    for order in data["orders"]:
        ordered_at_str = order.get("created_at")
        order_status = order.get("cancel_reason", "")

        print("Ordered At:", ordered_at_str)
        print("Order Status:", order_status)

        if ordered_at_str:
            ordered_at = datetime.fromisoformat(ordered_at_str.replace("Z", "+00:00"))
            ordered_at_naive = ordered_at.replace(tzinfo=None)
            if start_date <= ordered_at_naive <= end_date:
                filtered_orders.append(order)

    return filtered_orders

def extract_useful_info(data: json):
    return {
        "order_number": data.get("id"),
        "customer_name": data.get("name"),
        "created_at": data.get("created_at"),
        "items": [
            {
                "sku": item.get("sku"),
                "quantity": item.get("quantity"),
            }
            for item in data.get("line_items", [])
        ]
    }

def shopify():
    store_name = credentials["shopify_store_name"]
    access_token = credentials["shopify_access_token"]

    # 設置時間範圍
    start_date = datetime.now() - relativedelta(months=3)
    end_date = datetime.now()

    base_orders_endpoint = f"https://{store_name}.myshopify.com/admin/api/2023-10/orders.json"
    
    headers = {
        "X-Shopify-Access-Token": access_token,
        "Content-Type": "application/json"
    }

    all_orders = []
    next_page_url = None

    # 實現分頁邏輯，使用 Link header 解析
    while True:
        # 構建請求 URL
        if next_page_url:
            # 如果有下一頁，使用完整的下一頁 URL
            current_url = next_page_url
        else:
            # 第一頁請求，使用基礎 URL 並添加初始參數
            params = {
                'status': 'any',
                'limit': 250
            }
            current_url = f"{base_orders_endpoint}?{urlencode(params)}"
        
        print(f"Fetching: {current_url}")
        
        response = requests.get(current_url, headers=headers)
        
        if response.status_code != 200:
            print(f"Error: {response.status_code} - {response.text}")
            break
        
        data = response.json()
        
        # 檢查是否有訂單數據
        if 'orders' not in data or not data['orders']:
            print("No more orders found")
            break
        
        print(f"Retrieved {len(data['orders'])} orders")
        
        # 在客戶端進行時間過濾
        filtered_orders = filter_by_time(data, start_date, end_date)
        all_orders.extend(filtered_orders)
        
        # 檢查是否有下一頁 - 通過 Link header
        link_header = response.headers.get('Link', '')
        next_page_url = None
        
        # 解析 Link header 來找到下一頁
        if 'rel="next"' in link_header:
            links = link_header.split(',')
            for link in links:
                if 'rel="next"' in link:
                    # 提取 URL 部分
                    next_page_url = link.split(';')[0].strip('<> ')
                    break
        
        # 如果沒有下一頁，退出循環
        if not next_page_url:
            print("No more pages")
            break
    
    # 保留這段特定的數據處理邏輯
    result = []
    for order in all_orders:
        result.append(extract_useful_info(order))
    
    print(f"Total filtered orders retrieved: {len(result)}")
    return result

if __name__ == "__main__":
    result = shopify()

    if result:
        with open('shopify.csv', 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['order_number', 'customer_name', 'created_at', 'items']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            
            for row in result:
                # 將 items 列表轉換為 JSON 字符串
                row_copy = row.copy()
                row_copy['items'] = json.dumps(row_copy['items'])
                writer.writerow(row_copy)
        
        print(f"CSV file created with {len(result)} records")
    else:
        print("No records found")