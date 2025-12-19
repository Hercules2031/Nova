import requests
import json
from datetime import datetime
from dateutil.relativedelta import relativedelta
import pandas as pd
import csv
from shopify_cred import credentials
import os
from pathlib import Path

def filter_by_time(data :json, start_date: datetime, end_date: datetime):
    filtered_orders = []

    for order in data["orders"]:
        ordered_at_str = order.get("created_at")
        order_status = order.get("cancel_reason", "")

        print("Ordered At:", ordered_at_str)
        print("Order Status:", order_status)

        if ordered_at_str :
            ordered_at = datetime.fromisoformat(ordered_at_str.replace("Z", "+00:00"))
            ordered_at_naive = ordered_at.replace(tzinfo=None)
            if start_date <= ordered_at_naive <= end_date:
                filtered_orders.append(order)

    return filtered_orders

def extract_useful_info(data :json):
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

    # Set up API endpoints
    date_3_months_ago = datetime.now() - relativedelta(months=3)
    start_date = datetime.now() - relativedelta(months=3)
    end_date = datetime.now()

    base_orders_endpoint = f"https://{store_name}.myshopify.com/admin/api/2023-10/orders.json"
    
    headers = {
        "X-Shopify-Access-Token": access_token,
        "Content-Type": "application/json"
    }

    all_orders = []
    next_page_url = None

    # 實現分頁邏輯，類似你的 JavaScript 範例
    while True:
        # 構建請求 URL
        if next_page_url:
            # 如果有下一頁，使用完整的下一頁 URL
            current_url = next_page_url
        else:
            # 第一頁請求，使用基礎 URL 並添加初始參數
            current_url = f"{base_orders_endpoint}?limit=250&status=any"
        
        print(f"Fetching: {current_url}")
        
        response = requests.get(current_url, headers=headers)
        
        if response.status_code != 200:
            print(f"Error: {response.status_code} - {response.text}")
            break
        
        data = response.json()
        
        # 過濾當前頁面的訂單
        filtered_orders = filter_by_time(data, start_date, end_date)
        all_orders.extend(filtered_orders)
        
        # 檢查是否有下一頁
        link_header = response.headers.get('Link', '')
        next_page_url = None
        
        # 解析 Link header 來找到下一頁
        if 'rel="next"' in link_header:
            # 提取下一頁的 URL
            links = link_header.split(',')
            for link in links:
                if 'rel="next"' in link:
                    # 提取 URL 部分
                    next_page_url = link.split(';')[0].strip('<> ')
                    break
        
        # 如果沒有下一頁，退出循環
        if not next_page_url:
            break
    
    # 提取有用信息
    result = []
    for order in all_orders:
        result.append(extract_useful_info(order))
    
    return result

if __name__ == "__main__":
    result = shopify()

    if result:
        with open('shopify.csv', 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['order_number', 'customer_name', 'created_at', 'items']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            
            # 處理 items 欄位，將其轉換為字符串以便寫入 CSV
            for row in result:
                # 將 items 列表轉換為 JSON 字符串
                row_copy = row.copy()
                row_copy['items'] = json.dumps(row_copy['items'])
                writer.writerow(row_copy)
        
        print(f"CSV file created with {len(result)} records")
    else:
        print("No records found")