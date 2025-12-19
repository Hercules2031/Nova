import requests
import time
import csv
import urllib.parse
from datetime import datetime, timedelta
from amazon_cred import credentials

def amazon_order_specific_sku(log_message=None):
    start_time = time.time()

    if log_message is None:
        log_message = print

    print(f"開始執行 Amazon 訂單數據獲取: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # 目标 SKU
    target_sku = "DT-009"
    print(f"目標 SKU: {target_sku}")
    
    # 市场列表
    marketplaces = [
        {"id": "A2EUQ1WTGCTBG2", "name": "Canada"},
        {"id": "ATVPDKIKX0DER", "name": "United States of America"},
        {"id": "A1AM78C64UM0Y8", "name": "Mexico"},
        {"id": "A1RKKUPIHCS9HS", "name": "Spain"},
        {"id": "A1F83G8C2ARO7P", "name": "United Kingdom"},
        {"id": "A13V1IB3VIYZZH", "name": "France"},
        {"id": "A1PA6795UKMFR9", "name": "Germany"},
        {"id": "APJ6JRA9NG5V4", "name": "Italy"},
        {"id": "A39IBJ37TRP1C6", "name": "Australia"},
        {"id": "A1VC38T7YXB528", "name": "Japan"}
    ]

    log_message(f"將在 {len(marketplaces)} 個市場中查詢 SKU: {target_sku}")

    # 顺序处理每个市场
    api_start = time.time()
    log_message(f"開始順序處理 {len(marketplaces)} 個市場的 API 請求...")
    
    market_sales_data = []
    
    for i, marketplace in enumerate(marketplaces, 1):
        log_message(f"\n--- 處理第 {i}/{len(marketplaces)} 個市場: {marketplace['name']} ---")
        
        try:
            result = process_marketplace(target_sku, marketplace)
            market_sales_data.append(result)
            
            # 顯示進度
            elapsed = time.time() - api_start
            speed = i / elapsed if elapsed > 0 else 0
            remaining = (len(marketplaces) - i) / speed if speed > 0 else 0
            log_message(f"進度: {i}/{len(marketplaces)} | 速度: {speed:.2f} 市場/秒 | 預計剩餘時間: {remaining:.0f} 秒")
            
            # 每個市場處理完成後等待1秒（最後一個市場不用等）
            if i < len(marketplaces):
                log_message("等待1秒...")
                time.sleep(1)
                
        except Exception as e:
            log_message(f"處理市場 {marketplace['name']} 時出錯: {str(e)}")
            # 錯誤後也等待1秒
            time.sleep(1)

    api_time = time.time() - api_start
    log_message(f"API 請求處理完成，用時: {api_time:.2f} 秒")

    # 計算總銷售量
    total_30_days = sum(item['order_item_count_30_days'] for item in market_sales_data)
    total_4_months = sum(item['order_item_count_4_months_total'] for item in market_sales_data)
    total_monthly_avg = sum(item['monthly_avg_4_months'] for item in market_sales_data)

    # 添加總計行
    market_sales_data.append({
        'marketplace_id': 'TOTAL',
        'marketplace_name': 'ALL MARKETS',
        'sku': target_sku,
        'order_item_count_30_days': total_30_days,
        'order_item_count_4_months_total': total_4_months,
        'monthly_avg_4_months': round(total_monthly_avg, 2)
    })

    total_time = time.time() - start_time
    log_message(f"總執行時間: {total_time:.2f} 秒")
    log_message(f"總銷售量 - 最近30天: {total_30_days}")
    log_message(f"總銷售量 - 最近4個月: {total_4_months}")
    log_message(f"總月平均值: {total_monthly_avg:.2f}")
    
    return market_sales_data

def process_marketplace(sku, marketplace):
    """處理單個市場的銷售數據"""
    encoded_sku = urllib.parse.quote(sku, safe='')

    # 獲取 access token
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

    # 計算日期範圍
    today = datetime.now()
    last_30_days_start = (today - timedelta(days=30)).strftime('%Y-%m-%d')
    last_30_days_end = today.strftime('%Y-%m-%d')
    last_4_months_start = (today - timedelta(days=120)).strftime('%Y-%m-%d')
    last_4_months_end = today.strftime('%Y-%m-%d')

    # 構建兩個時間段的 URL
    url_30_days = f"https://sellingpartnerapi-na.amazon.com/sales/v1/orderMetrics?marketplaceIds={marketplace['id']}&interval={last_30_days_start}T00%3A00%3A00-08%3A00--{last_30_days_end}T00%3A00%3A00-08%3A00&granularity=Total&buyerType=All&firstDayOfWeek=Monday&sku={encoded_sku}"
    url_4_months = f"https://sellingpartnerapi-na.amazon.com/sales/v1/orderMetrics?marketplaceIds={marketplace['id']}&interval={last_4_months_start}T00%3A00%3A00-08%3A00--{last_4_months_end}T00%3A00%3A00-12%3A00&granularity=Month&buyerType=All&firstDayOfWeek=Monday&sku={encoded_sku}"

    headers = {"x-amz-access-token": access_token}

    order_item_count_30_days = 0
    order_item_count_4_months = 0

    try:
        # 順序請求兩個時間段的數據
        print(f"正在請求 {marketplace['name']} 的30天數據...")
        response_30 = requests.get(url_30_days, headers=headers, timeout=15)
        time.sleep(2)  # 兩個請求之間等待0.5秒
        
        print(f"正在請求 {marketplace['name']} 的4個月數據...")
        response_4 = requests.get(url_4_months, headers=headers, timeout=15)
        
        # 處理30天數據
        if response_30.status_code == 200:
            data_30 = response_30.json()
            if 'payload' in data_30 and len(data_30['payload']) > 0:
                order_item_count_30_days = data_30['payload'][0].get('orderItemCount', 0)
                print(f"✅ {marketplace['name']} 30天數據: {order_item_count_30_days}")
            else:
                print(f"ℹ️ {marketplace['name']} 沒有30天銷售數據")
        elif response_30.status_code in [403, 429]:
            print(f"⚠️ {marketplace['name']} 30天數據請求被限制: {response_30.status_code}")
        else:
            print(f"❌ {marketplace['name']} 30天數據請求失敗: {response_30.status_code}")
        
        # 處理4個月數據
        if response_4.status_code == 200:
            data_4 = response_4.json()
            if 'payload' in data_4 and len(data_4['payload']) > 0:
                # 計算4個月的總銷售量
                total_4_months = 0
                for month_data in data_4['payload']:
                    total_4_months += month_data.get('orderItemCount', 0)
                order_item_count_4_months = total_4_months
                print(f"✅ {marketplace['name']} 4個月數據: {order_item_count_4_months}")
            else:
                print(f"ℹ️ {marketplace['name']} 沒有4個月銷售數據")
        elif response_4.status_code in [403, 429]:
            print(f"⚠️ {marketplace['name']} 4個月數據請求被限制: {response_4.status_code}")
        else:
            print(f"❌ {marketplace['name']} 4個月數據請求失敗: {response_4.status_code}")
                
    except requests.exceptions.Timeout:
        print(f"⏰ {marketplace['name']} 請求超時")
    except Exception as e:
        print(f"❌ 處理市場 {marketplace['name']} 時出錯: {str(e)}")

    # 計算4個月平均值
    monthly_avg_4_months = round(order_item_count_4_months / 4, 2) if order_item_count_4_months > 0 else 0
    
    return {
        'marketplace_id': marketplace['id'],
        'marketplace_name': marketplace['name'],
        'sku': sku,
        'order_item_count_30_days': order_item_count_30_days,
        'order_item_count_4_months_total': order_item_count_4_months,
        'monthly_avg_4_months': monthly_avg_4_months
    }

if __name__ == "__main__":
    result = amazon_order_specific_sku()

    if result:
        # 輸出 CSV 文件
        filename = f'amazon_sales_DT009_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv'
        with open(filename, 'w', newline='', encoding='utf-8-sig') as csvfile:
            fieldnames = [
                'marketplace_id', 
                'marketplace_name', 
                'sku', 
                'order_item_count_30_days', 
                'order_item_count_4_months_total', 
                'monthly_avg_4_months'
            ]
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(result)
        
        print(f"\nCSV 文件創建完成: {filename}")
        print(f"包含 {len(result)} 條記錄 (包含總計行)")
        
        # 顯示總計信息
        total_row = result[-1]  # 最後一行是總計
        print(f"\n=== 銷售總計 ===")
        print(f"所有市場最近30天總銷售量: {total_row['order_item_count_30_days']}")
        print(f"所有市場最近4個月總銷售量: {total_row['order_item_count_4_months_total']}")
        print(f"所有市場月平均銷售量: {total_row['monthly_avg_4_months']}")