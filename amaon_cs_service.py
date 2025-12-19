import requests
import time
import gzip
import io
import csv
import urllib.parse
from datetime import datetime, timedelta
from amazon_cred import credentials
import concurrent.futures
import json

# Global log file
LOG_FILE = "amazon_api_logs.txt"

def log_to_file(content):
    """Append content to the single log file"""
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(content + "\n")

def clear_log_file():
    """Clear the log file at the start"""
    with open(LOG_FILE, 'w', encoding='utf-8') as f:
        f.write(f"=== AMAZON API LOGS - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===\n\n")

def amazon_order(log_message=None):
    start_time = time.time()
    clear_log_file()

    if log_message is None:
        log_message = print

    print(f"開始執行 Amazon 訂單數據獲取: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log_to_file(f"開始執行 Amazon 訂單數據獲取: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # Get access token
    token_start = time.time()
    token_request_data = {
        "grant_type": "refresh_token",
        "refresh_token": credentials["refresh_token"],
        "client_id": credentials["lwa_app_id"],
        "client_secret": credentials["lwa_client_secret"],
    }
    
    log_to_file("=== TOKEN REQUEST ===")
    log_to_file(f"URL: https://api.amazon.com/auth/o2/token")
    log_to_file(f"Method: POST")
    log_to_file(f"Headers: Content-Type: application/x-www-form-urlencoded")
    log_to_file(f"Body: {json.dumps(token_request_data, indent=2)}")
    
    token_response = requests.post(
        "https://api.amazon.com/auth/o2/token",
        data=token_request_data,
    )
    
    log_to_file("=== TOKEN RESPONSE ===")
    log_to_file(f"Status Code: {token_response.status_code}")
    log_to_file(f"Headers: {json.dumps(dict(token_response.headers), indent=2)}")
    log_to_file(f"Body: {token_response.text}")
    log_to_file("")  # Empty line for separation
    
    access_token = token_response.json()["access_token"]
    print(f"獲取 Access Token 用時: {time.time() - token_start:.2f} 秒")
    log_to_file(f"獲取 Access Token 用時: {time.time() - token_start:.2f} 秒")

    # API configuration
    endpoint = "https://sellingpartnerapi-na.amazon.com"
    marketplace_id = "ATVPDKIKX0DER"
    headers = {
        "x-amz-access-token": access_token,
        "Content-Type": "application/json"
    }

    # Request report generation
    report_start = time.time()
    report_url = f"{endpoint}/reports/2021-06-30/reports"
    report_data = {
        "reportType": "GET_MERCHANT_LISTINGS_ALL_DATA",
        "dataStartTime": "2025-06-10T20:11:24.000Z",
        "marketplaceIds": [marketplace_id],
    }

    log_to_file("=== REPORT REQUEST ===")
    log_to_file(f"URL: {report_url}")
    log_to_file(f"Method: POST")
    log_to_file(f"Headers: {json.dumps(headers, indent=2)}")
    log_to_file(f"Body: {json.dumps(report_data, indent=2)}")
    
    response = requests.post(report_url, headers=headers, json=report_data)
    
    log_to_file("=== REPORT RESPONSE ===")
    log_to_file(f"Status Code: {response.status_code}")
    log_to_file(f"Headers: {json.dumps(dict(response.headers), indent=2)}")
    log_to_file(f"Body: {response.text}")
    log_to_file("")

    if response.status_code != 202:
        log_message(f"Error creating report: {response.status_code} - {response.text}")
        exit()

    report_id = response.json()["reportId"]
    log_message(f"報告請求用時: {time.time() - report_start:.2f} 秒")
    log_to_file(f"報告請求用時: {time.time() - report_start:.2f} 秒")
    print(f"Report requested. Report ID: {report_id}")

    # Check report status and wait for processing
    max_attempts = 30
    attempt = 0
    report_status = ""
    
    status_start = time.time()
    while attempt < max_attempts:
        status_url = f"{endpoint}/reports/2021-06-30/reports/{report_id}"
        
        log_to_file(f"=== STATUS REQUEST {attempt} ===")
        log_to_file(f"URL: {status_url}")
        log_to_file(f"Method: GET")
        log_to_file(f"Headers: {json.dumps(headers, indent=2)}")
        log_to_file(f"Body: None")
        
        status_response = requests.get(status_url, headers=headers)
        
        log_to_file(f"=== STATUS RESPONSE {attempt} ===")
        log_to_file(f"Status Code: {status_response.status_code}")
        log_to_file(f"Headers: {json.dumps(dict(status_response.headers), indent=2)}")
        log_to_file(f"Body: {status_response.text}")
        log_to_file("")
        
        if status_response.status_code != 200:
            log_message(f"Error checking report status: {status_response.status_code}")
            time.sleep(30)
            attempt += 1
            continue
            
        report_status = status_response.json()["processingStatus"]
        
        if report_status == "DONE":
            log_message(f"報告處理完成，用時: {time.time() - status_start:.2f} 秒")
            log_to_file(f"報告處理完成，用時: {time.time() - status_start:.2f} 秒")
            break
        elif report_status in ["CANCELLED", "FATAL"]:
            print(f"Report processing failed with status: {report_status}")
            exit()
        else:
            print(f"Report status: {report_status}. Waiting... ({attempt + 1}/{max_attempts})")
            time.sleep(30)
            attempt += 1

    if report_status != "DONE":
        log_message("Report processing timed out.")
        exit()

    # Get report document ID
    report_document_id = status_response.json()["reportDocumentId"]
    print(f"Report document ID: {report_document_id}")

    # Download the report
    download_start = time.time()
    document_url = f"{endpoint}/reports/2021-06-30/documents/{report_document_id}"
    
    log_to_file("=== DOCUMENT REQUEST ===")
    log_to_file(f"URL: {document_url}")
    log_to_file(f"Method: GET")
    log_to_file(f"Headers: {json.dumps(headers, indent=2)}")
    log_to_file(f"Body: None")
    
    document_response = requests.get(document_url, headers=headers)

    log_to_file("=== DOCUMENT RESPONSE ===")
    log_to_file(f"Status Code: {document_response.status_code}")
    log_to_file(f"Headers: {json.dumps(dict(document_response.headers), indent=2)}")
    log_to_file(f"Body: {document_response.text}")
    log_to_file("")

    if document_response.status_code != 200:
        log_message(f"Error downloading report: {document_response.status_code}")
        exit()

    document_info = document_response.json()
    report_download_url = document_info["url"]

    # Download the actual report content
    log_to_file("=== CONTENT REQUEST ===")
    log_to_file(f"URL: {report_download_url}")
    log_to_file(f"Method: GET")
    log_to_file(f"Headers: None")
    log_to_file(f"Body: None")
    
    report_content_response = requests.get(report_download_url)
    
    log_to_file("=== CONTENT RESPONSE ===")
    log_to_file(f"Status Code: {report_content_response.status_code}")
    log_to_file(f"Headers: {json.dumps(dict(report_content_response.headers), indent=2)}")
    log_to_file(f"Body: Binary content (size: {len(report_content_response.content)} bytes)")
    log_to_file("")

    # Check if the report is compressed
    if 'compressionAlgorithm' in document_info and document_info['compressionAlgorithm'] == 'GZIP':
        compressed_file = io.BytesIO(report_content_response.content)
        with gzip.GzipFile(fileobj=compressed_file) as f:
            report_data = f.read().decode('utf-8')
    else:
        report_data = report_content_response.text

    log_message(f"報告下載和解壓用時: {time.time() - download_start:.2f} 秒")
    log_to_file(f"報告下載和解壓用時: {time.time() - download_start:.2f} 秒")

    # Split the report into lines
    lines = report_data.split('\n')

    # Find the column index for seller-sku
    headers = lines[0].split('\t')
    try:
        sku_index = headers.index('seller-sku')
    except ValueError:
        # Try alternative column names
        try:
            sku_index = headers.index('sku')
        except ValueError:
            try:
                sku_index = headers.index('Seller SKU')
            except ValueError:
                log_message("Could not find seller-sku column in headers:", headers)
                exit()

    # Extract all SKUs from the report
    skus = []
    for line in lines[1:]:  # Skip header row
        if line.strip():  # Skip empty lines
            columns = line.split('\t')
            if len(columns) > sku_index:
                skus.append(columns[sku_index])

    log_message(f"Found {len(skus)} SKUs to process")
    log_to_file(f"Found {len(skus)} SKUs to process")

    # 計算日期範圍
    today = datetime.now()
    last_30_days_start = (today - timedelta(days=30)).strftime('%Y-%m-%d')
    last_30_days_end = today.strftime('%Y-%m-%d')
    last_4_months_start = (today - timedelta(days=120)).strftime('%Y-%m-%d')
    last_4_months_end = today.strftime('%Y-%m-%d')

    log_message(f"時間範圍 - 最近30天: {last_30_days_start} 到 {last_30_days_end}")
    log_message(f"時間範圍 - 最近4個月: {last_4_months_start} 到 {last_4_months_end}")
    log_to_file(f"時間範圍 - 最近30天: {last_30_days_start} 到 {last_30_days_end}")
    log_to_file(f"時間範圍 - 最近4個月: {last_4_months_start} 到 {last_4_months_end}")

    # 使用多線程處理 SKU
    api_start = time.time()
    log_message(f"開始處理 {len(skus)} 個 SKU 的 API 請求...")
    log_to_file(f"開始處理 {len(skus)} 個 SKU 的 API 請求...")

    # 使用線程池並行處理
    sku_order_items = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
        futures = {executor.submit(process_sku, sku): sku for sku in skus}
        
        completed = 0
        for future in concurrent.futures.as_completed(futures):
            try:
                result = future.result()
                sku_order_items.append(result)
                completed += 1
                if completed % 10 == 0:
                    elapsed = time.time() - api_start
                    speed = completed / elapsed if elapsed > 0 else 0
                    remaining = (len(skus) - completed) / speed if speed > 0 else 0
                    progress_msg = f"進度: {completed}/{len(skus)} | 速度: {speed:.2f} SKU/秒 | 預計剩餘時間: {remaining:.0f} 秒"
                    log_message(progress_msg)
                    log_to_file(progress_msg)
                    
            except Exception as e:
                error_msg = f"處理 SKU 時出錯: {str(e)}"
                log_message(error_msg)
                log_to_file(error_msg)

    api_time = time.time() - api_start
    log_message(f"API 請求處理完成，用時: {api_time:.2f} 秒")
    log_message(f"平均速度: {len(skus) / api_time:.2f} SKU/秒")
    log_to_file(f"API 請求處理完成，用時: {api_time:.2f} 秒")
    log_to_file(f"平均速度: {len(skus) / api_time:.2f} SKU/秒")

    total_time = time.time() - start_time
    log_message(f"總執行時間: {total_time:.2f} 秒")
    log_to_file(f"總執行時間: {total_time:.2f} 秒")
    
    return sku_order_items

def process_sku(sku):
    encoded_sku = urllib.parse.quote(sku, safe='')

    # Get access token for this SKU
    token_request_data = {
        "grant_type": "refresh_token",
        "refresh_token": credentials["refresh_token"],
        "client_id": credentials["lwa_app_id"],
        "client_secret": credentials["lwa_client_secret"],
    }
    
    log_to_file(f"=== SKU {sku} - TOKEN REQUEST ===")
    log_to_file(f"URL: https://api.amazon.com/auth/o2/token")
    log_to_file(f"Method: POST")
    log_to_file(f"Headers: Content-Type: application/x-www-form-urlencoded")
    log_to_file(f"Body: {json.dumps(token_request_data, indent=2)}")
    
    token_response = requests.post(
        "https://api.amazon.com/auth/o2/token",
        data=token_request_data,
    )
    
    log_to_file(f"=== SKU {sku} - TOKEN RESPONSE ===")
    log_to_file(f"Status Code: {token_response.status_code}")
    log_to_file(f"Headers: {json.dumps(dict(token_response.headers), indent=2)}")
    log_to_file(f"Body: {token_response.text}")
    
    access_token = token_response.json()["access_token"]

    today = datetime.now()
    last_30_days_start = (today - timedelta(days=30)).strftime('%Y-%m-%d')
    last_30_days_end = today.strftime('%Y-%m-%d')
    last_4_months_start = (today - timedelta(days=120)).strftime('%Y-%m-%d')
    last_4_months_end = today.strftime('%Y-%m-%d')

    # 構建兩個時間段的 URL
    url_30_days = f"https://sellingpartnerapi-na.amazon.com/sales/v1/orderMetrics?marketplaceIds=ATVPDKIKX0DER&interval={last_30_days_start}T00%3A00%3A00-08%3A00--{last_30_days_end}T00%3A00%3A00-08%3A00&granularity=Total&buyerType=All&firstDayOfWeek=Monday&sku={encoded_sku}"
    url_4_months = f"https://sellingpartnerapi-na.amazon.com/sales/v1/orderMetrics?marketplaceIds=ATVPDKIKX0DER&interval={last_4_months_start}T00%3A00%3A00-08%3A00--{last_4_months_end}T00%3A00%3A00-12%3A00&granularity=Month&buyerType=All&firstDayOfWeek=Monday&sku={encoded_sku}"
    
    headers = {"x-amz-access-token": access_token}

    order_item_count_30_days = 0
    order_item_count_4_months = 0

    try:
        # 並行請求兩個時間段的數據
        with requests.Session() as session:
            # 30 days request
            log_to_file(f"=== SKU {sku} - 30 DAYS REQUEST ===")
            log_to_file(f"URL: {url_30_days}")
            log_to_file(f"Method: GET")
            log_to_file(f"Headers: {json.dumps(headers, indent=2)}")
            log_to_file(f"Body: None")
            
            response_30 = session.get(url_30_days, headers=headers, timeout=10)
            
            log_to_file(f"=== SKU {sku} - 30 DAYS RESPONSE ===")
            log_to_file(f"Status Code: {response_30.status_code}")
            log_to_file(f"Response Time: {response_30.elapsed.total_seconds()} seconds")
            log_to_file(f"Headers: {json.dumps(dict(response_30.headers), indent=2)}")
            log_to_file(f"Body: {response_30.text}")
            
            # 4 months request
            log_to_file(f"=== SKU {sku} - 4 MONTHS REQUEST ===")
            log_to_file(f"URL: {url_4_months}")
            log_to_file(f"Method: GET")
            log_to_file(f"Headers: {json.dumps(headers, indent=2)}")
            log_to_file(f"Body: None")
            
            response_4 = session.get(url_4_months, headers=headers, timeout=10)
            
            log_to_file(f"=== SKU {sku} - 4 MONTHS RESPONSE ===")
            log_to_file(f"Status Code: {response_4.status_code}")
            log_to_file(f"Response Time: {response_4.elapsed.total_seconds()} seconds")
            log_to_file(f"Headers: {json.dumps(dict(response_4.headers), indent=2)}")
            log_to_file(f"Body: {response_4.text}")
            log_to_file("")  # Empty line for separation
            
            if response_30.status_code == 200:
                data_30 = response_30.json()
                if 'payload' in data_30 and len(data_30['payload']) > 0:
                    order_item_count_30_days = data_30['payload'][0]['orderItemCount']
            
            if response_4.status_code == 200:
                data_4 = response_4.json()
                if 'payload' in data_4 and len(data_4['payload']) > 0:
                    order_item_count_4_months = data_4['payload'][0]['orderItemCount']
                    
    except Exception as e:
        error_msg = f"處理 SKU {sku} 時出錯: {str(e)}"
        print(error_msg)
        log_to_file(f"=== SKU {sku} - ERROR ===")
        log_to_file(f"Error: {str(e)}")
        log_to_file(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        log_to_file("")

    # 計算4個月平均值
    monthly_avg_4_months = round(order_item_count_4_months / 4, 2) if order_item_count_4_months > 0 else 0
    
    return {
        'sku': sku,
        'amazon_orderItemCount_30_days': order_item_count_30_days,
        'amazon_orderItemCount_4_months_total': order_item_count_4_months,
        'amazon_monthly_avg_4_months': monthly_avg_4_months
    }

if __name__ == "__main__":
    result = amazon_order()

    if result:
        with open('amazon_order.csv', 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['sku', 'amazon_orderItemCount_30_days', 'amazon_orderItemCount_4_months_total', 'amazon_monthly_avg_4_months']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(result)
        
        print(f"CSV 文件創建完成，包含 {len(result)} 條記錄")
        print(f"所有日誌已保存到: {LOG_FILE}")