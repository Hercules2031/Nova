import requests
import time
import gzip
import io
import csv
import urllib.parse
from datetime import datetime, timedelta
from amazon_cred import credentials
import concurrent.futures

def amazon_order(log_message=None):
    start_time = time.time()

    if log_message is None:
        log_message = print

    print(f"開始執行 Amazon 訂單數據獲取: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # Get access token
    token_start = time.time()
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
    print(f"獲取 Access Token 用時: {time.time() - token_start:.2f} 秒")

    # API configuration
    endpoint = "https://sellingpartnerapi-na.amazon.com"
    marketplace_id = "ATVPDKIKX0DER"
    headers = {
        "x-amz-access-token": access_token,
        "Content-Type": "application/json"
    }

    # Define all report types to download
    report_types = [
        "GET_XML_ALL_ORDERS_DATA_BY_ORDER_DATE_GENERAL"
    ]

    # File name mapping for each report type
    file_names = {
        # Original reports
        "GET_FLAT_FILE_ACTIONABLE_ORDER_DATA_SHIPPING": "actionable_order_shipping.txt",
        "GET_ORDER_REPORT_DATA_INVOICING": "order_report_invoicing.txt",
        "GET_ORDER_REPORT_DATA_TAX": "order_report_tax.txt",
        "GET_ORDER_REPORT_DATA_SHIPPING": "order_report_shipping.txt",
        "GET_FLAT_FILE_ORDER_REPORT_DATA_INVOICING": "flat_file_order_invoicing.txt",
        "GET_FLAT_FILE_ORDER_REPORT_DATA_SHIPPING": "flat_file_order_shipping.txt",
        "GET_FLAT_FILE_ORDER_REPORT_DATA_TAX": "flat_file_order_tax.txt",
        # New order reports
        "GET_FLAT_FILE_ALL_ORDERS_DATA_BY_LAST_UPDATE_GENERAL": "all_orders_by_last_update_flat.txt",
        "GET_FLAT_FILE_ALL_ORDERS_DATA_BY_ORDER_DATE_GENERAL": "all_orders_by_order_date_flat.txt",
        "GET_FLAT_FILE_ARCHIVED_ORDERS_DATA_BY_ORDER_DATE": "archived_orders_by_order_date_flat.txt",
        "GET_XML_ALL_ORDERS_DATA_BY_LAST_UPDATE_GENERAL": "all_orders_by_last_update_xml.xml",
        "GET_XML_ALL_ORDERS_DATA_BY_ORDER_DATE_GENERAL": "all_orders_by_order_date_xml.xml",
        "GET_FLAT_FILE_PENDING_ORDERS_DATA": "pending_orders_flat.txt",
        "GET_PENDING_ORDERS_DATA": "pending_orders.txt",
        "GET_CONVERGED_FLAT_FILE_PENDING_ORDERS_DATA": "converged_pending_orders_flat.txt"
    }

    # Process each report type
    for report_type in report_types:
        log_message(f"\n開始處理報告類型: {report_type}")
        
        # Request report generation
        report_start = time.time()
        report_url = f"{endpoint}/reports/2021-06-30/reports"
        l = {
            "reportType": report_type,
            "dataStartTime": "2025-06-10T20:11:24.000Z",
            "marketplaceIds": [marketplace_id],
        }

        response = requests.post(report_url, headers=headers, json=report_data)
        if response.status_code != 202:
            log_message(f"Error creating report {report_type}: {response.status_code} - {response.text}")
            continue  # Continue with next report instead of exiting

        report_id = response.json()["reportId"]
        log_message(f"報告請求用時: {time.time() - report_start:.2f} 秒")
        print(f"Report requested. Report ID: {report_id}")

        # Check report status and wait for processing
        max_attempts = 30
        attempt = 0
        report_status = ""
        
        status_start = time.time()
        while attempt < max_attempts:
            status_url = f"{endpoint}/reports/2021-06-30/reports/{report_id}"
            status_response = requests.get(status_url, headers=headers)
            
            if status_response.status_code != 200:
                log_message(f"Error checking report status: {status_response.status_code}")
                time.sleep(30)
                attempt += 1
                continue
                
            report_status = status_response.json()["processingStatus"]
            
            if report_status == "DONE":
                log_message(f"報告處理完成，用時: {time.time() - status_start:.2f} 秒")
                break
            elif report_status in ["CANCELLED", "FATAL"]:
                print(f"Report processing failed with status: {report_status}")
                break
            else:
                print(f"Report status: {report_status}. Waiting... ({attempt + 1}/{max_attempts})")
                time.sleep(30)
                attempt += 1

        if report_status != "DONE":
            log_message(f"Report {report_type} processing timed out or failed.")
            continue  # Continue with next report

        # Get report document ID
        report_document_id = status_response.json()["reportDocumentId"]
        print(f"Report document ID: {report_document_id}")

        # Download the report
        download_start = time.time()
        document_url = f"{endpoint}/reports/2021-06-30/documents/{report_document_id}"
        document_response = requests.get(document_url, headers=headers)

        if document_response.status_code != 200:
            log_message(f"Error downloading report: {document_response.status_code}")
            continue

        document_info = document_response.json()
        report_download_url = document_info["url"]

        # Download the actual report content
        report_content_response = requests.get(report_download_url)

        # Check if the report is compressed
        if 'compressionAlgorithm' in document_info and document_info['compressionAlgorithm'] == 'GZIP':
            compressed_file = io.BytesIO(report_content_response.content)
            with gzip.GzipFile(fileobj=compressed_file) as f:
                report_content = f.read().decode('utf-8')
        else:
            report_content = report_content_response.text

        log_message(f"報告下載和解壓用時: {time.time() - download_start:.2f} 秒")

        # Save report to file
        file_name = "test"
        try:
            with open(file_name, 'w', encoding='utf-8') as f:
                f.write(report_content)
            log_message(f"報告已保存至: {file_name}")
        except Exception as e:
            log_message(f"保存文件時出錯: {e}")

    total_time = time.time() - start_time
    log_message(f"\n所有報告處理完成! 總用時: {total_time:.2f} 秒")

# 使用示例
if __name__ == "__main__":
    # 請確保 credentials 已定義

    
    amazon_order()