import requests
import time
import gzip
import io
import json
from datetime import datetime, timedelta
from amazon_cred import credentials

def get_access_token():
    """獲取訪問令牌"""
    token_response = requests.post(
        "https://api.amazon.com/auth/o2/token",
        data={
            "grant_type": "refresh_token",
            "refresh_token": credentials["refresh_token"],
            "client_id": credentials["lwa_app_id"],
            "client_secret": credentials["lwa_client_secret"],
        },
    )
    return token_response.json()["access_token"]

def get_report():
    """獲取報告的完整流程"""
    # 1. 獲取訪問令牌
    print("獲取訪問令牌...")
    access_token = get_access_token()
    
    # 2. 設定時間範圍（最近30天）
    end_time = datetime.utcnow()
    start_time = end_time - timedelta(days=30)
    
    # 格式化為ISO 8601格式
    dataStartTime = start_time.isoformat()
    dataEndTime = end_time.isoformat()
    
    print(f"時間範圍: {dataStartTime} 到 {dataEndTime}")
    
    # 3. 創建報告請求
    url = "https://sellingpartnerapi-na.amazon.com/reports/2021-06-30/reports"
    
    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "x-amz-access-token": access_token
    }
    
    payload = {
        "reportType": "GET_XML_ALL_ORDERS_DATA_BY_ORDER_DATE_GENERAL",
        "dataStartTime": dataStartTime,
        "dataEndTime": dataEndTime,
        "marketplaceIds": ["ATVPDKIKX0DER"]  # 北美市場
    }
    
    print("創建報告請求...")
    response = requests.post(url, json=payload, headers=headers)
    
    if response.status_code != 202:
        print(f"創建報告失敗: {response.status_code}")
        print(response.text)
        return None
    
    report_info = response.json()
    report_id = report_info["reportId"]
    print(f"報告ID: {report_id}")
    
    # 4. 等待報告處理完成
    print("等待報告處理...")
    for i in range(30):  # 最多等待30次，每次30秒
        time.sleep(3)
        
        # 檢查報告狀態
        status_url = f"https://sellingpartnerapi-na.amazon.com/reports/2021-06-30/reports/{report_id}"
        status_response = requests.get(status_url, headers=headers)
        
        if status_response.status_code != 200:
            print(f"檢查狀態失敗: {status_response.status_code}")
            continue
        
        status_data = status_response.json()
        processing_status = status_data["processingStatus"]
        
        print(f"報告狀態 ({i+1}/30): {processing_status}")
        
        if processing_status == "DONE":
            report_document_id = status_data["reportDocumentId"]
            print(f"報告處理完成! 文檔ID: {report_document_id}")
            break
        elif processing_status in ["CANCELLED", "FATAL"]:
            print(f"報告處理失敗: {processing_status}")
            return None
    
    # 5. 下載報告
    print("下載報告...")
    document_url = f"https://sellingpartnerapi-na.amazon.com/reports/2021-06-30/documents/{report_document_id}"
    document_response = requests.get(document_url, headers=headers)
    
    if document_response.status_code != 200:
        print(f"下載報告失敗: {document_response.status_code}")
        return None
    
    document_info = document_response.json()
    download_url = document_info["url"]
    
    # 6. 從下載URL獲取報告內容
    report_response = requests.get(download_url)
    report_content = report_response.text
    
    # 7. 保存報告到文件
    filename = f"amazon_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(report_content)
    
    print(f"報告已保存到: {filename}")
    print(f"報告大小: {len(report_content)} 字符")
    
    return report_content

# 執行
if __name__ == "__main__":
    report = get_report()
    if report:
        print("\n報告前500字符:")
        print("=" * 50)
        print(report[:500])
        print("=" * 50)