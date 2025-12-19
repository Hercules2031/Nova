import requests
import time
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

def fetch_amazon_report():
    """獲取 Amazon 報告並直接存儲為 XML"""
    print("開始獲取 Amazon 銷售報告 (XML 存儲模式)")
    print("=" * 50)
    
    # 1. 獲取訪問令牌
    print("步驟 1: 獲取訪問令牌...")
    access_token = get_access_token()
    
    # 2. 設置報告參數
    report_type = "GET_XML_ALL_ORDERS_DATA_BY_ORDER_DATE_GENERAL"
    today = datetime.now()
    start_date = (today - timedelta(days=14)).strftime('%Y-%m-%d')
    end_date = today.strftime('%Y-%m-%d')
    
    # 3. 創建報告請求
    print("步驟 2: 創建報告請求...")
    url = "https://sellingpartnerapi-na.amazon.com/reports/2021-06-30/reports"
    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "x-amz-access-token": access_token
    }
    payload = {
        "reportType": report_type,
        "dataStartTime": f"{start_date}T00:00:00Z",
        "dataEndTime": f"{end_date}T23:59:59Z",
        "marketplaceIds": ["ATVPDKIKX0DER"]
    }
    
    response = requests.post(url, json=payload, headers=headers)
    if response.status_code != 202:
        print(f"創建報告失敗: {response.text}")
        return None
    
    report_id = response.json()["reportId"]
    print(f"報告ID: {report_id}")
    
    # 4. 等待報告處理完成
    print("步驟 3: 等待報告處理...")
    report_document_id = None
    for attempt in range(30):
        time.sleep(5) # 稍微增加等待時間以減少請求頻率
        status_url = f"https://sellingpartnerapi-na.amazon.com/reports/2021-06-30/reports/{report_id}"
        status_response = requests.get(status_url, headers=headers).json()
        status = status_response["processingStatus"]
        print(f"當前狀態: {status}")
        
        if status == "DONE":
            report_document_id = status_response["reportDocumentId"]
            break
        elif status in ["CANCELLED", "FATAL"]:
            return None

    if not report_document_id:
        return None

    # 5. 下載報告內容
    print("步驟 4: 下載並儲存 XML...")
    doc_url = f"https://sellingpartnerapi-na.amazon.com/reports/2021-06-30/documents/{report_document_id}"
    doc_response = requests.get(doc_url, headers=headers).json()
    
    # 獲取實際的下載連結並下載
    final_report = requests.get(doc_response["url"])
    
    # 6. 保存為 XML 文件
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"amazon_report_{timestamp}.xml"
    
    with open(filename, 'wb') as f: # 使用 'wb' 模式以確保正確處理編碼
        f.write(final_report.content)
    
    return filename

if __name__ == "__main__":
    xml_file = fetch_amazon_report()
    if xml_file:
        print(f"\n成功！原始 XML 已保存至: {xml_file}")
    else:
        print("\n任務失敗")