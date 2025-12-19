import requests
import urllib.parse
from datetime import datetime, timedelta
from amazon_cred import credentials

def test_single_market():
    # 目标 SKU
    target_sku = "DT-009"
    
    # 选择一个市场测试（这里用美国市场）
    marketplace_id = "ATVPDKIKX0DER"
    marketplace_name = "United States"
    
    print(f"測試 SKU: {target_sku}")
    print(f"測試市場: {marketplace_name} ({marketplace_id})")
    
    # 获取 access token
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
    print("Access Token 獲取成功")

    # 计算日期范围
    today = datetime.now()
    last_30_days_start = (today - timedelta(days=120)).strftime('%Y-%m-%d')
    last_30_days_end = today.strftime('%Y-%m-%d')
    
    # 编码 SKU
    encoded_sku = urllib.parse.quote(target_sku, safe='')
    
    # 构建 URL（只测试30天数据）
    url = f"https://sellingpartnerapi-na.amazon.com/sales/v1/orderMetrics?marketplaceIds={marketplace_id}&interval={last_30_days_start}T00:00:00-08:00--{last_30_days_end}T00:00:00-08:00&granularity=Total&buyerType=All&firstDayOfWeek=Monday&sku={encoded_sku}"
    
    print(f"請求 URL: {url}")
    
    headers = {"x-amz-access-token": access_token}
    
    try:
        response = requests.get(url, headers=headers, timeout=10)
        print(f"HTTP 狀態碼: {response.status_code}")
        
        if response.status_code == 200:
            data = response.json()
            print("響應數據:", data)
            
            if 'payload' in data and len(data['payload']) > 0:
                order_count = data['payload'][0].get('orderItemCount', 0)
                print(f"✅ 成功! SKU {target_sku} 在 {marketplace_name} 的30天銷售量: {order_count}")
            else:
                print("⚠️  沒有找到銷售數據（可能該SKU在此市場無銷售）")
                
        elif response.status_code == 400:
            print("❌ 請求錯誤 - 可能市場ID或SKU不正確")
            print("錯誤詳情:", response.text)
        elif response.status_code == 403:
            print("❌ 權限錯誤 - 檢查Access Token或API權限")
        else:
            print(f"❌ 其他錯誤: {response.status_code}")
            print("錯誤詳情:", response.text)
            
    except Exception as e:
        print(f"❌ 請求失敗: {str(e)}")

if __name__ == "__main__":
    test_single_market()