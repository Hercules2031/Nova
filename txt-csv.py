import csv
import sys
import os
from pathlib import Path

def convert_txt_to_csv(input_file, output_file=None):
    """
    將 Amazon 報告 TXT 文件轉換為 CSV 文件
    
    參數:
        input_file: 輸入的 TXT 文件路徑
        output_file: 輸出的 CSV 文件路徑 (可選)
    """
    
    # 如果沒有指定輸出文件，使用相同文件名但擴展名為 .csv
    if output_file is None:
        output_path = Path(input_file)
        output_file = output_path.with_suffix('.csv').name
    
    try:
        # 讀取 TXT 文件
        with open(input_file, 'r', encoding='utf-8') as txt_file:
            content = txt_file.read()
        
        # 分割行
        lines = content.strip().split('\n')
        
        if not lines:
            print("文件為空")
            return False
        
        # 檢測分隔符（可能是 Tab 或逗號）
        first_line = lines[0]
        if '\t' in first_line:
            delimiter = '\t'
            print("檢測到 Tab 分隔符")
        elif ',' in first_line:
            delimiter = ','
            print("檢測到逗號分隔符")
        else:
            # 嘗試空格分隔
            delimiter = ' '
            print("使用空格分隔符")
        
        # 解析數據
        data = []
        for i, line in enumerate(lines):
            if line.strip():  # 跳過空行
                # 使用分隔符分割行
                parts = line.split(delimiter)
                data.append(parts)
        
        # 寫入 CSV 文件
        with open(output_file, 'w', newline='', encoding='utf-8') as csv_file:
            csv_writer = csv.writer(csv_file)
            csv_writer.writerows(data)
        
        print(f"成功轉換！")
        print(f"輸入文件: {input_file}")
        print(f"輸出文件: {output_file}")
        print(f"總行數: {len(data)}")
        print(f"列數: {len(data[0]) if data else 0}")
        
        # 顯示前幾行作為預覽
        if data:
            print("\nCSV 文件預覽:")
            print("=" * 80)
            for i, row in enumerate(data[:5]):
                print(f"行 {i+1}: {row}")
            if len(data) > 5:
                print(f"... 還有 {len(data) - 5} 行")
            print("=" * 80)
        
        return True
        
    except Exception as e:
        print(f"轉換過程中出錯: {str(e)}")
        return False

def convert_folder(folder_path):
    """轉換文件夾中的所有 TXT 文件"""
    folder = Path(folder_path)
    txt_files = list(folder.glob('*.txt'))
    
    if not txt_files:
        print(f"在 {folder_path} 中沒有找到 TXT 文件")
        return
    
    print(f"找到 {len(txt_files)} 個 TXT 文件")
    
    success_count = 0
    for txt_file in txt_files:
        print(f"\n處理文件: {txt_file.name}")
        if convert_txt_to_csv(str(txt_file)):
            success_count += 1
    
    print(f"\n轉換完成！成功轉換 {success_count}/{len(txt_files)} 個文件")

def main():
    """主函數：提供命令行界面"""
    print("Amazon 報告 TXT 到 CSV 轉換器")
    print("=" * 50)
    
    if len(sys.argv) > 1:
        # 如果有命令行參數，使用第一個參數作為輸入文件
        input_file = sys.argv[1]
        
        if os.path.isdir(input_file):
            # 如果是文件夾，轉換文件夾中的所有文件
            convert_folder(input_file)
        else:
            # 如果是單個文件，轉換該文件
            convert_txt_to_csv(input_file)
    else:
        # 交互模式
        while True:
            print("\n請選擇操作:")
            print("1. 轉換單個文件")
            print("2. 轉換文件夾中的所有文件")
            print("3. 退出")
            
            choice = input("請輸入選項 (1-3): ").strip()
            
            if choice == '1':
                file_path = input("請輸入 TXT 文件路徑: ").strip()
                if os.path.exists(file_path):
                    convert_txt_to_csv(file_path)
                else:
                    print("文件不存在，請重新輸入")
            
            elif choice == '2':
                folder_path = input("請輸入文件夾路徑: ").strip()
                if os.path.exists(folder_path) and os.path.isdir(folder_path):
                    convert_folder(folder_path)
                else:
                    print("文件夾不存在，請重新輸入")
            
            elif choice == '3':
                print("再見！")
                break
            
            else:
                print("無效選項，請重新選擇")

if __name__ == "__main__":
    main()