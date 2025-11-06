#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
#3形式のCSVを月毎に縦に並ぶ形式のCSVに変換（標準ライブラリのみ使用）
"""

import csv
import re
from datetime import datetime

def convert_date_format(date_str):
    """日付形式を変換（Apr-25 -> 2025-04）"""
    month_map = {
        'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04',
        'May': '05', 'Jun': '06', 'Jul': '07', 'Aug': '08',
        'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12'
    }
    
    # Apr-25形式を検出
    match = re.match(r'([A-Za-z]+)-(\d+)', date_str)
    if match:
        month_str = match.group(1)
        year_str = match.group(2)
        
        if month_str in month_map:
            month = month_map[month_str]
            year = '20' + year_str if len(year_str) == 2 else year_str
            return f"{year}-{month}"
    
    return date_str

def parse_number(value):
    """数値を変換（カンマ区切りを削除）"""
    if not value or value == '-' or value == '':
        return 0
    
    # カンマと引用符を削除
    value_str = str(value).replace(',', '').replace('"', '').strip()
    
    try:
        return float(value_str)
    except:
        return 0

def load_budget_data(budget_file):
    """予算データを読み込む（#3形式と同じ構造）"""
    if not budget_file:
        return {}
    
    print(f"予算データを読み込み中: {budget_file}")
    
    budget_dict = {}
    
    try:
        with open(budget_file, 'r', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            rows = list(reader)
        
        # ヘッダー情報を抽出（行0-7）
        header_info = {
            'level1': rows[0] if len(rows) > 0 else [],
            'level2': rows[1] if len(rows) > 1 else [],
            'level3': rows[2] if len(rows) > 2 else [],
        }
        
        # データ部分（行8以降）
        data_rows = rows[8:] if len(rows) > 8 else []
        
        # 予算データを辞書に格納 {('Year_Month', 'Account_Code'): amount}
        for data_row in data_rows:
            if not data_row or len(data_row) == 0:
                continue
            
            date_value = data_row[0] if len(data_row) > 0 else ''
            date_str = convert_date_format(date_value)
            
            # 各列（Account Code）を処理
            for col_idx in range(1, len(data_row)):
                # Account Codeを抽出
                account_code = None
                
                if col_idx < len(header_info['level3']) and header_info['level3'][col_idx]:
                    account_code = header_info['level3'][col_idx].strip()
                elif col_idx < len(header_info['level2']) and header_info['level2'][col_idx]:
                    account_code = header_info['level2'][col_idx].strip()
                elif col_idx < len(header_info['level1']) and header_info['level1'][col_idx]:
                    account_code = header_info['level1'][col_idx].strip()
                
                if account_code and account_code != '':
                    amount = parse_number(data_row[col_idx])
                    key = (date_str, account_code)
                    budget_dict[key] = amount
        
        print(f"予算データ読み込み完了: {len(budget_dict)}件のレコード")
        
    except FileNotFoundError:
        print(f"警告: 予算データファイルが見つかりません: {budget_file}")
    except Exception as e:
        print(f"警告: 予算データの読み込み中にエラーが発生: {e}")
    
    return budget_dict

def convert_csv_to_long_format(input_file, output_file, budget_file=None):
    """CSVをロング形式に変換"""
    print(f"CSVファイルを読み込み中: {input_file}")
    
    # 予算データを読み込む（オプション）
    budget_dict = load_budget_data(budget_file)
    
    with open(input_file, 'r', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        rows = list(reader)
    
    # ヘッダー情報を抽出（行0-7）
    header_info = {
        'level1': rows[0] if len(rows) > 0 else [],
        'level2': rows[1] if len(rows) > 1 else [],
        'level3': rows[2] if len(rows) > 2 else [],
        'internal_code': rows[3] if len(rows) > 3 else [],
        'account_name': rows[4] if len(rows) > 4 else [],
        'keiei_houkoku': rows[5] if len(rows) > 5 else [],
        'detail_jp': rows[6] if len(rows) > 6 else [],
        'description_en': rows[7] if len(rows) > 7 else [],
    }
    
    # データ部分（行8以降）
    data_rows = rows[8:] if len(rows) > 8 else []
    
    print(f"ヘッダー行数: 8行")
    print(f"データ行数: {len(data_rows)}行")
    
    # 結果を格納
    result_rows = []
    
    # 各データ行を処理
    for data_row in data_rows:
        if not data_row or len(data_row) == 0:
            continue
        
        date_value = data_row[0] if len(data_row) > 0 else ''
        date_str = convert_date_format(date_value)
        
        # 各列（Account Code）を処理
        for col_idx in range(1, len(data_row)):
            # Account Codeを抽出
            account_code = None
            level = 0
            
            # レベル3から順に確認
            if col_idx < len(header_info['level3']) and header_info['level3'][col_idx]:
                account_code = header_info['level3'][col_idx].strip()
                if account_code:
                    level = 3
            elif col_idx < len(header_info['level2']) and header_info['level2'][col_idx]:
                account_code = header_info['level2'][col_idx].strip()
                if account_code:
                    level = 2
            elif col_idx < len(header_info['level1']) and header_info['level1'][col_idx]:
                account_code = header_info['level1'][col_idx].strip()
                if account_code:
                    level = 1
            
            if account_code and account_code != '':
                # 数値を変換
                amount = parse_number(data_row[col_idx])
                
                # メタ情報を取得
                internal_code = header_info['internal_code'][col_idx] if col_idx < len(header_info['internal_code']) else ''
                account_name = header_info['account_name'][col_idx] if col_idx < len(header_info['account_name']) else ''
                keiei_houkoku = header_info['keiei_houkoku'][col_idx] if col_idx < len(header_info['keiei_houkoku']) else ''
                detail_jp = header_info['detail_jp'][col_idx] if col_idx < len(header_info['detail_jp']) else ''
                description_en = header_info['description_en'][col_idx] if col_idx < len(header_info['description_en']) else ''
                
                # 予算データを取得
                budget_key = (date_str, account_code)
                budget_amount = budget_dict.get(budget_key, 0)
                
                # 結果に追加
                # Plan用とActual用の2行を作成（ロング形式）
                # これによりLooker StudioでTypeでフィルタリングできる
                
                # Plan行（予算データ）
                result_rows.append({
                    'Year_Month': date_str,
                    'Date': date_value,
                    'Account_Code': account_code,
                    'Level': level,
                    'Internal_code': internal_code,
                    'Internal_Account_Name': account_name,
                    'KEIEI_houkoku': keiei_houkoku,
                    '詳細': detail_jp,
                    'Description': description_en,
                    'Type': 'Plan',
                    'Amount': budget_amount  # 予算データ
                })
                
                # Actual行（実績データ）
                result_rows.append({
                    'Year_Month': date_str,
                    'Date': date_value,
                    'Account_Code': account_code,
                    'Level': level,
                    'Internal_code': internal_code,
                    'Internal_Account_Name': account_name,
                    'KEIEI_houkoku': keiei_houkoku,
                    '詳細': detail_jp,
                    'Description': description_en,
                    'Type': 'Actual',
                    'Amount': amount  # 実績データ
                })
    
    print(f"変換完了: {len(result_rows)}件のレコード")
    
    # CSVに出力
    print(f"CSVファイルに出力中: {output_file}")
    
    with open(output_file, 'w', encoding='utf-8-sig', newline='') as f:
        fieldnames = [
            'Year_Month',
            'Date',
            'Account_Code',
            'Level',
            'Internal_code',
            'Internal_Account_Name',
            'KEIEI_houkoku',
            '詳細',
            'Description',
            'Type',
            'Amount'
        ]
        
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        
        # ソート（年月、Account Code順）
        result_rows.sort(key=lambda x: (x['Year_Month'], x['Account_Code']))
        
        for row in result_rows:
            writer.writerow(row)
    
    print(f"出力完了: {output_file}")
    print(f"総レコード数: {len(result_rows)}件")
    
    # 統計情報
    if result_rows:
        year_months = set(row['Year_Month'] for row in result_rows)
        account_codes = set(row['Account_Code'] for row in result_rows)
        print(f"年月数: {len(year_months)}ヶ月")
        print(f"Account Code数: {len(account_codes)}件")
        
        # サンプルデータを表示
        print("\nサンプルデータ（最初の5件）:")
        for i, row in enumerate(result_rows[:5], 1):
            print(f"{i}. {row['Year_Month']} | {row['Account_Code']} | {row['Type']} | {row['Amount']:,.0f}")

def main():
    """メイン処理"""
    input_file = '/Users/sugenotomohiro/Desktop/CURSOR/Manabox/20_DATA_Input/01.3consulting/Glory　Format - #3 sample(DB for BI report).csv'
    
    # 予算データファイル（オプション）
    # 予算データがある場合は、ファイルパスを指定してください
    budget_file = None  # 例: 'path/to/budget_data.csv'
    
    output_file = '/Users/sugenotomohiro/Desktop/CURSOR/Manabox/20_DATA_Input/01.3consulting/月次DB_SharePoint用.csv'
    
    try:
        convert_csv_to_long_format(input_file, output_file, budget_file)
        print("\n✅ 処理完了！")
        print(f"出力ファイル: {output_file}")
        print("\n次のステップ:")
        print("1. 出力されたCSVファイルをエクセルで開く")
        print("2. エクセルでスタイルを適用（必要に応じて）")
        print("3. SharePointにアップロード")
        
    except Exception as e:
        print(f"❌ エラーが発生しました: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    main()

