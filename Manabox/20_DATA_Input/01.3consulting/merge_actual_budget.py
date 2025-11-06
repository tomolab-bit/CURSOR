#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
#1（実績）と#2（予算）のCSVを結合して#3形式（Plan/Actual）の月次DBを作成
"""

import csv
import re
from collections import defaultdict

def convert_date_format(date_str):
    """日付形式を変換（Apr-25 -> 2025-04, Apr-26 -> 2025-04）"""
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
            # #2のApr-26は実際には2025年4月の予算データとして扱う
            # 実績と予算の年を統一（2025年として扱う）
            year = '2025'
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

def extract_account_code(row, col_idx):
    """行からAccount Codeを抽出"""
    if col_idx < len(row) and row[col_idx]:
        code = row[col_idx].strip()
        # Account Codeのレベルを判定（インデントの数で判定）
        level = 0
        if code:
            # インデントがない場合（レベル1）
            if not code.startswith(' '):
                level = 1
            # 1つのインデント（レベル2）
            elif code.startswith(' ') and not code.startswith('  '):
                level = 2
            # 2つのインデント（レベル3）
            else:
                level = 3
            return code.strip(), level
    return None, 0

def load_actual_data(actual_file):
    """#1実績データを読み込む"""
    print(f"実績データを読み込み中: {actual_file}")
    
    actual_data = defaultdict(dict)  # {(date_str, account_code): amount}
    account_info = {}  # {account_code: {level, internal_code, account_name, ...}}
    
    with open(actual_file, 'r', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        rows = list(reader)
    
    # ヘッダー行（行1）
    header_row = rows[0] if len(rows) > 0 else []
    
    # 月の列を特定（Apr-25, May-25, Jun-25など）
    month_columns = {}
    for col_idx, header in enumerate(header_row):
        if header and re.match(r'[A-Za-z]+-\d+', header):
            month_columns[col_idx] = header
    
    print(f"月の列を検出: {list(month_columns.values())}")
    
    # データ行（行3以降、行2は空行）
    data_rows = rows[2:] if len(rows) > 2 else []
    
    current_account_code = None
    current_level = 0
    
    for row in data_rows:
        if not row or len(row) == 0:
            continue
        
        # Account Codeを取得（最初の非空列を探す：列0、列1、列2）
        account_code = None
        level = 0
        
        for col_idx in range(3):  # 列0、列1、列2をチェック
            if col_idx < len(row) and row[col_idx] and row[col_idx].strip():
                account_code = row[col_idx].strip()
                level = col_idx + 1  # 列0=レベル1, 列1=レベル2, 列2=レベル3
                break
        
        # Description（Chỉ tiêu列、列8）を取得
        description_chi_tieu = row[8] if len(row) > 8 else ''
        
        # Account Codeがある場合、新しいAccount Codeとして更新
        if account_code:
            current_account_code = account_code
            current_level = level
            
            # メタ情報を保存
            internal_code = row[3] if len(row) > 3 else ''
            account_name = row[4] if len(row) > 4 else ''
            keiei_houkoku = row[5] if len(row) > 5 else ''
            detail_jp = row[6] if len(row) > 6 else ''
            description_en = row[7] if len(row) > 7 else ''
            
            # Description（Chỉ tiêu）がない場合は、Description（列7）を使用
            if not description_chi_tieu:
                description_chi_tieu = description_en
            
            account_info[current_account_code] = {
                'level': current_level,
                'internal_code': internal_code,
                'account_name': account_name,
                'keiei_houkoku': keiei_houkoku,
                'detail_jp': detail_jp,
                'description_en': description_en,
                'description_chi_tieu': description_chi_tieu
            }
        else:
            # Account Codeがない場合、Description（Chỉ tiêu）があるかチェック
            if description_chi_tieu and description_chi_tieu.strip():
                # Description（Chỉ tiêu）がある場合は、前のAccount Codeに紐付ける
                # ただし、金額が入っている場合のみ
                pass
        
        # 月の列のデータを取得（Account Codeがある場合、または金額が入っている場合）
        if current_account_code:
            # 金額が入っているかチェック
            has_amount = False
            for col_idx, month_header in month_columns.items():
                if col_idx < len(row):
                    amount = parse_number(row[col_idx])
                    if amount != 0:
                        has_amount = True
                        break
            
            # Account Codeがある場合、または金額が入っている場合
            if account_code or has_amount:
                # Description（Chỉ tiêu）を取得（なければDescription列を使用）
                if not description_chi_tieu:
                    description_chi_tieu = row[7] if len(row) > 7 else ''
                
                # キーにDescriptionを含める（サブ項目を区別するため）
                # Descriptionがない場合はAccount Codeのみ
                if description_chi_tieu and description_chi_tieu.strip():
                    unique_key = f"{current_account_code}_{description_chi_tieu.strip()}"
                else:
                    unique_key = current_account_code
                
                for col_idx, month_header in month_columns.items():
                    if col_idx < len(row):
                        amount = parse_number(row[col_idx])
                        # 金額が0でもAccount Codeを保持（全てのAccount Codeを記録）
                        date_str = convert_date_format(month_header)
                        key = (date_str, unique_key)
                        actual_data[key] = amount
                        
                        # Descriptionがある場合は、account_infoにも保存
                        if description_chi_tieu and description_chi_tieu.strip():
                            if unique_key not in account_info:
                                # 親Account Codeの情報をコピー
                                parent_info = account_info.get(current_account_code, {})
                                account_info[unique_key] = {
                                    'level': parent_info.get('level', 0),
                                    'internal_code': parent_info.get('internal_code', ''),
                                    'account_name': parent_info.get('account_name', ''),
                                    'keiei_houkoku': parent_info.get('keiei_houkoku', ''),
                                    'detail_jp': parent_info.get('detail_jp', ''),
                                    'description_en': description_chi_tieu.strip(),
                                    'description_chi_tieu': description_chi_tieu.strip(),
                                    'parent_account_code': current_account_code
                                }
    
    print(f"実績データ読み込み完了: {len(actual_data)}件のレコード")
    return actual_data, account_info

def load_budget_data(budget_file):
    """#2予算データを読み込む"""
    print(f"予算データを読み込み中: {budget_file}")
    
    budget_data = defaultdict(dict)  # {(date_str, account_code): amount}
    account_info = {}  # {account_code: {level, internal_code, account_name, ...}}
    
    with open(budget_file, 'r', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        rows = list(reader)
    
    # ヘッダー行（行6）
    header_row = rows[5] if len(rows) > 5 else []
    
    # 月の列を特定（Apr-26, May-26, Jun-26など）
    month_columns = {}
    for col_idx, header in enumerate(header_row):
        if header and re.match(r'[A-Za-z]+-\d+', header):
            month_columns[col_idx] = header
    
    print(f"月の列を検出: {list(month_columns.values())}")
    
    # データ行（行8以降、行7は空行）
    data_rows = rows[7:] if len(rows) > 7 else []
    
    current_account_code = None
    current_level = 0
    
    for row in data_rows:
        if not row or len(row) == 0:
            continue
        
        # Account Codeを取得（最初の非空列を探す：列0、列1、列2）
        account_code = None
        level = 0
        
        for col_idx in range(3):  # 列0、列1、列2をチェック
            if col_idx < len(row) and row[col_idx] and row[col_idx].strip():
                account_code = row[col_idx].strip()
                level = col_idx + 1  # 列0=レベル1, 列1=レベル2, 列2=レベル3
                break
        
        # Description（Chỉ tiêu列、列8）を取得
        description_chi_tieu = row[8] if len(row) > 8 else ''
        
        # Account Codeがある場合、新しいAccount Codeとして更新
        if account_code:
            current_account_code = account_code
            current_level = level
            
            # メタ情報を保存
            internal_code = row[3] if len(row) > 3 else ''
            account_name = row[4] if len(row) > 4 else ''
            keiei_houkoku = row[5] if len(row) > 5 else ''
            detail_jp = row[6] if len(row) > 6 else ''
            description_en = row[7] if len(row) > 7 else ''
            
            # Description（Chỉ tiêu）がない場合は、Description（列7）を使用
            if not description_chi_tieu:
                description_chi_tieu = description_en
            
            account_info[current_account_code] = {
                'level': current_level,
                'internal_code': internal_code,
                'account_name': account_name,
                'keiei_houkoku': keiei_houkoku,
                'detail_jp': detail_jp,
                'description_en': description_en,
                'description_chi_tieu': description_chi_tieu
            }
        
        # 月の列のデータを取得（Account Codeがある場合、または金額が入っている場合）
        if current_account_code:
            # 金額が入っているかチェック
            has_amount = False
            for col_idx, month_header in month_columns.items():
                if col_idx < len(row):
                    amount = parse_number(row[col_idx])
                    if amount != 0:
                        has_amount = True
                        break
            
            # Account Codeがある場合、または金額が入っている場合
            if account_code or has_amount:
                # Description（Chỉ tiêu）を取得（なければDescription列を使用）
                if not description_chi_tieu:
                    description_chi_tieu = row[7] if len(row) > 7 else ''
                
                # キーにDescriptionを含める（サブ項目を区別するため）
                # Descriptionがない場合はAccount Codeのみ
                if description_chi_tieu and description_chi_tieu.strip():
                    unique_key = f"{current_account_code}_{description_chi_tieu.strip()}"
                else:
                    unique_key = current_account_code
                
                for col_idx, month_header in month_columns.items():
                    if col_idx < len(row):
                        amount = parse_number(row[col_idx])
                        # 金額が0でもAccount Codeを保持（全てのAccount Codeを記録）
                        date_str = convert_date_format(month_header)
                        key = (date_str, unique_key)
                        budget_data[key] = amount
                        
                        # Descriptionがある場合は、account_infoにも保存
                        if description_chi_tieu and description_chi_tieu.strip():
                            if unique_key not in account_info:
                                # 親Account Codeの情報をコピー
                                parent_info = account_info.get(current_account_code, {})
                                account_info[unique_key] = {
                                    'level': parent_info.get('level', 0),
                                    'internal_code': parent_info.get('internal_code', ''),
                                    'account_name': parent_info.get('account_name', ''),
                                    'keiei_houkoku': parent_info.get('keiei_houkoku', ''),
                                    'detail_jp': parent_info.get('detail_jp', ''),
                                    'description_en': description_chi_tieu.strip(),
                                    'description_chi_tieu': description_chi_tieu.strip(),
                                    'parent_account_code': current_account_code
                                }
    
    print(f"予算データ読み込み完了: {len(budget_data)}件のレコード")
    return budget_data, account_info

def merge_to_long_format(actual_file, budget_file, output_file):
    """実績と予算データを結合してロング形式で出力"""
    
    # データを読み込む
    actual_data, actual_account_info = load_actual_data(actual_file)
    budget_data, budget_account_info = load_budget_data(budget_file)
    
    # アカウント情報を統合（実績を優先、なければ予算）
    account_info = {**budget_account_info, **actual_account_info}
    
    # すべてのキー（date_str, account_code）を収集
    all_keys = set(actual_data.keys()) | set(budget_data.keys())
    
    # 結果を格納
    result_rows = []
    
    for date_str, unique_key in sorted(all_keys):
        # メタ情報を取得
        info = account_info.get(unique_key, {})
        
        # unique_keyからAccount Codeを抽出（親Account Codeを取得）
        if '_' in unique_key:
            # unique_keyが"account_code_description"形式の場合
            parent_account_code = info.get('parent_account_code', unique_key.split('_')[0])
            account_code = parent_account_code
        else:
            account_code = unique_key
        
        # Descriptionを取得（description_chi_tieuがあれば使用、なければdescription_en）
        description = info.get('description_chi_tieu', '') or info.get('description_en', '')
        
        # 実績と予算の金額を取得
        actual_amount = actual_data.get((date_str, unique_key), 0)
        budget_amount = budget_data.get((date_str, unique_key), 0)
        
        # Plan行（予算データ）
        result_rows.append({
            'Year_Month': date_str,
            'Date': date_str.replace('-', ' '),  # 2025-04 -> 2025 04
            'Account_Code': account_code,
            'Level': info.get('level', 0),
            'Internal_code': info.get('internal_code', ''),
            'Internal_Account_Name': info.get('account_name', ''),
            'KEIEI_houkoku': info.get('keiei_houkoku', ''),
            '詳細': info.get('detail_jp', ''),
            'Description': description,
            'Type': 'Plan',
            'Amount': budget_amount
        })
        
        # Actual行（実績データ）
        result_rows.append({
            'Year_Month': date_str,
            'Date': date_str.replace('-', ' '),
            'Account_Code': account_code,
            'Level': info.get('level', 0),
            'Internal_code': info.get('internal_code', ''),
            'Internal_Account_Name': info.get('account_name', ''),
            'KEIEI_houkoku': info.get('keiei_houkoku', ''),
            '詳細': info.get('detail_jp', ''),
            'Description': description,
            'Type': 'Actual',
            'Amount': actual_amount
        })
    
    # CSVに出力
    fieldnames = [
        'Year_Month', 'Date', 'Account_Code', 'Level', 'Internal_code',
        'Internal_Account_Name', 'KEIEI_houkoku', '詳細', 'Description',
        'Type', 'Amount'
    ]
    
    with open(output_file, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(result_rows)
    
    print(f"\n✅ 処理完了！")
    print(f"出力ファイル: {output_file}")
    print(f"総レコード数: {len(result_rows)}行（Plan/Actual各{len(result_rows)//2}件）")
    
    # サンプル出力
    if result_rows:
        print("\n【サンプルデータ（最初の5行）】")
        for i, row in enumerate(result_rows[:5]):
            print(f"{i+1}. {row['Year_Month']} | {row['Account_Code']} | {row['Type']} | {row['Amount']}")

def main():
    """メイン処理"""
    actual_file = '/Users/sugenotomohiro/Desktop/CURSOR/Manabox/20_DATA_Input/01.3consulting/Glory　Format - #1MonthlyPL(Actual) .csv'
    budget_file = '/Users/sugenotomohiro/Desktop/CURSOR/Manabox/20_DATA_Input/01.3consulting/Glory　Format - #2Monthly PL(Budget).csv'
    output_file = '/Users/sugenotomohiro/Desktop/CURSOR/Manabox/20_DATA_Input/01.3consulting/月次DB_SharePoint用.csv'
    
    try:
        merge_to_long_format(actual_file, budget_file, output_file)
        print("\n次のステップ:")
        print("1. 出力されたCSVファイルをエクセルで開く")
        print("2. エクセルでスタイルを適用（必要に応じて）")
        print("3. SharePointにアップロード")
        print("4. Looker Studioで接続してType列でPlan/Actualをフィルタリング")
        
    except Exception as e:
        print(f"エラーが発生しました: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    main()

