#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
#1（実績）と#2（予算）のCSVを結合して#3形式（横型、ワイド形式）の月次DBを作成
各Account Codeに対してActual列とBudget列を追加
"""

import csv
import re
from collections import defaultdict, OrderedDict

def convert_date_format(date_str):
    """日付形式を変換（Apr-25 -> 2025-04, Apr-26 -> 2025-04）"""
    month_map = {
        'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04',
        'May': '05', 'Jun': '06', 'Jul': '07', 'Aug': '08',
        'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12'
    }
    
    match = re.match(r'([A-Za-z]+)-(\d+)', date_str)
    if match:
        month_str = match.group(1)
        year_str = match.group(2)
        
        if month_str in month_map:
            month = month_map[month_str]
            year = '2025'
            return f"{year}-{month}"
    
    return date_str

def parse_number(value):
    """数値を変換（カンマ区切りを削除）"""
    if not value or value == '-' or value == '':
        return 0
    
    value_str = str(value).replace(',', '').replace('"', '').strip()
    
    try:
        return float(value_str)
    except:
        return 0

def load_master_format(master_file):
    """マスターフォーマット#3からAccount Code構造を読み込む"""
    print(f"マスターフォーマットを読み込み中: {master_file}")
    
    account_structure = OrderedDict()  # {account_code: {level, internal_code, account_name, ...}}
    
    with open(master_file, 'r', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        rows = list(reader)
    
    # ヘッダー情報を取得（行0-7）
    if len(rows) < 8:
        print("警告: マスターフォーマットの形式が正しくありません")
        return account_structure
    
    # 行3-7からAccount Code構造を抽出
    level3 = rows[2] if len(rows) > 2 else []  # レベル3 Account Code
    level2 = rows[1] if len(rows) > 1 else []  # レベル2 Account Code
    level1 = rows[0] if len(rows) > 0 else []  # レベル1 Account Code
    internal_code = rows[3] if len(rows) > 3 else []  # Internal Code
    account_name = rows[4] if len(rows) > 4 else []  # Internal Account Name
    keiei_houkoku = rows[5] if len(rows) > 5 else []  # KEIEI houkoku
    detail_jp = rows[6] if len(rows) > 6 else []  # 詳細
    description = rows[7] if len(rows) > 7 else []  # Description
    
    # 各列を処理
    for col_idx in range(len(level1)):
        account_code = None
        level = 0
        
        # レベル3から順に確認
        if col_idx < len(level3) and level3[col_idx] and level3[col_idx].strip():
            account_code = level3[col_idx].strip()
            level = 3
        elif col_idx < len(level2) and level2[col_idx] and level2[col_idx].strip():
            account_code = level2[col_idx].strip()
            level = 2
        elif col_idx < len(level1) and level1[col_idx] and level1[col_idx].strip():
            account_code = level1[col_idx].strip()
            level = 1
        
        if account_code:
            account_structure[account_code] = {
                'col_idx': col_idx,
                'level': level,
                'internal_code': internal_code[col_idx] if col_idx < len(internal_code) else '',
                'account_name': account_name[col_idx] if col_idx < len(account_name) else '',
                'keiei_houkoku': keiei_houkoku[col_idx] if col_idx < len(keiei_houkoku) else '',
                'detail_jp': detail_jp[col_idx] if col_idx < len(detail_jp) else '',
                'description': description[col_idx] if col_idx < len(description) else '',
            }
    
    print(f"マスターフォーマット読み込み完了: {len(account_structure)}件のAccount Code")
    return account_structure

def load_actual_data(actual_file):
    """#1実績データを読み込む"""
    print(f"実績データを読み込み中: {actual_file}")
    
    actual_data = defaultdict(dict)  # {(date_str, account_code): amount}
    
    with open(actual_file, 'r', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        rows = list(reader)
    
    header_row = rows[0] if len(rows) > 0 else []
    
    # 月の列を特定
    month_columns = {}
    for col_idx, header in enumerate(header_row):
        if header and re.match(r'[A-Za-z]+-\d+', header):
            month_columns[col_idx] = header
    
    print(f"月の列を検出: {list(month_columns.values())}")
    
    data_rows = rows[2:] if len(rows) > 2 else []
    
    current_account_code = None
    
    for row in data_rows:
        if not row or len(row) == 0:
            continue
        
        # Account Codeを取得
        account_code = None
        for col_idx in range(3):
            if col_idx < len(row) and row[col_idx] and row[col_idx].strip():
                account_code = row[col_idx].strip()
                break
        
        if account_code:
            current_account_code = account_code
        
        # Description（Chỉ tiêu列、列8）を取得
        description_chi_tieu = row[8] if len(row) > 8 else ''
        
        # 月の列のデータを取得
        if current_account_code:
            has_amount = False
            for col_idx, month_header in month_columns.items():
                if col_idx < len(row):
                    amount = parse_number(row[col_idx])
                    if amount != 0:
                        has_amount = True
                    if account_code or has_amount:
                        date_str = convert_date_format(month_header)
                        # Descriptionがある場合はキーに含める
                        if description_chi_tieu and description_chi_tieu.strip():
                            unique_key = f"{current_account_code}_{description_chi_tieu.strip()}"
                        else:
                            unique_key = current_account_code
                        key = (date_str, unique_key)
                        actual_data[key] = amount
    
    print(f"実績データ読み込み完了: {len(actual_data)}件のレコード")
    return actual_data

def load_budget_data(budget_file):
    """#2予算データを読み込む"""
    print(f"予算データを読み込み中: {budget_file}")
    
    budget_data = defaultdict(dict)  # {(date_str, account_code): amount}
    
    with open(budget_file, 'r', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        rows = list(reader)
    
    header_row = rows[5] if len(rows) > 5 else []
    
    # 月の列を特定
    month_columns = {}
    for col_idx, header in enumerate(header_row):
        if header and re.match(r'[A-Za-z]+-\d+', header):
            month_columns[col_idx] = header
    
    print(f"月の列を検出: {list(month_columns.values())}")
    
    data_rows = rows[7:] if len(rows) > 7 else []
    
    current_account_code = None
    
    for row in data_rows:
        if not row or len(row) == 0:
            continue
        
        # Account Codeを取得
        account_code = None
        for col_idx in range(3):
            if col_idx < len(row) and row[col_idx] and row[col_idx].strip():
                account_code = row[col_idx].strip()
                break
        
        if account_code:
            current_account_code = account_code
        
        # Description（Chỉ tiêu列、列8）を取得
        description_chi_tieu = row[8] if len(row) > 8 else ''
        
        # 月の列のデータを取得
        if current_account_code:
            has_amount = False
            for col_idx, month_header in month_columns.items():
                if col_idx < len(row):
                    amount = parse_number(row[col_idx])
                    if amount != 0:
                        has_amount = True
                    if account_code or has_amount:
                        date_str = convert_date_format(month_header)
                        # Descriptionがある場合はキーに含める
                        if description_chi_tieu and description_chi_tieu.strip():
                            unique_key = f"{current_account_code}_{description_chi_tieu.strip()}"
                        else:
                            unique_key = current_account_code
                        key = (date_str, unique_key)
                        budget_data[key] = amount
    
    print(f"予算データ読み込み完了: {len(budget_data)}件のレコード")
    return budget_data

def create_wide_format_db(master_file, actual_file, budget_file, output_file):
    """横型（ワイド形式）のDBを作成"""
    
    # マスターフォーマットを読み込む
    account_structure = load_master_format(master_file)
    
    # 実績と予算データを読み込む
    actual_data = load_actual_data(actual_file)
    budget_data = load_budget_data(budget_file)
    
    # すべての日付を収集
    all_dates = set()
    for date_str, _ in actual_data.keys():
        all_dates.add(date_str)
    for date_str, _ in budget_data.keys():
        all_dates.add(date_str)
    
    all_dates = sorted(all_dates)
    
    # すべてのAccount Code（Description含む）を収集
    all_account_codes = set()
    account_code_map = {}  # {parent_code: [description付きキーのリスト]}
    
    for _, account_code in actual_data.keys():
        all_account_codes.add(account_code)
        # Description付きのAccount Codeの場合、親コードを抽出
        if '_' in account_code:
            parent_code = account_code.split('_')[0]
            if parent_code not in account_code_map:
                account_code_map[parent_code] = []
            account_code_map[parent_code].append(account_code)
        else:
            if account_code not in account_code_map:
                account_code_map[account_code] = []
    
    for _, account_code in budget_data.keys():
        all_account_codes.add(account_code)
        # Description付きのAccount Codeの場合、親コードを抽出
        if '_' in account_code:
            parent_code = account_code.split('_')[0]
            if parent_code not in account_code_map:
                account_code_map[parent_code] = []
            if account_code not in account_code_map[parent_code]:
                account_code_map[parent_code].append(account_code)
        else:
            if account_code not in account_code_map:
                account_code_map[account_code] = []
    
    # マスターフォーマットのAccount Codeも含める
    for account_code in account_structure.keys():
        all_account_codes.add(account_code)
        if account_code not in account_code_map:
            account_code_map[account_code] = []
    
    # Account Codeを順序付きで整理（マスターフォーマットの順序を維持）
    ordered_account_codes = []
    seen_codes = set()
    
    # まず、マスターフォーマットの順序で追加
    for account_code in account_structure.keys():
        ordered_account_codes.append(account_code)
        seen_codes.add(account_code)
    
    # 次に、Description付きのAccount Codeを追加（親コードの後に追加）
    for parent_code in account_code_map.keys():
        if parent_code in ordered_account_codes:
            # 親コードの位置を取得
            parent_idx = ordered_account_codes.index(parent_code)
            # Description付きのAccount Codeを親コードの直後に追加
            for desc_code in account_code_map[parent_code]:
                if desc_code not in seen_codes and '_' in desc_code:
                    ordered_account_codes.insert(parent_idx + 1, desc_code)
                    seen_codes.add(desc_code)
                    parent_idx += 1
        else:
            # 親コードがマスターフォーマットにない場合、追加
            if parent_code not in seen_codes:
                ordered_account_codes.append(parent_code)
                seen_codes.add(parent_code)
            # Description付きのAccount Codeを追加
            for desc_code in account_code_map[parent_code]:
                if desc_code not in seen_codes and '_' in desc_code:
                    ordered_account_codes.append(desc_code)
                    seen_codes.add(desc_code)
    
    # ヘッダー行を作成（マスターフォーマット#3の形式に合わせる）
    header_rows = []
    
    # 行1-7: マスターフォーマットのヘッダー構造を再現
    # 行1: レベル1 Account Code
    row1 = ['Date']  # Date列
    row2 = ['']  # レベル2
    row3 = ['']  # レベル3
    row4 = ['']  # Internal Code
    row5 = ['']  # Internal Account Name
    row6 = ['']  # KEIEI houkoku
    row7 = ['']  # 詳細
    row8 = ['Date']  # Description（列名行）
    
    # 各Account Codeに対してActual列とBudget列を追加
    for account_code in ordered_account_codes:
        # Account Codeの構造を取得
        if account_code in account_structure:
            info = account_structure[account_code]
            level = info['level']
            internal_code = info['internal_code']
            account_name = info['account_name']
            keiei_houkoku = info['keiei_houkoku']
            detail_jp = info['detail_jp']
            description = info['description']
        else:
            # Description付きのAccount Codeの場合
            if '_' in account_code:
                parent_code = account_code.split('_')[0]
                description = account_code.split('_', 1)[1]
                if parent_code in account_structure:
                    info = account_structure[parent_code]
                    level = info['level']
                    internal_code = info['internal_code']
                    account_name = info['account_name']
                    keiei_houkoku = info['keiei_houkoku']
                    detail_jp = info['detail_jp']
                else:
                    level = 3
                    internal_code = ''
                    account_name = ''
                    keiei_houkoku = ''
                    detail_jp = ''
            else:
                level = 3
                internal_code = ''
                account_name = ''
                keiei_houkoku = ''
                detail_jp = ''
                description = account_code
        
        # Actual列とBudget列を追加
        for col_type in ['Actual', 'Budget']:
            # 行1-7: ヘッダー情報
            if level == 1:
                row1.append(account_code)
                row2.append('')
                row3.append('')
            elif level == 2:
                row1.append('')
                row2.append(account_code)
                row3.append('')
            else:
                row1.append('')
                row2.append('')
                row3.append(account_code)
            
            row4.append(internal_code)
            row5.append(account_name)
            row6.append(keiei_houkoku)
            row7.append(detail_jp)
            
            # 行8: 列名（Description + Actual/Budget）
            if description:
                row8.append(f"{description} {col_type}")
            else:
                row8.append(f"{account_code} {col_type}")
    
    header_rows = [row1, row2, row3, row4, row5, row6, row7, row8]
    
    # データ行を作成
    data_rows = []
    
    for date_str in all_dates:
        data_row = [date_str.replace('-', ' ')]  # Date列（Apr-25形式）
        
        # 各Account Codeに対してActual列とBudget列を追加
        for account_code in ordered_account_codes:
            # Actual列のデータを取得
            actual_amount = 0
            # まず、直接マッチするキーを確認
            if (date_str, account_code) in actual_data:
                actual_amount = actual_data[(date_str, account_code)]
            else:
                # Description付きのAccount Codeのデータを集約（親コードの場合）
                if account_code not in account_code_map:
                    account_code_map[account_code] = []
                for desc_code in account_code_map.get(account_code, []):
                    if '_' in desc_code and (date_str, desc_code) in actual_data:
                        actual_amount += actual_data[(date_str, desc_code)]
            
            # Budget列のデータを取得
            budget_amount = 0
            # まず、直接マッチするキーを確認
            if (date_str, account_code) in budget_data:
                budget_amount = budget_data[(date_str, account_code)]
            else:
                # Description付きのAccount Codeのデータを集約（親コードの場合）
                for desc_code in account_code_map.get(account_code, []):
                    if '_' in desc_code and (date_str, desc_code) in budget_data:
                        budget_amount += budget_data[(date_str, desc_code)]
            
            data_row.append(actual_amount)
            data_row.append(budget_amount)
        
        data_rows.append(data_row)
    
    # CSVに出力
    with open(output_file, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        # ヘッダー行を書き込み
        for header_row in header_rows:
            writer.writerow(header_row)
        # データ行を書き込み
        writer.writerows(data_rows)
    
    print(f"\n✅ 処理完了！")
    print(f"出力ファイル: {output_file}")
    print(f"総Account Code数: {len(ordered_account_codes)}")
    print(f"総列数: {len(ordered_account_codes) * 2 + 1} (Date + Actual/Budget各{len(ordered_account_codes)})")
    print(f"データ行数: {len(data_rows)}行")

def main():
    """メイン処理"""
    master_file = '/Users/sugenotomohiro/Desktop/CURSOR/Manabox/20_DATA_Input/01.3consulting/Glory Format - #3 sample(DB for BI report).csv'
    actual_file = '/Users/sugenotomohiro/Desktop/CURSOR/Manabox/20_DATA_Input/01.3consulting/Glory　Format - #1MonthlyPL(Actual) .csv'
    budget_file = '/Users/sugenotomohiro/Desktop/CURSOR/Manabox/20_DATA_Input/01.3consulting/Glory　Format - #2Monthly PL(Budget).csv'
    output_file = '/Users/sugenotomohiro/Desktop/CURSOR/Manabox/20_DATA_Input/01.3consulting/月次DB_SharePoint用.csv'
    
    try:
        create_wide_format_db(master_file, actual_file, budget_file, output_file)
        print("\n次のステップ:")
        print("1. 出力されたCSVファイルをエクセルで開く")
        print("2. SharePointにアップロード")
        print("3. Looker Studioで接続")
        
    except Exception as e:
        print(f"エラーが発生しました: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    main()

