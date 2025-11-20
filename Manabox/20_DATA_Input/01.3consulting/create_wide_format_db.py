#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
#1（実績）と#2（予算）のCSVを結合して#3形式（横型、ワイド形式）の月次DBを作成
Looker Studio用：各Account Codeごとに1列、Actual/Budgetは行で分ける
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

def extract_account_structure_from_data(actual_file, budget_file):
    """#1と#2のデータからAccount Code構造を抽出"""
    print(f"Account Code構造を抽出中...")
    
    account_structure = OrderedDict()  # {account_code: {level, internal_code, account_name, ...}}
    account_code_map = {}  # {parent_code: [description付きキーのリスト]}
    account_code_order = []  # Account Codeの出現順序を保持
    
    # #1（実績）から構造を抽出
    with open(actual_file, 'r', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        rows = list(reader)
    
    data_rows = rows[2:] if len(rows) > 2 else []
    current_account_code = None
    current_level = 0
    
    for row in data_rows:
        if not row or len(row) == 0:
            continue
        
        account_code = None
        level = 0
        
        # Account Codeを取得（列0, 1, 2を確認）
        for col_idx in range(3):
            if col_idx < len(row) and row[col_idx] and row[col_idx].strip():
                account_code = row[col_idx].strip()
                level = col_idx + 1
                break
        
        description_chi_tieu = row[8] if len(row) > 8 else ''
        internal_code = row[3] if len(row) > 3 else ''
        account_name = row[4] if len(row) > 4 else ''
        keiei_houkoku = row[5] if len(row) > 5 else ''
        detail_jp = row[6] if len(row) > 6 else ''
        description_en = row[7] if len(row) > 7 else ''
        
        if account_code:
            current_account_code = account_code
            current_level = level
            
            if not description_chi_tieu:
                description_chi_tieu = description_en
            
            if account_code not in account_structure:
                account_structure[account_code] = {
                    'level': level,
                    'internal_code': internal_code,
                    'account_name': account_name,
                    'keiei_houkoku': keiei_houkoku,
                    'detail_jp': detail_jp,
                    'description_en': description_en,
                    'description_chi_tieu': description_chi_tieu
                }
                if account_code not in account_code_order:
                    account_code_order.append(account_code)
        
        # Description付きのAccount Code（サブアイテム）を処理
        if current_account_code and description_chi_tieu and description_chi_tieu.strip():
            # 金額があるか確認
            has_amount = False
            for col_idx in range(9, len(row)):
                if col_idx < len(row):
                    amount = parse_number(row[col_idx])
                    if amount != 0:
                        has_amount = True
                        break
            
            if has_amount:
                unique_key = f"{current_account_code}_{description_chi_tieu.strip()}"
                if current_account_code not in account_code_map:
                    account_code_map[current_account_code] = []
                if unique_key not in account_code_map[current_account_code]:
                    account_code_map[current_account_code].append(unique_key)
                
                if unique_key not in account_structure:
                    parent_info = account_structure.get(current_account_code, {})
                    account_structure[unique_key] = {
                        'level': parent_info.get('level', current_level),
                        'internal_code': parent_info.get('internal_code', ''),
                        'account_name': parent_info.get('account_name', ''),
                        'keiei_houkoku': parent_info.get('keiei_houkoku', ''),
                        'detail_jp': parent_info.get('detail_jp', ''),
                        'description_en': description_chi_tieu.strip(),
                        'description_chi_tieu': description_chi_tieu.strip(),
                        'parent_account_code': current_account_code
                    }
    
    # #2（予算）からも構造を抽出（不足しているAccount Codeを追加）
    with open(budget_file, 'r', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        rows = list(reader)
    
    data_rows = rows[7:] if len(rows) > 7 else []
    current_account_code = None
    current_level = 0
    
    for row in data_rows:
        if not row or len(row) == 0:
            continue
        
        account_code = None
        level = 0
        
        for col_idx in range(3):
            if col_idx < len(row) and row[col_idx] and row[col_idx].strip():
                account_code = row[col_idx].strip()
                level = col_idx + 1
                break
        
        description_chi_tieu = row[8] if len(row) > 8 else ''
        internal_code = row[3] if len(row) > 3 else ''
        account_name = row[4] if len(row) > 4 else ''
        keiei_houkoku = row[5] if len(row) > 5 else ''
        detail_jp = row[6] if len(row) > 6 else ''
        description_en = row[7] if len(row) > 7 else ''
        
        if account_code:
            current_account_code = account_code
            current_level = level
            
            if not description_chi_tieu:
                description_chi_tieu = description_en
            
            if account_code not in account_structure:
                account_structure[account_code] = {
                    'level': level,
                    'internal_code': internal_code,
                    'account_name': account_name,
                    'keiei_houkoku': keiei_houkoku,
                    'detail_jp': detail_jp,
                    'description_en': description_en,
                    'description_chi_tieu': description_chi_tieu
                }
                if account_code not in account_code_order:
                    account_code_order.append(account_code)
        
        # Description付きのAccount Codeを処理
        if current_account_code and description_chi_tieu and description_chi_tieu.strip():
            has_amount = False
            for col_idx in range(10, len(row)):  # 予算は列10から
                if col_idx < len(row):
                    amount = parse_number(row[col_idx])
                    if amount != 0:
                        has_amount = True
                        break
            
            if has_amount:
                unique_key = f"{current_account_code}_{description_chi_tieu.strip()}"
                if current_account_code not in account_code_map:
                    account_code_map[current_account_code] = []
                if unique_key not in account_code_map[current_account_code]:
                    account_code_map[current_account_code].append(unique_key)
                
                if unique_key not in account_structure:
                    parent_info = account_structure.get(current_account_code, {})
                    account_structure[unique_key] = {
                        'level': parent_info.get('level', current_level),
                        'internal_code': parent_info.get('internal_code', ''),
                        'account_name': parent_info.get('account_name', ''),
                        'keiei_houkoku': parent_info.get('keiei_houkoku', ''),
                        'detail_jp': parent_info.get('detail_jp', ''),
                        'description_en': description_chi_tieu.strip(),
                        'description_chi_tieu': description_chi_tieu.strip(),
                        'parent_account_code': current_account_code
                    }
    
    print(f"Account Code構造抽出完了: {len(account_structure)}件")
    return account_structure, account_code_map, account_code_order

def load_actual_data(actual_file):
    """#1実績データを読み込む"""
    print(f"実績データを読み込み中: {actual_file}")
    
    actual_data = {}  # {(date_str, account_code): amount}
    
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
        
        account_code = None
        for col_idx in range(3):
            if col_idx < len(row) and row[col_idx] and row[col_idx].strip():
                account_code = row[col_idx].strip()
                break
        
        if account_code:
            current_account_code = account_code
        
        description_chi_tieu = row[8] if len(row) > 8 else ''
        
        if current_account_code:
            has_amount = False
            for col_idx, month_header in month_columns.items():
                if col_idx < len(row):
                    amount = parse_number(row[col_idx])
                    if amount != 0:
                        has_amount = True
            
            if account_code or has_amount:
                if not description_chi_tieu:
                    description_chi_tieu = row[7] if len(row) > 7 else ''
                
                if description_chi_tieu and description_chi_tieu.strip():
                    unique_key = f"{current_account_code}_{description_chi_tieu.strip()}"
                else:
                    unique_key = current_account_code
                
                for col_idx, month_header in month_columns.items():
                    if col_idx < len(row):
                        amount = parse_number(row[col_idx])
                        date_str = convert_date_format(month_header)
                        key = (date_str, unique_key)
                        actual_data[key] = amount
    
    print(f"実績データ読み込み完了: {len(actual_data)}件のレコード")
    return actual_data

def load_budget_data(budget_file):
    """#2予算データを読み込む"""
    print(f"予算データを読み込み中: {budget_file}")
    
    budget_data = {}  # {(date_str, account_code): amount}
    
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
        
        account_code = None
        for col_idx in range(3):
            if col_idx < len(row) and row[col_idx] and row[col_idx].strip():
                account_code = row[col_idx].strip()
                break
        
        if account_code:
            current_account_code = account_code
        
        description_chi_tieu = row[8] if len(row) > 8 else ''
        
        if current_account_code:
            has_amount = False
            for col_idx, month_header in month_columns.items():
                if col_idx < len(row):
                    amount = parse_number(row[col_idx])
                    if amount != 0:
                        has_amount = True
            
            if account_code or has_amount:
                if not description_chi_tieu:
                    description_chi_tieu = row[7] if len(row) > 7 else ''
                
                if description_chi_tieu and description_chi_tieu.strip():
                    unique_key = f"{current_account_code}_{description_chi_tieu.strip()}"
                else:
                    unique_key = current_account_code
                
                for col_idx, month_header in month_columns.items():
                    if col_idx < len(row):
                        amount = parse_number(row[col_idx])
                        date_str = convert_date_format(month_header)
                        key = (date_str, unique_key)
                        budget_data[key] = amount
    
    print(f"予算データ読み込み完了: {len(budget_data)}件のレコード")
    return budget_data

def get_column_name(account_code, account_structure, all_column_names):
    """列名を取得（Account Codeを含めて一意にする、Looker Studio用にシンプルに）"""
    info = account_structure.get(account_code, {})
    
    # Description付きのAccount Codeの場合
    if '_' in account_code:
        # Account CodeとDescriptionを組み合わせ（カンマを削除）
        parent_code = account_code.split('_')[0]
        description = account_code.split('_', 1)[1].replace(',', '').replace('"', '')
        column_name = f"{parent_code}_{description}"
        
        # 重複チェック
        if column_name in all_column_names:
            return account_code
        return column_name
    
    # Account Codeを基本として使用
    base_code = account_code
    
    # Descriptionを取得（優先順位：description_en > description_chi_tieu）
    description_en = info.get('description_en', '')
    description_chi_tieu = info.get('description_chi_tieu', '')
    
    # シンプルな列名を作成（Account Code + Description、カンマを削除）
    if description_en:
        description_en = description_en.replace(',', '').replace('"', '')
        column_name = f"{base_code}_{description_en}"
        if column_name in all_column_names:
            return base_code
        return column_name
    elif description_chi_tieu:
        description_chi_tieu = description_chi_tieu.replace(',', '').replace('"', '')
        column_name = f"{base_code}_{description_chi_tieu}"
        if column_name in all_column_names:
            return base_code
        return column_name
    else:
        return account_code

def create_wide_format_db(actual_file, budget_file, output_file):
    """横型（ワイド形式）のDBを作成"""
    
    # Account Code構造を抽出
    account_structure, account_code_map, account_code_order = extract_account_structure_from_data(actual_file, budget_file)
    
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
    
    # Account Codeを元の出現順序で整理（階層構造を保持）
    ordered_account_codes = []
    seen_codes = set()
    
    # 元の順序を保持しながら階層構造に従って並べる
    for code in account_code_order:
        if code in seen_codes:
            continue
        
        info = account_structure.get(code, {})
        level = info.get('level', 3)
        
        # 親コードを先に追加
        if level == 1:
            ordered_account_codes.append(code)
            seen_codes.add(code)
            
            # 子コードを追加（元の順序を保持）
            for code2 in account_code_order:
                if code2 in seen_codes:
                    continue
                info2 = account_structure.get(code2, {})
                level2 = info2.get('level', 3)
                if level2 == 2 and code2.startswith(code):
                    ordered_account_codes.append(code2)
                    seen_codes.add(code2)
                    
                    # レベル3の子コードを追加
                    for code3 in account_code_order:
                        if code3 in seen_codes:
                            continue
                        info3 = account_structure.get(code3, {})
                        level3 = info3.get('level', 3)
                        if level3 == 3 and code3.startswith(code2):
                            ordered_account_codes.append(code3)
                            seen_codes.add(code3)
                            
                            # Description付きのコードを追加
                            for desc_code in account_code_map.get(code3, []):
                                if desc_code not in seen_codes:
                                    ordered_account_codes.append(desc_code)
                                    seen_codes.add(desc_code)
        elif level == 2:
            # 親コードが既に追加されているか確認
            parent_added = False
            for parent_code in account_code_order:
                if parent_code.startswith(code[:1]) and account_structure.get(parent_code, {}).get('level') == 1:
                    if parent_code in seen_codes:
                        parent_added = True
                        break
            
            if not parent_added:
                ordered_account_codes.append(code)
                seen_codes.add(code)
            else:
                ordered_account_codes.append(code)
                seen_codes.add(code)
                
                # レベル3の子コードを追加
                for code3 in account_code_order:
                    if code3 in seen_codes:
                        continue
                    info3 = account_structure.get(code3, {})
                    level3 = info3.get('level', 3)
                    if level3 == 3 and code3.startswith(code):
                        ordered_account_codes.append(code3)
                        seen_codes.add(code3)
                        
                        # Description付きのコードを追加
                        for desc_code in account_code_map.get(code3, []):
                            if desc_code not in seen_codes:
                                ordered_account_codes.append(desc_code)
                                seen_codes.add(desc_code)
        else:
            # レベル3のコード
            if code not in seen_codes:
                ordered_account_codes.append(code)
                seen_codes.add(code)
                
                # Description付きのコードを追加
                for desc_code in account_code_map.get(code, []):
                    if desc_code not in seen_codes:
                        ordered_account_codes.append(desc_code)
                        seen_codes.add(desc_code)
    
    # 残りのコードを追加（元の順序を保持）
    for code in account_code_order:
        if code not in seen_codes:
            ordered_account_codes.append(code)
            seen_codes.add(code)
            
            # Description付きのコードを追加
            if code in account_code_map:
                for desc_code in account_code_map[code]:
                    if desc_code not in seen_codes:
                        ordered_account_codes.append(desc_code)
                        seen_codes.add(desc_code)
    
    # ヘッダー行を作成（#3サンプルフォーマットに合わせる）
    row1 = ['']  # レベル1 Account Code
    row2 = ['']  # レベル2
    row3 = ['']  # レベル3
    row4 = ['']  # Internal Code
    row5 = ['']  # Internal Account Name
    row6 = ['']  # KEIEI houkoku
    row7 = ['']  # 詳細
    row8 = ['Date', '予算/実績']  # Description（列名行）- Date列とType列を追加
    all_column_names = set(['Date', '予算/実績'])  # 重複チェック用
    
    # 各Account Codeに対して1列ずつ追加
    for account_code in ordered_account_codes:
        info = account_structure.get(account_code, {})
        level = info.get('level', 3)
        internal_code = info.get('internal_code', '')
        account_name = info.get('account_name', '')
        keiei_houkoku = info.get('keiei_houkoku', '')
        detail_jp = info.get('detail_jp', '')
        
        # Description付きのAccount Codeの場合
        if '_' in account_code:
            parent_code = account_code.split('_')[0]
            if parent_code in account_structure:
                parent_info = account_structure[parent_code]
                level = parent_info.get('level', 3)
                internal_code = parent_info.get('internal_code', '')
                account_name = parent_info.get('account_name', '')
                keiei_houkoku = parent_info.get('keiei_houkoku', '')
                detail_jp = parent_info.get('detail_jp', '')
        
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
        
        # 行8: 列名（Account Codeを含めて一意にする）
        column_name = get_column_name(account_code, account_structure, all_column_names)
        # まだ重複している場合はAccount Codeのみ
        if column_name in all_column_names:
            column_name = account_code
        all_column_names.add(column_name)
        row8.append(column_name)
    
    header_rows = [row1, row2, row3, row4, row5, row6, row7, row8]
    
    # データ行を作成（各期間に対してActual行とBudget行を作成）
    data_rows = []
    
    for date_str in all_dates:
        date_display = date_str.replace('-', ' ')  # 2025-04 -> 2025 04
        
        # Actual行を作成
        actual_row = [date_display, 'Actual']  # Date列とType列
        
        # 各Account Codeに対してActual列のデータを追加
        for account_code in ordered_account_codes:
            actual_amount = 0
            # まず、直接マッチするキーを確認
            if (date_str, account_code) in actual_data:
                actual_amount = actual_data[(date_str, account_code)]
            else:
                # Description付きのAccount Codeのデータを集約（親コードの場合）
                for desc_code in account_code_map.get(account_code, []):
                    if '_' in desc_code and (date_str, desc_code) in actual_data:
                        actual_amount += actual_data[(date_str, desc_code)]
            
            actual_row.append(actual_amount)
        
        data_rows.append(actual_row)
        
        # Budget行を作成
        budget_row = [date_display, 'Budget']  # Date列とType列
        
        # 各Account Codeに対してBudget列のデータを追加
        for account_code in ordered_account_codes:
            budget_amount = 0
            # まず、直接マッチするキーを確認
            if (date_str, account_code) in budget_data:
                budget_amount = budget_data[(date_str, account_code)]
            else:
                # Description付きのAccount Codeのデータを集約（親コードの場合）
                for desc_code in account_code_map.get(account_code, []):
                    if '_' in desc_code and (date_str, desc_code) in budget_data:
                        budget_amount += budget_data[(date_str, desc_code)]
            
            budget_row.append(budget_amount)
        
        data_rows.append(budget_row)
    
    # CSVに出力（Looker Studio用：列名行のみ）
    with open(output_file, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f, quoting=csv.QUOTE_MINIMAL)
        # 列名行のみを書き込み（行1-7のヘッダーは削除してLooker Studioで使いやすく）
        writer.writerow(row8)
        # データ行を書き込み
        writer.writerows(data_rows)
    
    print(f"\n✅ 処理完了！")
    print(f"出力ファイル: {output_file}")
    print(f"総Account Code数: {len(ordered_account_codes)}")
    print(f"総列数: {len(ordered_account_codes) + 2} (Date + Type + Account Codes各{len(ordered_account_codes)})")
    print(f"データ行数: {len(data_rows)}行（各期間に対してActual行とBudget行）")

def main():
    """メイン処理"""
    actual_file = '/Users/sugenotomohiro/Desktop/CURSOR/Manabox/20_DATA_Input/01.3consulting/Glory　Format - #1MonthlyPL(Actual) .csv'
    budget_file = '/Users/sugenotomohiro/Desktop/CURSOR/Manabox/20_DATA_Input/01.3consulting/Glory　Format - #2Monthly PL(Budget).csv'
    output_file = '/Users/sugenotomohiro/Desktop/CURSOR/Manabox/20_DATA_Input/01.3consulting/月次DB_SharePoint用.csv'
    
    try:
        create_wide_format_db(actual_file, budget_file, output_file)
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
