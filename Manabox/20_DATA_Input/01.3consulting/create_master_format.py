#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
会計コードマスターフォーマット作成スクリプト
#1（実績）と#2（予算）を結合して、網羅的なマスターフォーマット#3を作成
"""

import csv
import re
from collections import OrderedDict

def extract_account_code(row, start_col, end_col):
    """Account Code アカウントを抽出"""
    for i in range(start_col, end_col):
        if row[i] and row[i].strip():
            return row[i].strip()
    return None

def get_level(row, start_col, end_col):
    """階層レベルを取得（左から何列目にAccount Codeがあるか）"""
    for i in range(start_col, end_col):
        if row[i] and row[i].strip():
            return i - start_col
    return None

def normalize_account_code(code):
    """Account Codeを正規化（空白削除など）"""
    if not code:
        return None
    code = str(code).strip()
    if code == '':
        return None
    return code

def parse_csv(filename):
    """CSVファイルを解析して#1と#2のデータを抽出"""
    data_1 = []  # 実績データ
    data_2 = []  # 予算データ
    
    with open(filename, 'r', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        rows = list(reader)
        
        # ヘッダー行をスキップ（1-3行目）
        for row_num, row in enumerate(rows[3:], start=4):
            # #1のデータ（列0-8）
            row_1 = row[0:9] if len(row) > 9 else row[0:] + [''] * (9 - len(row))
            # #2のデータ（列10-19）
            row_2 = row[10:19] if len(row) > 19 else ([''] * (19 - len(row)) + row[10:]) if len(row) > 10 else [''] * 9
            
            # Account Codeを抽出
            acc_code_1 = extract_account_code(row_1, 0, 3)
            acc_code_2 = extract_account_code(row_2, 0, 3)
            
            # レベルを取得
            level_1 = get_level(row_1, 0, 3)
            level_2 = get_level(row_2, 0, 3)
            
            # データが存在する場合のみ追加
            if any(cell.strip() for cell in row_1):
                data_1.append({
                    'row_num': row_num,
                    'account_code': acc_code_1,
                    'level': level_1,
                    'data': row_1,
                    'full_row': row
                })
            
            if any(cell.strip() for cell in row_2):
                data_2.append({
                    'row_num': row_num,
                    'account_code': acc_code_2,
                    'level': level_2,
                    'data': row_2,
                    'full_row': row
                })
    
    return data_1, data_2

def merge_data(data_1, data_2):
    """#1と#2を結合してマスターフォーマット#3を作成"""
    master = OrderedDict()
    
    # Account Codeをキーとして、両方のデータを統合
    # まず#1のデータを処理
    for item in data_1:
        acc_code = normalize_account_code(item['account_code'])
        if acc_code:
            if acc_code not in master:
                master[acc_code] = {
                    'account_code': acc_code,
                    'level': item['level'],
                    'data_1': item['data'],
                    'data_2': None,
                    'has_data_1': True,
                    'has_data_2': False,
                    'row_num_1': item['row_num'],
                    'row_num_2': None
                }
            else:
                # 既存のデータを更新（より詳細な情報を優先）
                master[acc_code]['data_1'] = item['data']
                master[acc_code]['row_num_1'] = item['row_num']
    
    # #2のデータを処理
    for item in data_2:
        acc_code = normalize_account_code(item['account_code'])
        if acc_code:
            if acc_code not in master:
                # #2のみに存在する項目
                master[acc_code] = {
                    'account_code': acc_code,
                    'level': item['level'],
                    'data_1': None,
                    'data_2': item['data'],
                    'has_data_1': False,
                    'has_data_2': True,
                    'row_num_1': None,
                    'row_num_2': item['row_num']
                }
            else:
                # #1にも存在する項目
                master[acc_code]['data_2'] = item['data']
                master[acc_code]['has_data_2'] = True
                master[acc_code]['row_num_2'] = item['row_num']
                # レベルが異なる場合は#2を優先（より詳細な可能性）
                if item['level'] is not None:
                    master[acc_code]['level'] = item['level']
    
    return master

def get_account_code_from_row(row, start_idx=0):
    """行からAccount Codeを抽出（階層構造を考慮）"""
    for i in range(start_idx, min(start_idx + 3, len(row))):
        code = normalize_account_code(row[i])
        if code:
            return code, i - start_idx
    return None, None

def process_sub_items(data_1, data_2, master):
    """Account Codeがなく名前だけの行を処理（親に紐づける）"""
    sub_items = []
    
    # #1のサブアイテムを処理
    for i, item in enumerate(data_1):
        acc_code = normalize_account_code(item['account_code'])
        if not acc_code:
            # Account Codeがない行は、前の行のAccount Codeに紐づける
            if i > 0:
                prev_item = data_1[i-1]
                prev_code = normalize_account_code(prev_item['account_code'])
                if prev_code:
                    sub_items.append({
                        'parent_code': prev_code,
                        'source': 'data_1',
                        'data': item['data'],
                        'row_num': item['row_num']
                    })
    
    # #2のサブアイテムを処理
    for i, item in enumerate(data_2):
        acc_code = normalize_account_code(item['account_code'])
        if not acc_code:
            if i > 0:
                prev_item = data_2[i-1]
                prev_code = normalize_account_code(prev_item['account_code'])
                if prev_code:
                    sub_items.append({
                        'parent_code': prev_code,
                        'source': 'data_2',
                        'data': item['data'],
                        'row_num': item['row_num']
                    })
    
    return sub_items

def create_master_csv(master, sub_items, output_filename):
    """マスターフォーマット#3をCSVとして出力"""
    header = [
        'Account Code アカウント',
        '',
        '',
        'Internal code',
        'Internal Account Name',
        'KEIEI houkoku',
        '詳細',
        'Description',
        'Chỉ tiêu',
        'Đơn giá',
        'Source',
        'Level',
        'Has Data #1',
        'Has Data #2'
    ]
    
    with open(output_filename, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(header)
        
        # Account Code順にソート（階層構造を維持）
        sorted_items = sorted(master.items(), key=lambda x: (
            x[1]['level'] if x[1]['level'] is not None else 999,
            x[0]
        ))
        
        for acc_code, item in sorted_items:
            row = [''] * 14
            level = item['level'] if item['level'] is not None else 0
            
            # Account Codeを階層に応じて配置
            row[level] = acc_code
            
            # データを統合（#2を優先、なければ#1）
            data = item['data_2'] if item['data_2'] else item['data_1']
            if data:
                # Internal code
                if len(data) > 3 and data[3]:
                    row[3] = data[3]
                # Internal Account Name
                if len(data) > 4 and data[4]:
                    row[4] = data[4]
                # KEIEI houkoku
                if len(data) > 5 and data[5]:
                    row[5] = data[5]
                # 詳細
                if len(data) > 6 and data[6]:
                    row[6] = data[6]
                # Description
                if len(data) > 7 and data[7]:
                    row[7] = data[7]
                # Chỉ tiêu
                if len(data) > 8 and data[8]:
                    row[8] = data[8]
            
            # Đơn giá（#2のデータから取得）
            if item['data_2'] and len(item['data_2']) > 8:
                # #2の元データからĐơn giá列を取得（元のCSVの列19）
                pass  # 元のCSVから直接取得する必要がある
            
            # メタデータ
            row[9] = ''  # Đơn giáは後で追加
            row[10] = 'Both' if item['has_data_1'] and item['has_data_2'] else ('#1' if item['has_data_1'] else '#2')
            row[11] = str(level)
            row[12] = 'Yes' if item['has_data_1'] else 'No'
            row[13] = 'Yes' if item['has_data_2'] else 'No'
            
            writer.writerow(row)
            
            # サブアイテムを追加（同じ親に紐づくもの）
            for sub_item in sub_items:
                if sub_item['parent_code'] == acc_code:
                    sub_row = [''] * 14
                    sub_row[level + 1] = ''  # Account Codeは空
                    sub_data = sub_item['data']
                    if len(sub_data) > 3:
                        sub_row[3] = sub_data[3] if len(sub_data) > 3 else ''
                        sub_row[4] = sub_data[4] if len(sub_data) > 4 else ''
                        sub_row[5] = sub_data[5] if len(sub_data) > 5 else ''
                        sub_row[6] = sub_data[6] if len(sub_data) > 6 else ''
                        sub_row[7] = sub_data[7] if len(sub_data) > 7 else ''
                        sub_row[8] = sub_data[8] if len(sub_data) > 8 else ''
                    sub_row[10] = sub_item['source']
                    writer.writerow(sub_row)

def main():
    input_file = '/Users/sugenotomohiro/Desktop/CURSOR/Manabox/20_DATA_Input/01.3consulting/Glory　Format - Compara.csv'
    output_file = '/Users/sugenotomohiro/Desktop/CURSOR/Manabox/20_DATA_Input/01.3consulting/マスターフォーマット#3.csv'
    
    print("CSVファイルを解析中...")
    data_1, data_2 = parse_csv(input_file)
    
    print(f"#1（実績）: {len(data_1)}件")
    print(f"#2（予算）: {len(data_2)}件")
    
    print("データを結合中...")
    master = merge_data(data_1, data_2)
    
    print(f"マスター項目数: {len(master)}件")
    
    print("サブアイテムを処理中...")
    sub_items = process_sub_items(data_1, data_2, master)
    
    print(f"サブアイテム数: {len(sub_items)}件")
    
    print("マスターフォーマット#3を作成中...")
    create_master_csv(master, sub_items, output_file)
    
    print(f"完了: {output_file}")

if __name__ == '__main__':
    main()

