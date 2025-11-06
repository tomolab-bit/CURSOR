#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
会計コードマスターフォーマット作成スクリプト（最終版）
#1（実績）と#2（予算）を結合して、網羅的なマスターフォーマット#3を作成
"""

import csv
from collections import OrderedDict

def extract_account_code(row, start_col, end_col):
    """Account Code アカウントを抽出し、レベルも返す"""
    for i in range(start_col, min(end_col, len(row))):
        if row[i] and row[i].strip():
            return row[i].strip(), i - start_col
    return None, None

def parse_csv_complete(filename):
    """CSVファイルを完全に解析"""
    items_dict = OrderedDict()
    current_parent_1 = None  # #1の現在の親Account Code
    current_parent_2 = None  # #2の現在の親Account Code
    
    with open(filename, 'r', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        rows = list(reader)
        
        # ヘッダー行をスキップ（1-3行目）
        for row_num, row in enumerate(rows[3:], start=4):
            # #1のデータ（列0-8）
            row_1 = row[0:9] if len(row) > 9 else row[0:] + [''] * (9 - len(row))
            # #2のデータ（列10-18）
            row_2 = row[10:19] if len(row) > 19 else ([''] * (19 - len(row)) + row[10:]) if len(row) > 10 else [''] * 9
            # Đơn giá（列19）
            don_gia = row[19].strip() if len(row) > 19 and row[19] else ''
            
            # Account Codeを抽出
            acc_code_1, level_1 = extract_account_code(row_1, 0, 3)
            acc_code_2, level_2 = extract_account_code(row_2, 0, 3)
            
            # #1の処理
            if acc_code_1:
                current_parent_1 = acc_code_1
                if acc_code_1 not in items_dict:
                    items_dict[acc_code_1] = {
                        'account_code': acc_code_1,
                        'level': level_1,
                        'data_1': None,
                        'data_2': None,
                        'don_gia': None,
                        'sub_items': []
                    }
                items_dict[acc_code_1]['data_1'] = row_1
                items_dict[acc_code_1]['level'] = level_1 if level_1 is not None else items_dict[acc_code_1]['level']
            elif current_parent_1 and any(cell.strip() for cell in row_1):
                # Account Codeがないがデータがある行はサブアイテム
                items_dict[current_parent_1]['sub_items'].append({
                    'data_1': row_1,
                    'data_2': None,
                    'don_gia': None
                })
            
            # #2の処理
            if acc_code_2:
                current_parent_2 = acc_code_2
                if acc_code_2 not in items_dict:
                    items_dict[acc_code_2] = {
                        'account_code': acc_code_2,
                        'level': level_2,
                        'data_1': None,
                        'data_2': None,
                        'don_gia': None,
                        'sub_items': []
                    }
                items_dict[acc_code_2]['data_2'] = row_2
                items_dict[acc_code_2]['don_gia'] = don_gia if don_gia else items_dict[acc_code_2]['don_gia']
                items_dict[acc_code_2]['level'] = level_2 if level_2 is not None else items_dict[acc_code_2]['level']
            elif current_parent_2 and (any(cell.strip() for cell in row_2) or don_gia):
                # Account Codeがないがデータがある行はサブアイテム
                # 既存のサブアイテムを更新するか、新しいサブアイテムを追加
                found = False
                for sub_item in items_dict[current_parent_2]['sub_items']:
                    if not sub_item['data_2']:
                        sub_item['data_2'] = row_2
                        sub_item['don_gia'] = don_gia if don_gia else sub_item['don_gia']
                        found = True
                        break
                if not found:
                    items_dict[current_parent_2]['sub_items'].append({
                        'data_1': None,
                        'data_2': row_2,
                        'don_gia': don_gia
                    })
    
    return items_dict

def create_master_csv_final(items_dict, output_filename):
    """マスターフォーマット#3をCSVとして出力（最終版）"""
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
        sorted_items = sorted(items_dict.items(), key=lambda x: (
            x[1]['level'] if x[1]['level'] is not None else 999,
            x[0]
        ))
        
        for acc_code, item in sorted_items:
            # メインデータ行
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
            
            # Đơn giá
            row[9] = item['don_gia'] if item['don_gia'] else ''
            
            # メタデータ
            has_1 = item['data_1'] is not None
            has_2 = item['data_2'] is not None
            if has_1 and has_2:
                source = 'Both'
            elif has_1:
                source = '#1'
            else:
                source = '#2'
            
            row[10] = source
            row[11] = str(level)
            row[12] = 'Yes' if has_1 else 'No'
            row[13] = 'Yes' if has_2 else 'No'
            
            writer.writerow(row)
            
            # サブアイテムを追加
            for sub_item in item['sub_items']:
                sub_row = [''] * 14
                sub_row[level + 1] = ''  # Account Codeは空（親に紐づく）
                
                # サブアイテムのデータを統合
                sub_data = sub_item['data_2'] if sub_item['data_2'] else sub_item['data_1']
                if sub_data:
                    if len(sub_data) > 3:
                        sub_row[3] = sub_data[3] if len(sub_data) > 3 else ''
                        sub_row[4] = sub_data[4] if len(sub_data) > 4 else ''
                        sub_row[5] = sub_data[5] if len(sub_data) > 5 else ''
                        sub_row[6] = sub_data[6] if len(sub_data) > 6 else ''
                        sub_row[7] = sub_data[7] if len(sub_data) > 7 else ''
                        sub_row[8] = sub_data[8] if len(sub_data) > 8 else ''
                
                sub_row[9] = sub_item['don_gia'] if sub_item['don_gia'] else ''
                
                sub_has_1 = sub_item['data_1'] is not None
                sub_has_2 = sub_item['data_2'] is not None
                if sub_has_1 and sub_has_2:
                    sub_source = 'Both'
                elif sub_has_1:
                    sub_source = '#1'
                else:
                    sub_source = '#2'
                
                sub_row[10] = sub_source
                sub_row[11] = str(level + 1)
                sub_row[12] = 'Yes' if sub_has_1 else 'No'
                sub_row[13] = 'Yes' if sub_has_2 else 'No'
                
                writer.writerow(sub_row)

def main():
    input_file = '/Users/sugenotomohiro/Desktop/CURSOR/Manabox/20_DATA_Input/01.3consulting/Glory　Format_Source Data.csv'
    output_file = '/Users/sugenotomohiro/Desktop/CURSOR/Manabox/20_DATA_Input/01.3consulting/マスターフォーマット#3_最終版.csv'
    
    print("CSVファイルを完全解析中...")
    items_dict = parse_csv_complete(input_file)
    
    print(f"マスター項目数: {len(items_dict)}件")
    total_sub_items = sum(len(item['sub_items']) for item in items_dict.values())
    print(f"サブアイテム総数: {total_sub_items}件")
    
    print("マスターフォーマット#3を作成中...")
    create_master_csv_final(items_dict, output_file)
    
    print(f"完了: {output_file}")
    
    # 統計情報
    both_count = sum(1 for item in items_dict.values() if item['data_1'] and item['data_2'])
    only_1 = sum(1 for item in items_dict.values() if item['data_1'] and not item['data_2'])
    only_2 = sum(1 for item in items_dict.values() if not item['data_1'] and item['data_2'])
    print(f"\n統計:")
    print(f"  #1と#2の両方: {both_count}件")
    print(f"  #1のみ: {only_1}件")
    print(f"  #2のみ: {only_2}件")

if __name__ == '__main__':
    main()

