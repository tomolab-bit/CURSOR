#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
会計コードマスターフォーマット作成スクリプト（改良版）
#1（実績）と#2（予算）を結合して、網羅的なマスターフォーマット#3を作成
"""

import csv
from collections import OrderedDict

class AccountItem:
    """会計項目を表すクラス"""
    def __init__(self, account_code, level, parent_code=None):
        self.account_code = account_code
        self.level = level
        self.parent_code = parent_code
        self.data_1 = None  # 実績データ
        self.data_2 = None  # 予算データ
        self.don_gia = None  # Đơn giá
        self.sub_items = []  # Account Codeがないサブアイテム
        
    def has_data_1(self):
        return self.data_1 is not None
    
    def has_data_2(self):
        return self.data_2 is not None
    
    def get_source(self):
        if self.has_data_1() and self.has_data_2():
            return 'Both'
        elif self.has_data_1():
            return '#1'
        else:
            return '#2'

def extract_account_code(row, start_col, end_col):
    """Account Code アカウントを抽出"""
    for i in range(start_col, min(end_col, len(row))):
        if row[i] and row[i].strip():
            return row[i].strip(), i - start_col
    return None, None

def parse_csv_detailed(filename):
    """CSVファイルを詳細に解析"""
    items_dict = OrderedDict()
    sub_items_list = []  # Account Codeがない行
    
    with open(filename, 'r', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        rows = list(reader)
        
        # ヘッダー行をスキップ（1-3行目）
        for row_num, row in enumerate(rows[3:], start=4):
            # #1のデータ（列0-8）
            row_1 = row[0:9] if len(row) > 9 else row[0:] + [''] * (9 - len(row))
            # #2のデータ（列10-19、ただし列19はĐơn giá）
            row_2 = row[10:19] if len(row) > 19 else ([''] * (19 - len(row)) + row[10:]) if len(row) > 10 else [''] * 9
            don_gia_2 = row[19] if len(row) > 19 else ''
            
            # Account Codeを抽出
            acc_code_1, level_1 = extract_account_code(row_1, 0, 3)
            acc_code_2, level_2 = extract_account_code(row_2, 0, 3)
            
            # Account Codeがない行はサブアイテムとして処理
            if not acc_code_1 and not acc_code_2:
                # 前の行のAccount Codeを親として使用
                if sub_items_list:
                    last_item = sub_items_list[-1]
                    if last_item.get('parent_code'):
                        sub_items_list.append({
                            'parent_code': last_item['parent_code'],
                            'data_1': row_1 if any(cell.strip() for cell in row_1) else None,
                            'data_2': row_2 if any(cell.strip() for cell in row_2) else None,
                            'don_gia': don_gia_2,
                            'row_num': row_num
                        })
                continue
            
            # #1の処理
            if acc_code_1:
                if acc_code_1 not in items_dict:
                    items_dict[acc_code_1] = AccountItem(acc_code_1, level_1)
                items_dict[acc_code_1].data_1 = row_1
                items_dict[acc_code_1].level = level_1 if level_1 is not None else items_dict[acc_code_1].level
            
            # #2の処理
            if acc_code_2:
                if acc_code_2 not in items_dict:
                    items_dict[acc_code_2] = AccountItem(acc_code_2, level_2)
                items_dict[acc_code_2].data_2 = row_2
                items_dict[acc_code_2].don_gia = don_gia_2 if don_gia_2 else items_dict[acc_code_2].don_gia
                items_dict[acc_code_2].level = level_2 if level_2 is not None else items_dict[acc_code_2].level
            
            # Account Codeがない行の処理（前のAccount Codeを親として使用）
            if not acc_code_1 and acc_code_2:
                # #2にのみAccount Codeがある場合
                if acc_code_2 in items_dict:
                    items_dict[acc_code_2].sub_items.append({
                        'data_1': row_1 if any(cell.strip() for cell in row_1) else None,
                        'data_2': row_2 if any(cell.strip() for cell in row_2) else None,
                        'don_gia': don_gia_2
                    })
            elif acc_code_1 and not acc_code_2:
                # #1にのみAccount Codeがある場合
                if acc_code_1 in items_dict:
                    items_dict[acc_code_1].sub_items.append({
                        'data_1': row_1 if any(cell.strip() for cell in row_1) else None,
                        'data_2': row_2 if any(cell.strip() for cell in row_2) else None,
                        'don_gia': don_gia_2
                    })
            elif not acc_code_1 and not acc_code_2:
                # 両方ともAccount Codeがない場合
                # 前の有効なAccount Codeを探す
                prev_code = None
                for code in reversed(list(items_dict.keys())):
                    if code:
                        prev_code = code
                        break
                if prev_code:
                    items_dict[prev_code].sub_items.append({
                        'data_1': row_1 if any(cell.strip() for cell in row_1) else None,
                        'data_2': row_2 if any(cell.strip() for cell in row_2) else None,
                        'don_gia': don_gia_2
                    })
    
    return items_dict, sub_items_list

def get_parent_code(account_code, items_dict):
    """親のAccount Codeを取得"""
    if not account_code:
        return None
    
    # 階層構造から親を推測
    # 例: 51131の親は5113、5113の親は511
    if len(account_code) > 3:
        parent = account_code[:-1]
        if parent in items_dict:
            return parent
    elif len(account_code) == 3:
        # 最上位レベル
        return None
    
    return None

def create_master_csv_v2(items_dict, output_filename):
    """マスターフォーマット#3をCSVとして出力（改良版）"""
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
            x[1].level if x[1].level is not None else 999,
            x[0]
        ))
        
        for acc_code, item in sorted_items:
            # メインデータ行
            row = [''] * 14
            level = item.level if item.level is not None else 0
            
            # Account Codeを階層に応じて配置
            row[level] = acc_code
            
            # データを統合（#2を優先、なければ#1）
            data = item.data_2 if item.data_2 else item.data_1
            
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
            row[9] = item.don_gia if item.don_gia else ''
            
            # メタデータ
            row[10] = item.get_source()
            row[11] = str(level)
            row[12] = 'Yes' if item.has_data_1() else 'No'
            row[13] = 'Yes' if item.has_data_2() else 'No'
            
            writer.writerow(row)
            
            # サブアイテムを追加
            for sub_item in item.sub_items:
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
                sub_row[10] = 'Both' if sub_item['data_1'] and sub_item['data_2'] else ('#1' if sub_item['data_1'] else '#2')
                
                writer.writerow(sub_row)

def main():
    input_file = '/Users/sugenotomohiro/Desktop/CURSOR/Manabox/20_DATA_Input/01.3consulting/Glory　Format - Compara.csv'
    output_file = '/Users/sugenotomohiro/Desktop/CURSOR/Manabox/20_DATA_Input/01.3consulting/マスターフォーマット#3_v2.csv'
    
    print("CSVファイルを詳細解析中...")
    items_dict, sub_items_list = parse_csv_detailed(input_file)
    
    print(f"マスター項目数: {len(items_dict)}件")
    print(f"サブアイテム総数: {sum(len(item.sub_items) for item in items_dict.values())}件")
    
    print("マスターフォーマット#3を作成中...")
    create_master_csv_v2(items_dict, output_file)
    
    print(f"完了: {output_file}")

if __name__ == '__main__':
    main()

