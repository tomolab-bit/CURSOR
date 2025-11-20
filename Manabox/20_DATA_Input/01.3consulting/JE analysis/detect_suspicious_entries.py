#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ベトナム仕訳帳異常検出ツール
「賄賂」の疑いがある異常な仕訳を抽出します。
"""

import csv
import re
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Tuple

# 検出パターンの設定
SUSPICIOUS_ACCOUNT_CODES = {
    '6428',  # 旅費・交際費・会議費・その他費用（管理）
    '6418',  # 販売促進費・倉庫料・旅費・交際費・会議費（販売）
    '6278',  # その他製造間接費（その他）
    '6427',  # その他管理費
    '811',   # その他費用
}

ADVANCE_PAYMENT_KEYWORDS = [
    '立替', 'advance', 'reimbursement', 'tạm ứng', 'ứng trước',
    'prepayment', 'tạm ứng', 'ứng tiền'
]

CONSULTING_KEYWORDS = [
    'consulting', 'consultant', 'コンサル', 'advisory', 'tư vấn',
    'consult', 'adviser', 'コンサルタント'
]

SUNDRY_KEYWORDS = [
    '雑費', 'sundry', 'miscellaneous', 'chi phí khác', 'chi phí linh tinh',
    'その他', 'other expense', 'chi phí khác'
]

# 賄賂関連キーワード（明確に賄賂を示すキーワードのみ）
BRIBERY_KEYWORDS = [
    # 日本語（明確な賄賂関連）
    'キックバック', 'リベート', '裏金', '袖の下', 'わいろ', '賄賂',
    '見返り', '口利き', '便宜', '特別扱い', '不正', '違法', '闇', '隠し', '内密',
    '謝礼金', '謝金', '手数料（不正）', '仲介料（不正）', '紹介料（不正）',
    '黒い金', '闇金', '裏取引', '不正取引', '違法取引',
    
    # 英語（明確な賄賂関連）
    'bribe', 'kickback', 'rebate', 'facilitation payment', 'grease payment',
    'under the table', 'off the books', 'slush fund', 'brown envelope',
    'sweetener', 'payoff', 'hush money', 'corruption', 'illegal payment',
    'under table', 'black money', 'hidden payment', 'secret payment',
    'unofficial payment', 'backhander', 'baksheesh', 'boodle',
    
    # ベトナム語（明確な賄賂関連）
    'hối lộ', 'đút lót', 'tiền hối lộ', 'tiền bôi trơn', 'phong bì',
    'tiền lót tay', 'tiền đen', 'tiền ngầm', 'thu nhập bất hợp pháp',
    'tiền hối lộ ngầm', 'đút lót ngầm', 'quà biếu bất hợp pháp',
    
    # その他（明確な賄賂関連）
    'gratuity (illegal)', 'under table money', 'off books payment'
]

# 個人名パターン（ベトナム語の一般的な個人名形式）
PERSON_NAME_PATTERNS = [
    r'^[A-ZĐ][a-záàảãạăắằẳẵặâấầẩẫậéèẻẽẹêếềểễệíìỉĩịóòỏõọôốồổỗộơớờởỡợúùủũụưứừửữựýỳỷỹỵđ]+\s+[A-ZĐ][a-záàảãạăắằẳẵặâấầẩẫậéèẻẽẹêếềểễệíìỉĩịóòỏõọôốồổỗộơớờởỡợúùủũụưứừửữựýỳỷỹỵđ]+',  # Nguyen Van A形式
    r'^Mr\.\s+',  # Mr. で始まる
    r'^Ms\.\s+',  # Ms. で始まる
    r'^Mrs\.\s+',  # Mrs. で始まる
    r'^Ông\s+',  # Ông (ベトナム語のMr.)
    r'^Bà\s+',   # Bà (ベトナム語のMrs.)
    r'^Anh\s+',  # Anh (ベトナム語のMr. - 若い男性)
    r'^Chị\s+',  # Chị (ベトナム語のMs. - 若い女性)
]


def parse_amount(amount_str: str) -> float:
    """金額文字列を数値に変換（カンマ区切り対応）"""
    if not amount_str or amount_str.strip() == '' or amount_str == ',':
        return 0.0
    
    # カンマとドットを削除して数値に変換
    # ベトナム形式: 1.423.942,68 -> 1423942.68
    amount_str = amount_str.replace('.', '').replace(',', '.')
    
    try:
        return float(amount_str)
    except (ValueError, AttributeError):
        return 0.0


def is_person_name(name: str) -> bool:
    """取引先名が個人名かどうかを判定"""
    if not name or name.strip() == '':
        return False
    
    name = name.strip()
    
    # 会社名のキーワードが含まれている場合は個人名ではない
    company_keywords = ['Công ty', 'Cty', 'CT', 'TNHH', 'CP', 'Co.,Ltd', 'Ltd', 'Corp', 'Corporation', 'Inc']
    if any(keyword in name for keyword in company_keywords):
        return False
    
    # 個人名パターンに一致するかチェック
    for pattern in PERSON_NAME_PATTERNS:
        if re.match(pattern, name, re.IGNORECASE):
            return True
    
    # 短い名前（2-3単語）で、会社名キーワードがない場合は個人名の可能性
    words = name.split()
    if 2 <= len(words) <= 4:
        # すべての単語が大文字で始まる場合は個人名の可能性が高い
        if all(word[0].isupper() if word else False for word in words):
            return True
    
    return False


def detect_pattern_1_personal_advance(row: Dict) -> bool:
    """パターン1: 個人への立替払い"""
    supplier_name = row.get('supplier_name', '').strip()
    description = row.get('description', '').strip()
    debit_account = row.get('debit_account', '').strip()
    credit_account = row.get('credit_account', '').strip()
    
    # 個人名かどうかチェック
    if is_person_name(supplier_name):
        # 立替関連の科目コードか摘要に立替キーワードがあるか
        if debit_account in ['338', '141', '3388'] or any(keyword.lower() in description.lower() for keyword in ADVANCE_PAYMENT_KEYWORDS):
            return True
    
    # 摘要に立替キーワードがあり、かつ個人名の場合
    if any(keyword.lower() in description.lower() for keyword in ADVANCE_PAYMENT_KEYWORDS):
        if is_person_name(supplier_name):
            return True
    
    return False


def detect_pattern_2_sundry(row: Dict) -> bool:
    """パターン2: 雑費"""
    debit_account = row.get('debit_account', '').strip()
    description = row.get('description', '').strip()
    
    # 雑費関連の科目コードかチェック
    if debit_account in SUSPICIOUS_ACCOUNT_CODES:
        return True
    
    # 摘要に雑費キーワードがあるかチェック
    if any(keyword.lower() in description.lower() for keyword in SUNDRY_KEYWORDS):
        return True
    
    return False


def detect_pattern_3_local_consulting(row: Dict) -> bool:
    """パターン3: ローカルコンサルへの支払い"""
    supplier_name = row.get('supplier_name', '').strip()
    description = row.get('description', '').strip()
    debit_account = row.get('debit_account', '').strip()
    
    # 取引先名や摘要にコンサル関連キーワードがあるか
    text_to_check = f"{supplier_name} {description}".lower()
    
    if any(keyword.lower() in text_to_check for keyword in CONSULTING_KEYWORDS):
        # 雑費関連の科目コードと組み合わせられている場合
        if debit_account in ['6427', '811', '6418', '6428']:
            return True
    
    return False


def detect_pattern_4_bribery_keywords(row: Dict) -> bool:
    """パターン4: 備考欄に賄賂関連キーワードを含む仕訳"""
    description = row.get('description', '').strip()
    supplier_name = row.get('supplier_name', '').strip()
    
    # 備考欄と取引先名を結合してチェック
    text_to_check = f"{supplier_name} {description}".lower()
    
    # 賄賂関連キーワードが含まれているかチェック
    if any(keyword.lower() in text_to_check for keyword in BRIBERY_KEYWORDS):
        return True
    
    return False


def detect_pattern_5_amount_anomaly(row: Dict, all_entries: List[Dict]) -> bool:
    """パターン5: 金額ベースの異常検出"""
    amount = row.get('amount', 0)
    debit_account = row.get('debit_account', '').strip()
    
    # 異常に大きな金額の閾値（100,000 VND以上）
    LARGE_AMOUNT_THRESHOLD = 100000
    
    # ラウンドナンバーのチェック（1,000,000、5,000,000など）
    def is_round_number(amt):
        if amt == 0:
            return False
        # 100万、500万、1000万など
        round_numbers = [1000000, 2000000, 5000000, 10000000, 20000000, 50000000, 100000000]
        return any(abs(amt - rn) < rn * 0.01 for rn in round_numbers)
    
    # 特定科目での異常に大きな金額
    if debit_account in SUSPICIOUS_ACCOUNT_CODES:
        if amount >= LARGE_AMOUNT_THRESHOLD:
            return True
        if is_round_number(amount):
            return True
    
    # 個人への高額支払い
    supplier_name = row.get('supplier_name', '').strip()
    if is_person_name(supplier_name):
        if amount >= 50000:  # 個人への5万VND以上の支払い
            return True
    
    return False


def detect_pattern_6_vague_description(row: Dict) -> bool:
    """パターン6: 摘要の空欄・曖昧な記載の検出"""
    description = row.get('description', '').strip()
    debit_account = row.get('debit_account', '').strip()
    
    # 摘要が空欄
    if not description or description == '':
        # 雑費関連の科目コードの場合のみ検出
        if debit_account in SUSPICIOUS_ACCOUNT_CODES:
            return True
        return False
    
    # 曖昧な記載のキーワード
    vague_keywords = [
        'その他', 'other', 'miscellaneous', 'sundry', 'chi phí khác',
        'chi phí linh tinh', 'other expense', 'その他費用', 'その他経費',
        '雑費', 'misc', 'etc', 'various', 'khác'
    ]
    
    # 摘要が極端に短い（10文字以下）
    if len(description) <= 10:
        if any(keyword.lower() in description.lower() for keyword in vague_keywords):
            return True
    
    # 摘要が曖昧なキーワードのみ
    description_lower = description.lower()
    if any(keyword.lower() == description_lower.strip() for keyword in vague_keywords):
        return True
    
    # 摘要に具体的な情報がない（日付、金額、取引先名、商品名などが含まれていない）
    has_specific_info = any([
        re.search(r'\d{1,2}[/-]\d{1,2}[/-]\d{2,4}', description),  # 日付
        re.search(r'\d+[,.]\d+', description),  # 金額
        re.search(r'inv[:\s]*\d+', description, re.IGNORECASE),  # インボイス番号
        re.search(r'[A-Z]{2,}\d+', description),  # コード
    ])
    
    if not has_specific_info and debit_account in SUSPICIOUS_ACCOUNT_CODES:
        if len(description) < 30:  # 30文字未満で具体的な情報がない
            return True
    
    return False


def detect_pattern_7_supplier_pattern(row: Dict, all_entries: List[Dict]) -> bool:
    """パターン7: 取引先パターン分析"""
    supplier_name = row.get('supplier_name', '').strip()
    amount = row.get('amount', 0)
    date = row.get('date', '')
    
    if not supplier_name:
        return False
    
    # 同じ取引先への支払いを集計
    supplier_entries = [e for e in all_entries if e.get('supplier_name', '').strip() == supplier_name]
    
    if len(supplier_entries) < 2:
        return False
    
    # 短期間での複数回支払い（30日以内に3回以上）
    from datetime import datetime, timedelta
    
    try:
        entry_date = datetime.strptime(date, '%d/%m/%Y')
    except:
        try:
            entry_date = datetime.strptime(date, '%Y-%m-%d')
        except:
            return False
    
    recent_payments = 0
    for entry in supplier_entries:
        try:
            entry_date2 = datetime.strptime(entry.get('date', ''), '%d/%m/%Y')
        except:
            try:
                entry_date2 = datetime.strptime(entry.get('date', ''), '%Y-%m-%d')
            except:
                continue
        
        if abs((entry_date - entry_date2).days) <= 30:
            recent_payments += 1
    
    if recent_payments >= 3:
        return True
    
    # 個人名への高額支払い（10万VND以上）
    if is_person_name(supplier_name) and amount >= 100000:
        return True
    
    # 新規取引先への初回高額支払い（同じ取引先への最初の支払いが10万VND以上）
    supplier_entries_sorted = sorted(supplier_entries, key=lambda x: x.get('date', ''))
    if len(supplier_entries_sorted) > 0:
        first_entry = supplier_entries_sorted[0]
        if first_entry.get('date', '') == date and first_entry.get('amount', 0) >= 100000:
            return True
    
    return False


def detect_pattern_8_timeseries_pattern(row: Dict, all_entries: List[Dict]) -> bool:
    """パターン8: 時系列パターン分析"""
    date = row.get('date', '')
    debit_account = row.get('debit_account', '').strip()
    amount = row.get('amount', 0)
    
    if not date:
        return False
    
    from datetime import datetime
    
    try:
        entry_date = datetime.strptime(date, '%d/%m/%Y')
    except:
        try:
            entry_date = datetime.strptime(date, '%Y-%m-%d')
        except:
            return False
    
    # 月末に集中（月の最後の3日以内）
    if entry_date.day >= 28:
        if debit_account in SUSPICIOUS_ACCOUNT_CODES:
            return True
    
    # 年末に集中（12月の最後の5日以内）
    if entry_date.month == 12 and entry_date.day >= 27:
        if debit_account in SUSPICIOUS_ACCOUNT_CODES:
            return True
    
    # 同じ月に同じ科目で異常に多い取引（10件以上）
    month_entries = [
        e for e in all_entries
        if e.get('debit_account', '').strip() == debit_account
        and e.get('date', '').startswith(f"{entry_date.month:02d}/")
    ]
    
    if len(month_entries) >= 10:
        return True
    
    return False


def detect_pattern_9_correlation(row: Dict, all_entries: List[Dict]) -> bool:
    """パターン9: 相関分析"""
    supplier_name = row.get('supplier_name', '').strip()
    description = row.get('description', '').strip()
    amount = row.get('amount', 0)
    debit_account = row.get('debit_account', '').strip()
    
    if not supplier_name:
        return False
    
    # 同じ取引先への複数科目での支払い（3つ以上の異なる科目）
    supplier_entries = [e for e in all_entries if e.get('supplier_name', '').strip() == supplier_name]
    unique_accounts = set(e.get('debit_account', '').strip() for e in supplier_entries)
    
    if len(unique_accounts) >= 3:
        # そのうち雑費関連の科目が含まれている場合
        if any(acc in SUSPICIOUS_ACCOUNT_CODES for acc in unique_accounts):
            return True
    
    # 同じ摘要での複数回支払い（5回以上）
    same_description_entries = [
        e for e in all_entries
        if e.get('description', '').strip() == description
        and e.get('description', '').strip() != ''
    ]
    
    if len(same_description_entries) >= 5:
        # そのうち雑費関連の科目が含まれている場合
        if any(e.get('debit_account', '').strip() in SUSPICIOUS_ACCOUNT_CODES for e in same_description_entries):
            return True
    
    # 同じ金額での複数回支払い（3回以上、かつラウンドナンバー）
    round_numbers = [1000000, 2000000, 5000000, 10000000, 20000000, 50000000, 100000000]
    is_round = any(abs(amount - rn) < rn * 0.01 for rn in round_numbers)
    
    if is_round:
        same_amount_entries = [
            e for e in all_entries
            if abs(e.get('amount', 0) - amount) < amount * 0.01
        ]
        
        if len(same_amount_entries) >= 3:
            return True
    
    return False


def load_gl_data(csv_file_path: str) -> List[Dict]:
    """GLデータをCSVから読み込む"""
    entries = []
    
    # エンコーディングを自動検出
    encodings = ['utf-8-sig', 'utf-8', 'cp1252', 'latin-1', 'iso-8859-1', 'shift_jis']
    rows = None
    used_encoding = None
    
    for encoding in encodings:
        try:
            with open(csv_file_path, 'r', encoding=encoding) as f:
                reader = csv.reader(f, delimiter=';')
                rows = list(reader)
                used_encoding = encoding
                break
        except (UnicodeDecodeError, UnicodeError):
            continue
    
    if rows is None:
        raise ValueError(f"CSVファイルのエンコーディングを判別できませんでした。")
    
    print(f"エンコーディング: {used_encoding} で読み込みました")
    
    # ヘッダー行を探す（6行目がヘッダー）
    header_row_idx = 5  # 0-indexedなので5
    if len(rows) <= header_row_idx:
        raise ValueError("CSVファイルにヘッダー行が見つかりません")
    
    # データ行を処理（7行目から）
    for i in range(header_row_idx + 1, len(rows)):
        row = rows[i]
        
        if len(row) < 12:
            continue
        
        # 空行をスキップ
        if not any(row):
            continue
        
        # 金額を取得（借方または貸方の大きい方）
        debit_amount = parse_amount(row[8])
        credit_amount = parse_amount(row[9])
        amount = max(debit_amount, credit_amount)
        
        # 金額が0の場合はスキップ
        if amount == 0:
            continue
        
        entry = {
            'date': row[0].strip(),
            'voucher_code': row[1].strip(),
            'voucher_number': row[2].strip(),
            'customer_code': row[3].strip(),
            'supplier_name': row[4].strip(),
            'description': row[5].strip(),
            'debit_account': row[6].strip(),
            'credit_account': row[7].strip(),
            'debit_amount': debit_amount,
            'credit_amount': credit_amount,
            'amount': amount,
            'voucher_code2': row[10].strip() if len(row) > 10 else '',
            'department_code': row[11].strip() if len(row) > 11 else '',
        }
        
        entries.append(entry)
    
    return entries


def calculate_risk_score(entry: Dict, detected_patterns: List[str], all_entries: List[Dict]) -> Tuple[int, Dict]:
    """リスクスコアを計算"""
    score = 0
    score_details = {}
    
    amount = entry.get('amount', 0)
    description = entry.get('description', '').strip()
    supplier_name = entry.get('supplier_name', '').strip()
    debit_account = entry.get('debit_account', '').strip()
    
    # 基本スコア（パターン別）
    pattern_scores = {
        '個人への立替払い': 20,
        '雑費': 15,
        'ローカルコンサルへの支払い': 25,
        'Bribery Keywords in Description': 30,
        '金額ベースの異常': 20,
        '摘要の空欄・曖昧': 15,
        '取引先パターン異常': 20,
        '時系列パターン異常': 15,
        '相関分析異常': 25,
    }
    
    for pattern in detected_patterns:
        if pattern in pattern_scores:
            score += pattern_scores[pattern]
            score_details[f'パターン: {pattern}'] = f"+{pattern_scores[pattern]}点"
    
    # 金額スコア
    if amount >= 1000000:
        score += 30
        score_details['金額（100万VND以上）'] = "+30点"
    elif amount >= 100000:
        score += 20
        score_details['金額（10万〜100万VND）'] = "+20点"
    elif amount >= 10000:
        score += 10
        score_details['金額（1万〜10万VND）'] = "+10点"
    
    # パターン数スコア
    pattern_count = len(detected_patterns)
    if pattern_count >= 3:
        score += 25
        score_details['複数パターン（3つ以上）'] = "+25点"
    elif pattern_count == 2:
        score += 15
        score_details['複数パターン（2つ）'] = "+15点"
    
    # 摘要の状態スコア
    if not description or description == '':
        score += 10
        score_details['摘要が空欄'] = "+10点"
    elif len(description) < 20:
        vague_keywords = ['その他', 'other', 'miscellaneous', 'sundry', 'chi phí khác']
        if any(kw.lower() in description.lower() for kw in vague_keywords):
            score += 5
            score_details['摘要が曖昧'] = "+5点"
    
    # ラウンドナンバースコア
    round_numbers = [1000000, 2000000, 5000000, 10000000, 20000000, 50000000, 100000000]
    if any(abs(amount - rn) < rn * 0.01 for rn in round_numbers):
        score += 5
        score_details['ラウンドナンバー'] = "+5点"
    
    # スコアを100点満点に制限
    score = min(score, 100)
    
    return score, score_details


def detect_suspicious_entries(entries: List[Dict]) -> List[Dict]:
    """異常な仕訳を検出"""
    suspicious_entries = []
    
    for entry in entries:
        detected_patterns = []
        
        # パターン1: 個人への立替払い
        if detect_pattern_1_personal_advance(entry):
            detected_patterns.append('個人への立替払い')
        
        # パターン2: 雑費
        if detect_pattern_2_sundry(entry):
            detected_patterns.append('雑費')
        
        # パターン3: ローカルコンサルへの支払い
        if detect_pattern_3_local_consulting(entry):
            detected_patterns.append('ローカルコンサルへの支払い')
        
        # パターン4: 備考欄に賄賂関連キーワード
        if detect_pattern_4_bribery_keywords(entry):
            detected_patterns.append('Bribery Keywords in Description')
        
        # パターン5: 金額ベースの異常検出
        if detect_pattern_5_amount_anomaly(entry, entries):
            detected_patterns.append('金額ベースの異常')
        
        # パターン6: 摘要の空欄・曖昧な記載
        if detect_pattern_6_vague_description(entry):
            detected_patterns.append('摘要の空欄・曖昧')
        
        # パターン7: 取引先パターン分析
        if detect_pattern_7_supplier_pattern(entry, entries):
            detected_patterns.append('取引先パターン異常')
        
        # パターン8: 時系列パターン分析
        if detect_pattern_8_timeseries_pattern(entry, entries):
            detected_patterns.append('時系列パターン異常')
        
        # パターン9: 相関分析
        if detect_pattern_9_correlation(entry, entries):
            detected_patterns.append('相関分析異常')
        
        if detected_patterns:
            entry['detected_patterns'] = ' / '.join(detected_patterns)
            
            # リスクスコアを計算
            risk_score, score_details = calculate_risk_score(entry, detected_patterns, entries)
            entry['risk_score'] = risk_score
            entry['score_details'] = score_details
            
            suspicious_entries.append(entry)
    
    return suspicious_entries


def save_results(suspicious_entries: List[Dict], output_file: str):
    """検出結果をCSVに保存"""
    if not suspicious_entries:
        print("検出された異常な仕訳はありませんでした。")
        return
    
    with open(output_file, 'w', encoding='utf-8-sig', newline='') as f:
        fieldnames = [
            'date', 'voucher_code', 'voucher_number', 'customer_code',
            'supplier_name', 'description', 'debit_account', 'credit_account',
            'debit_amount', 'credit_amount', 'amount', 'detected_patterns',
            'risk_score', 'score_details', 'voucher_code2', 'department_code'
        ]
        
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        
        # リスクスコアの降順でソート（スコアが同じ場合は金額の降順）
        sorted_entries = sorted(suspicious_entries, key=lambda x: (x.get('risk_score', 0), x['amount']), reverse=True)
        
        for entry in sorted_entries:
            # スコア詳細を文字列に変換
            score_details_str = '; '.join([f"{k}: {v}" for k, v in entry.get('score_details', {}).items()])
            
            row = {
                'date': entry['date'],
                'voucher_code': entry['voucher_code'],
                'voucher_number': entry['voucher_number'],
                'customer_code': entry['customer_code'],
                'supplier_name': entry['supplier_name'],
                'description': entry['description'],
                'debit_account': entry['debit_account'],
                'credit_account': entry['credit_account'],
                'debit_amount': entry['debit_amount'],
                'credit_amount': entry['credit_amount'],
                'amount': entry['amount'],
                'detected_patterns': entry['detected_patterns'],
                'risk_score': entry.get('risk_score', 0),
                'score_details': score_details_str,
                'voucher_code2': entry.get('voucher_code2', ''),
                'department_code': entry.get('department_code', ''),
            }
            writer.writerow(row)


def print_summary(suspicious_entries: List[Dict]):
    """サマリー情報を表示"""
    if not suspicious_entries:
        print("\n検出された異常な仕訳はありませんでした。")
        return
    
    total_count = len(suspicious_entries)
    total_amount = sum(entry['amount'] for entry in suspicious_entries)
    
    # パターン別の集計
    pattern_counts = {}
    pattern_amounts = {}
    
    # リスクスコア別の集計
    high_risk_count = 0  # 80点以上
    medium_risk_count = 0  # 50-79点
    low_risk_count = 0  # 50点未満
    
    for entry in suspicious_entries:
        patterns = entry['detected_patterns'].split(' / ')
        for pattern in patterns:
            pattern_counts[pattern] = pattern_counts.get(pattern, 0) + 1
            pattern_amounts[pattern] = pattern_amounts.get(pattern, 0) + entry['amount']
        
        # リスクスコア別集計
        risk_score = entry.get('risk_score', 0)
        if risk_score >= 80:
            high_risk_count += 1
        elif risk_score >= 50:
            medium_risk_count += 1
        else:
            low_risk_count += 1
    
    print("\n" + "="*80)
    print("異常仕訳検出結果サマリー")
    print("="*80)
    print(f"\n総検出件数: {total_count:,}件")
    print(f"総金額: {total_amount:,.2f} VND")
    
    # リスクスコア別内訳
    print(f"\nリスクスコア別内訳:")
    print("-" * 80)
    print(f"  高リスク（80点以上）: {high_risk_count:,}件")
    print(f"  中リスク（50-79点）: {medium_risk_count:,}件")
    print(f"  低リスク（50点未満）: {low_risk_count:,}件")
    
    print(f"\nパターン別内訳:")
    print("-" * 80)
    
    for pattern in sorted(pattern_counts.keys()):
        count = pattern_counts[pattern]
        amount = pattern_amounts[pattern]
        print(f"  {pattern}:")
        print(f"    件数: {count:,}件")
        print(f"    金額: {amount:,.2f} VND")
    
    print("\n" + "="*80)
    print(f"\n上位10件（リスクスコア順）:")
    print("-" * 80)
    
    sorted_entries = sorted(suspicious_entries, key=lambda x: (x.get('risk_score', 0), x['amount']), reverse=True)
    for i, entry in enumerate(sorted_entries[:10], 1):
        risk_score = entry.get('risk_score', 0)
        risk_level = "高リスク" if risk_score >= 80 else "中リスク" if risk_score >= 50 else "低リスク"
        
        print(f"\n{i}. {entry['date']} | {entry['supplier_name']}")
        print(f"   摘要: {entry['description'][:80]}...")
        print(f"   借方: {entry['debit_account']} | 貸方: {entry['credit_account']}")
        print(f"   金額: {entry['amount']:,.2f} VND")
        print(f"   リスクスコア: {risk_score}点 ({risk_level})")
        print(f"   検出パターン: {entry['detected_patterns']}")


def main():
    """メイン処理"""
    # ファイルパス
    input_file = '/Users/sugenotomohiro/Desktop/CURSOR/Manabox/20_DATA_Input/01.3consulting/JE analysis/28_GL data for Analysis.csv'
    output_file = '/Users/sugenotomohiro/Desktop/CURSOR/Manabox/20_DATA_Input/01.3consulting/JE analysis/suspicious_entries_detected.csv'
    
    print("="*80)
    print("ベトナム仕訳帳異常検出ツール")
    print("="*80)
    print(f"\n入力ファイル: {input_file}")
    print(f"出力ファイル: {output_file}")
    
    try:
        # データ読み込み
        print("\nデータを読み込み中...")
        entries = load_gl_data(input_file)
        print(f"読み込み完了: {len(entries):,}件の仕訳")
        
        # 異常検出
        print("\n異常な仕訳を検出中...")
        suspicious_entries = detect_suspicious_entries(entries)
        print(f"検出完了: {len(suspicious_entries):,}件の異常な仕訳を検出")
        
        # 高リスクのみをフィルタリング（80点以上）
        high_risk_entries = [e for e in suspicious_entries if e.get('risk_score', 0) >= 80]
        
        # 結果保存（高リスクのみ）
        if high_risk_entries:
            print("\n結果を保存中（高リスクのみ）...")
            save_results(high_risk_entries, output_file)
            print(f"保存完了: {output_file}")
        elif suspicious_entries:
            print("\n高リスクの仕訳は検出されませんでした。")
            print("全検出結果を保存しますか？")
            save_results(suspicious_entries, output_file)
            print(f"保存完了: {output_file}")
        
        # サマリー表示（高リスクのみ）
        print_summary(high_risk_entries)
        
        print("\n✅ 処理完了！")
        
    except FileNotFoundError:
        print(f"❌ エラー: ファイルが見つかりません: {input_file}")
    except Exception as e:
        print(f"❌ エラーが発生しました: {e}")
        import traceback
        traceback.print_exc()


if __name__ == '__main__':
    main()

