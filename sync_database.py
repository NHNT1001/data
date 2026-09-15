# -*- coding: utf-8 -*-
"""
Script tong hop du lieu tu 4 file Excel vao file database.json tich luy vinh vien.
- Giu nguyen du lieu cua cac thang/ngay cu (khong bao gio bi ghi de mat lich su).
- Cap nhat/Upsert du lieu cua cac ngay co trong file Excel hien tai.
- Tuan thu 100% logic nghiep vu:
  + Phan loai Nguon/Marketer
  + Loc bo don Google khoi Shopify
  + Khop dung 'Loai ket qua' cua tung kenh Ads
"""

import os
import re
import sys
import json
import warnings
from datetime import datetime, date
import openpyxl

# Suppress openpyxl styling warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# Force stdout to UTF-8
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

PLATFORM_CDTQ    = 'Chuyển Đổi Tệp Trung Quốc'
PLATFORM_MESSWS  = 'Mess/WS Tệp Trung Quốc'
PLATFORM_SHOPIFY = 'Shopify'
PLATFORM_UNKNOWN = 'Chưa phân loại'
PLATFORM_ORDER   = [PLATFORM_CDTQ, PLATFORM_MESSWS, PLATFORM_SHOPIFY, PLATFORM_UNKNOWN]

VALID_RESULT_TYPE = {
    PLATFORM_SHOPIFY: 'lượt mua trên web',
    PLATFORM_CDTQ: 'lượt đăng ký hoàn tất trên trang web',
    PLATFORM_MESSWS: 'lượt bắt đầu cuộc trò chuyện qua tin nhắn'
}

def classify_platform(marketer):
    s = str(marketer or '').strip().lower()
    if not s:
        return PLATFORM_UNKNOWN
    if 'shopify' in s:
        return PLATFORM_SHOPIFY
    if 'cđ tệp trung' in s or 'cd tep trung' in s or 'chuyển đổi' in s or 'chuyen doi' in s:
        return PLATFORM_CDTQ
    if 'mess' in s or 'ws tệp trung' in s or 'ws tep trung' in s:
        return PLATFORM_MESSWS
    return f"{PLATFORM_UNKNOWN}: {marketer}"

def platform_sort_key(p):
    for i, base in enumerate(PLATFORM_ORDER):
        if p.startswith(base):
            return i
    return 999

def parse_num(v):
    if v is None or v == '':
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    # VN format: 284.696,00 or 284696,00
    if re.match(r'^-?\d{1,3}(\.\d{3})*(,\d+)?$', s):
        s = s.replace('.', '').replace(',', '.')
    elif re.match(r'^-?\d+,\d+$', s):
        s = s.replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return 0.0

def norm_date(v):
    if isinstance(v, datetime):
        return v.strftime('%Y-%m-%d')
    if isinstance(v, date):
        return v.strftime('%Y-%m-%d')
    s = str(v).strip()
    if re.match(r'^\d{4}-\d{2}-\d{2}', s):
        return s[:10]
    m1 = re.match(r'^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})', s)
    if m1:
        d, mo, y = m1.groups()
        return f"{y}-{int(mo):02d}-{int(d):02d}"
    return s

def find_col(header_row, candidates):
    lower_cands = [c.lower() for c in candidates]
    headers_lower = [str(h or '').strip().lower() for h in header_row]
    for idx, h in enumerate(headers_lower):
        if h in lower_cands:
            return idx
    for idx, h in enumerate(headers_lower):
        for c in lower_cands:
            if c in h:
                return idx
    return -1

def read_excel_rows(file_path):
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    ws = wb.active
    rows = []
    for row in ws.iter_rows(values_only=True):
        rows.append(list(row))
    wb.close()
    return rows

def parse_orders(file_path):
    if not os.path.exists(file_path):
        return None, 0, []
    rows = read_excel_rows(file_path)
    if len(rows) < 2:
        return None, 0, []
    
    header = rows[0]
    i_date = find_col(header, ['ngày tạo đơn', 'ngày', 'date'])
    i_marketer = find_col(header, ['marketer'])
    i_amount = find_col(header, ['tổng tiền đơn hàng (trừ chiết khấu)', 'tổng tiền đơn hàng', 'tổng tiền'])
    i_source = find_col(header, ['utm source'])

    if i_date == -1 or i_amount == -1:
        print(f"[!] File {file_path} thieu cot bat buoc: Ngay / Tong tien")
        return None, 0, []

    out = []
    unmapped = set()
    excluded_google = 0

    for r in range(1, len(rows)):
        row = rows[r]
        if not row or all(v is None or v == '' for v in row):
            continue
        date_raw = row[i_date] if i_date < len(row) else None
        if date_raw is None or date_raw == '':
            continue
        
        amount = parse_num(row[i_amount] if i_amount < len(row) else 0)
        marketer = row[i_marketer] if (i_marketer != -1 and i_marketer < len(row)) else None
        platform = classify_platform(marketer)
        
        if platform.startswith(PLATFORM_UNKNOWN) and marketer:
            unmapped.add(str(marketer))

        # Loai don Google khoi Shopify
        if platform == PLATFORM_SHOPIFY and i_source != -1 and i_source < len(row):
            src = str(row[i_source] or '').strip().lower()
            if src == 'google':
                excluded_google += 1
                continue

        d_str = norm_date(date_raw)
        out.append({'date': d_str, 'platform': platform, 'amount': amount})

    return out, excluded_google, sorted(list(unmapped))

def parse_ads(file_path, platform_label):
    if not os.path.exists(file_path):
        return None
    rows = read_excel_rows(file_path)
    if len(rows) < 2:
        return None

    header = rows[0]
    i_date = find_col(header, ['ngày'])
    i_spend = find_col(header, ['số tiền đã chi tiêu (vnd)', 'số tiền đã chi tiêu', 'so tien da chi tieu'])
    i_result = find_col(header, ['kết quả', 'ket qua'])
    i_result_type = find_col(header, ['loại kết quả', 'loai ket qua'])

    if i_date == -1 or i_spend == -1:
        print(f"[!] File Ads {file_path} thieu cot bat buoc: Ngay / So tien")
        return None

    expected_type = VALID_RESULT_TYPE.get(platform_label)
    out = []

    for r in range(1, len(rows)):
        row = rows[r]
        if not row or all(v is None or v == '' for v in row):
            continue
        date_raw = row[i_date] if i_date < len(row) else None
        if date_raw is None or date_raw == '':
            continue

        spend = parse_num(row[i_spend] if i_spend < len(row) else 0)
        ketqua = 0.0
        if i_result != -1 and i_result < len(row):
            raw_kq = parse_num(row[i_result])
            if raw_kq > 0:
                type_val = str(row[i_result_type] or '').strip().lower() if (i_result_type != -1 and i_result_type < len(row)) else ''
                if not expected_type or type_val == expected_type:
                    ketqua = raw_kq

        d_str = norm_date(date_raw)
        out.append({'date': d_str, 'platform': platform_label, 'spend': spend, 'ketqua': ketqua})

    return out

def find_matching_file(folder, patterns):
    for fn in os.listdir(folder):
        if not fn.endswith('.xlsx') or fn.startswith('~$'):
            continue
        fn_lower = fn.lower()
        for p in patterns:
            if p.lower() in fn_lower:
                return os.path.join(folder, fn)
    return None

def main():
    target_dir = sys.argv[1] if len(sys.argv) > 1 else os.getcwd()
    db_path = os.path.join(target_dir, 'database.json')

    print(f"[>] Bat dau dong bo du lieu tai: {target_dir}")

    # 1. Doc database.json cu (neu co) de giu nguyen lich su cac ngay truoc
    existing_db = {'dates': [], 'platforms': [], 'data': {}, 'stats': {}}
    if os.path.exists(db_path):
        try:
            with open(db_path, 'r', encoding='utf-8') as f:
                existing_db = json.load(f)
            print(f"[+] Da load database cu: {len(existing_db.get('dates', []))} ngay lich su.")
        except Exception as e:
            print(f"[!] Khong the doc database cu ({e}), se tao moi.")

    db_data = existing_db.get('data', {})

    # 2. Tim cac file nguon
    orders_file = os.path.join(target_dir, 'orders.xlsx')
    if not os.path.exists(orders_file):
        orders_file = find_matching_file(target_dir, ['orders', 'don_hang'])

    cdtq_file = find_matching_file(target_dir, ['DATA-CĐ-Tệp-Trung', 'DATA-CD-Tep-Trung', 'DATA-C%C4%90'])
    messws_file = find_matching_file(target_dir, ['DATA-Mess-WS', 'DATA-Mess'])
    shopify_file = find_matching_file(target_dir, ['DATA-SHOPIFY'])

    print(f"  - Orders:   {os.path.basename(orders_file) if orders_file else 'KHONG THAY'}")
    print(f"  - Ads CĐTQ: {os.path.basename(cdtq_file) if cdtq_file else 'KHONG THAY'}")
    print(f"  - Ads MESS: {os.path.basename(messws_file) if messws_file else 'KHONG THAY'}")
    print(f"  - Ads SHOP: {os.path.basename(shopify_file) if shopify_file else 'KHONG THAY'}")

    # 3. Parse du lieu dot nay
    batch_data = {} # { date: { platform: {doanh_so, don_chot, ads, don_tao_moi} } }
    batch_dates = set()
    batch_platforms = set()

    def get_cell(d, p):
        if d not in batch_data:
            batch_data[d] = {}
        if p not in batch_data[d]:
            batch_data[d][p] = {'doanh_so': 0.0, 'don_chot': 0, 'ads': 0.0, 'don_tao_moi': 0}
        batch_dates.add(d)
        batch_platforms.add(p)
        return batch_data[d][p]

    total_orders = 0
    excluded_google = 0
    unmapped_marketers = []

    if orders_file and os.path.exists(orders_file):
        orders_rows, excluded_google, unmapped_marketers = parse_orders(orders_file)
        if orders_rows:
            total_orders = len(orders_rows)
            for o in orders_rows:
                cell = get_cell(o['date'], o['platform'])
                cell['doanh_so'] += o['amount']
                cell['don_chot'] += 1

    if cdtq_file and os.path.exists(cdtq_file):
        ads_cdtq = parse_ads(cdtq_file, PLATFORM_CDTQ)
        if ads_cdtq:
            for a in ads_cdtq:
                cell = get_cell(a['date'], a['platform'])
                cell['ads'] += a['spend']
                cell['don_tao_moi'] += int(a['ketqua'])

    if messws_file and os.path.exists(messws_file):
        ads_messws = parse_ads(messws_file, PLATFORM_MESSWS)
        if ads_messws:
            for a in ads_messws:
                cell = get_cell(a['date'], a['platform'])
                cell['ads'] += a['spend']
                cell['don_tao_moi'] += int(a['ketqua'])

    if shopify_file and os.path.exists(shopify_file):
        ads_shopify = parse_ads(shopify_file, PLATFORM_SHOPIFY)
        if ads_shopify:
            for a in ads_shopify:
                cell = get_cell(a['date'], a['platform'])
                cell['ads'] += a['spend']
                cell['don_tao_moi'] += int(a['ketqua'])

    print(f"[i] Parse thanh cong: {total_orders} don hang, {len(batch_dates)} ngay trong dot nay.")
    print(f"[i] Loai {excluded_google} don Google khoi Shopify.")
    if unmapped_marketers:
        print(f"[!] Phat hien Marketer chua phan loai: {', '.join(unmapped_marketers)}")

    # 4. UPSERT vao database tich luy (Giu nguyen cac ngay cu khong co trong batch)
    for d, plats in batch_data.items():
        if d not in db_data:
            db_data[d] = {}
        for p, metrics in plats.items():
            db_data[d][p] = metrics

    # 5. Cap nhat danh sach toan bo cac ngay va platforms
    all_dates = sorted(list(db_data.keys()))
    all_platforms_set = set()
    for d in all_dates:
        for p in db_data[d].keys():
            all_platforms_set.add(p)

    all_platforms = sorted(list(all_platforms_set), key=platform_sort_key)

    output = {
        'last_updated': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'dates': all_dates,
        'platforms': all_platforms,
        'data': db_data,
        'stats': {
            'excluded_google_shopify': excluded_google,
            'unmapped_marketers': unmapped_marketers
        }
    }

    # Ghi file atomic (qua file temp de tranh bi loi giua chung)
    temp_path = db_path + '.tmp'
    with open(temp_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    if os.path.exists(db_path):
        os.remove(db_path)
    os.rename(temp_path, db_path)

    print(f"[OK] Da cap nhat thanh cong database.json! Tong cong {len(all_dates)} ngay lich su.")

if __name__ == '__main__':
    main()
