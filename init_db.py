# -*- coding: utf-8 -*-
"""
init_db.py v4.2 - 資料庫管理模組
v4.2 變更：修正 get_available_months 的 conn.cursor() 重複呼叫問題、補上所有遺缺函式
"""
import sqlite3, os, sys
from datetime import date, datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

import openpyxl
from config import DB_PATH, CATEGORY_MAP, CATEGORY_ORDER, MANUAL_CODE_MAPPING, NAME_OVERRIDE, DEFAULT_DISCOUNT_RATES


def safe_float(val):
    if val is None:
        return None
    try:
        return float(val)
    except (ValueError, TypeError):
        return None


def get_category(product_code):
    if not product_code or len(str(product_code)) < 2:
        return "其他"
    prefix = str(product_code)[:2]
    return CATEGORY_MAP.get(prefix, "其他")


def ensure_db(conn):
    cursor = conn.cursor()
    cursor.execute("""CREATE TABLE IF NOT EXISTS customer_mapping (
        id INTEGER PRIMARY KEY AUTOINCREMENT, cust_code TEXT NOT NULL UNIQUE,
        cust_short_name TEXT, cust_full_name TEXT, collector_code TEXT,
        collector_name TEXT, distributor_name TEXT NOT NULL,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)""")
    cursor.execute("""CREATE TABLE IF NOT EXISTS distributor_info (
        id INTEGER PRIMARY KEY AUTOINCREMENT, std_name TEXT NOT NULL UNIQUE,
        sheet_name TEXT, company_full TEXT, rate_dan_jin REAL DEFAULT 0.8,
        rate_fu_jin REAL DEFAULT 0.5, rate_dan_nong REAL DEFAULT 0.5,
        rate_fu_nong REAL DEFAULT 0.45, rate_otc REAL DEFAULT 0.5,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)""")
    cursor.execute("""CREATE TABLE IF NOT EXISTS import_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT, source_file TEXT, action TEXT,
        records_cust INTEGER DEFAULT 0, records_dist INTEGER DEFAULT 0,
        note TEXT, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)""")
    cursor.execute("""CREATE TABLE IF NOT EXISTS sales_records (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sale_date TEXT NOT NULL, invoice_no TEXT NOT NULL,
        cust_code TEXT, cust_name TEXT, product_code TEXT,
        category TEXT, amount REAL DEFAULT 0, distributor TEXT,
        input_by TEXT, batch_id INTEGER,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)""")
    cursor.execute("""CREATE TABLE IF NOT EXISTS import_batches (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        batch_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        input_by TEXT, total_lines INTEGER DEFAULT 0,
        new_records INTEGER DEFAULT 0, skipped_duplicates INTEGER DEFAULT 0,
        source TEXT, month_key TEXT)""")
    cursor.execute("""CREATE TABLE IF NOT EXISTS user_profiles (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT NOT NULL UNIQUE, full_name TEXT,
        is_active INTEGER DEFAULT 1,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)""")
    conn.commit()


def _ensure_v4_tables(conn):
    cursor = conn.cursor()
    cur_check = conn.cursor()
    cur_check.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='DIST_DISCOUNTS'")
    if not cur_check.fetchone():
        cur_check.execute("""CREATE TABLE DIST_DISCOUNTS (
                DISTINCT_NAME TEXT NOT NULL UNIQUE,
                DISCOUNT_VALUE REAL DEFAULT 0.8)""")
        rows = conn.execute("SELECT std_name, rate_dan_jin FROM distributor_info").fetchall()
        for r in rows:
            cur_check.execute("INSERT OR IGNORE INTO DIST_DISCOUNTS (DISTINCT_NAME, DISCOUNT_VALUE) VALUES (?, ?)", (r[0], r[1]))
    cur_check.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='CODE_TO_DIST'")
    if not cur_check.fetchone():
        cur_check.execute("""CREATE TABLE CODE_TO_DIST (
                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                CUSTOMER_CODE TEXT NOT NULL,
                CUSTOMER_NAME TEXT,
                DISTINCT_NAME TEXT NOT NULL,
                SHEET_NAME TEXT,
                IS_ACTIVE INTEGER DEFAULT 1)""")
        rows = conn.execute("SELECT cust_code, cust_short_name, cust_full_name, distributor_name FROM customer_mapping").fetchall()
        for r in rows:
            cur_check.execute("INSERT INTO CODE_TO_DIST (CUSTOMER_CODE, CUSTOMER_NAME, DISTINCT_NAME, SHEET_NAME) VALUES (?, ?, ?, ?)",
                               (r[0], r[2] or r[1], r[3], r[1]))
    conn.commit()

def ensure_user(conn, username):
    cursor = conn.cursor()
    cursor.execute("SELECT id FROM user_profiles WHERE username = ?", (username,))
    if not cursor.fetchone():
        cursor.execute("INSERT INTO user_profiles (username, full_name) VALUES (?, ?)", (username, username))
    conn.commit()


def insert_sales_records(records_list, input_by, batch_info):
    if not records_list:
        return {"new": 0, "skipped": 0}
    conn = sqlite3.connect(DB_PATH)
    ensure_db(conn)
    cursor = conn.cursor()
    batch_id = batch_info.get("id", None)
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new_count = 0
    skipped_count = 0
    for rec in records_list:
        inv_no = rec.get("inv_no", "") or ""
        d_val = rec.get("date")
        if d_val and isinstance(d_val, (datetime, date)):
            sale_date = d_val.strftime("%Y-%m-%d")
        else:
            sale_date = str(d_val) if d_val else now_str[:10]
        cursor.execute("SELECT id FROM sales_records WHERE invoice_no = ? AND cust_code = ? AND product_code = ?",
                       (inv_no, rec.get("cust_code", ""), rec.get("product_code", "")))
        if cursor.fetchone():
            skipped_count += 1
            continue
        cursor.execute("INSERT INTO sales_records (sale_date, invoice_no, cust_code, cust_name, product_code, category, amount, distributor, input_by, batch_id) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                       (sale_date, inv_no, rec.get("cust_code", ""), rec.get("cust_name", ""),
                        rec.get("product_code", ""), rec.get("category", "其他"),
                        rec.get("amount", 0.0), rec.get("distributor", ""), input_by, batch_id))
        new_count += 1
    if batch_id:
        cursor.execute("UPDATE import_batches SET new_records = ?, skipped_duplicates = ? WHERE id = ?", (new_count, skipped_count, batch_id))
    conn.commit()
    conn.close()
    return {"new": new_count, "skipped": skipped_count}


def get_sales_by_month(conn, month_key):
    cursor = conn.cursor()
    cursor.execute("SELECT sale_date, invoice_no, cust_code, cust_name, product_code, category, amount, distributor, input_by FROM sales_records WHERE sale_date LIKE ? ORDER BY sale_date, invoice_no", (month_key + "%",))
    rows = cursor.fetchall()
    result = []
    for r in rows:
        result.append({"date": r[0], "invoice_no": r[1], "cust_code": r[2], "cust_name": r[3],
                       "product_code": r[4], "category": r[5], "amount": r[6], "distributor": r[7], "input_by": r[8]})
    return result


def get_available_months(conn):
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT substr(sale_date, 1, 7) as month_key FROM sales_records WHERE sale_date IS NOT NULL AND sale_date != '' GROUP BY month_key ORDER BY month_key DESC")
    months = [r[0] for r in cursor.fetchall()]
    return months


def delete_month_data(conn, month_key, input_by):
    cursor = conn.cursor()
    cursor.execute("DELETE FROM sales_records WHERE sale_date LIKE ?", (month_key + "%",))
    deleted = cursor.rowcount
    conn.commit()
    return deleted


def get_recent_imports(conn, limit=50):
    cursor = conn.cursor()
    cursor.execute("SELECT id, batch_time, input_by, total_lines, new_records, skipped_duplicates, month_key FROM import_batches ORDER BY batch_time DESC LIMIT ?", (limit,))
    return cursor.fetchall()


def get_daily_trend(conn, month_key):
    cursor = conn.cursor()
    cursor.execute("SELECT substr(sale_date, 1, 10) as day, SUM(amount) as total FROM sales_records WHERE sale_date LIKE ? GROUP BY day ORDER BY day", (month_key + "%",))
    return cursor.fetchall()


def get_distributor_ranking(conn, month_key):
    cursor = conn.cursor()
    cursor.execute("SELECT distributor, COUNT(*) as cnt, SUM(amount) as total FROM sales_records WHERE sale_date LIKE ? AND distributor IS NOT NULL AND distributor != '' GROUP BY distributor ORDER BY total DESC", (month_key + "%",))
    return cursor.fetchall()


def get_category_breakdown(conn, month_key):
    cursor = conn.cursor()
    cursor.execute("SELECT category, SUM(amount) as total FROM sales_records WHERE sale_date LIKE ? GROUP BY category ORDER BY total DESC", (month_key + "%",))
    return cursor.fetchall()


def get_overview_stats(conn, month_key):
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*), COALESCE(SUM(amount),0), COUNT(DISTINCT invoice_no), COUNT(DISTINCT distributor) FROM sales_records WHERE sale_date LIKE ?", (month_key + "%",))
    row = cursor.fetchone()
    return {"total_lines": row[0], "total_amount": row[1], "invoice_count": row[2], "distributor_count": row[3]}


def import_from_excel(file_path):
    wb = openpyxl.load_workbook(file_path)
    conn = sqlite3.connect(DB_PATH)
    ensure_db(conn)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM customer_mapping")
    cursor.execute("DELETE FROM distributor_info")
    ws_cust = wb["客戶對照表"]
    cust_count = 0
    for row in ws_cust.iter_rows(min_row=2, max_col=5):
        code = str(row[0].value).strip() if row[0].value else ""
        if not code: continue
        short_name = str(row[1].value).strip() if row[1].value else ""
        full_name = str(row[2].value).strip() if row[2].value else ""
        collector_code = str(row[3].value).strip() if row[3].value else ""
        collector_name_raw = str(row[4].value).strip() if row[4].value else ""
        if not collector_name_raw or collector_name_raw == "#N/A":
            dist_name = MANUAL_CODE_MAPPING.get(code)
            if not dist_name: continue
        else:
            dist_name = NAME_OVERRIDE.get(collector_name_raw, collector_name_raw)
        cursor.execute("INSERT OR REPLACE INTO customer_mapping (cust_code, cust_short_name, cust_full_name, collector_code, collector_name, distributor_name) VALUES (?, ?, ?, ?, ?, ?)",
              (code, short_name, full_name, collector_code, collector_name_raw, dist_name))
        cust_count += 1
    ws_disc = wb["經銷商折數基準表"]
    dist_count = 0
    for row in ws_disc.iter_rows(min_row=2, max_col=8):
        std_name_raw = row[0].value
        if not std_name_raw: continue
        std_name = str(std_name_raw).strip()
        sheet_name = str(row[1].value) if row[1].value is not None else std_name
        company_full = str(row[2].value).strip() if row[2].value else std_name
        cursor.execute("INSERT OR REPLACE INTO distributor_info (std_name, sheet_name, company_full, rate_dan_jin, rate_fu_jin, rate_dan_nong, rate_fu_nong, rate_otc) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
              (std_name, sheet_name, company_full, safe_float(row[3].value), safe_float(row[4].value),
               safe_float(row[5].value), safe_float(row[6].value), safe_float(row[7].value)))
        dist_count += 1
    cursor.execute("INSERT INTO import_log (source_file, action, records_cust, records_dist, note) VALUES (?, 'import', ?, ?, ?)",
          (file_path, cust_count, dist_count, f"Import from {os.path.basename(file_path)}"))
    conn.commit()
    wb.close()
    ensure_user(conn, "system")
    return cust_count, dist_count


def load_mappings_from_db():
    if not os.path.exists(DB_PATH):
        raise FileNotFoundError("Database not found.")
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT cust_code, distributor_name FROM customer_mapping")
    code_to_dist = {row[0]: row[1] for row in cursor.fetchall()}
    cursor.execute("SELECT std_name, sheet_name, company_full, rate_dan_jin, rate_fu_jin, rate_dan_nong, rate_fu_nong, rate_otc FROM distributor_info")
    dist_info = {}
    for row in cursor.fetchall():
        sn = row[1] if row[1] is not None else row[0]
        cf = row[2] if row[2] is not None else row[0]
        dist_info[row[0]] = {"sheet_name": sn, "company_full": cf,
            "discount_rates": {"單斤": row[3] if row[3] is not None else 0.8, "複斤": row[4] if row[4] is not None else 0.5,
                                "單濃": row[5] if row[5] is not None else 0.5, "複濃": row[6] if row[6] is not None else 0.45,
                                "O.T.C.": row[7] if row[7] is not None else 0.5}}
    conn.close()
    return code_to_dist, dist_info


def get_distributor_list():
    if not os.path.exists(DB_PATH):
        return []
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT std_name FROM distributor_info ORDER BY std_name")
    result = [row[0] for row in cursor.fetchall()]
    conn.close()
    return result


if __name__ == "__main__":
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        file_path = os.path.expanduser("~/Desktop/Temp/02_經銷商對應與篩選對照表.xlsm")
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        sys.exit(1)
    try:
        import_from_excel(file_path)
    except Exception as e:
        print(f"Import failed: {e}")
        import traceback; traceback.print_exc()
        sys.exit(1)
