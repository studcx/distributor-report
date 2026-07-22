# -*- coding: utf-8 -*-
"""
app.py v4.3 - 經銷商銷售對帳系統 (Streamlit Web)
使用方式: cd dist_report && streamlit run app.py
"""

import streamlit as st
import io, os, sqlite3, traceback
from datetime import datetime, date

from config import (
    PAGE_TITLE, UPLOAD_HINT_A, UPLOAD_HINT_B, DB_PATH,
    CATEGORY_ORDER, CATEGORY_MAP, SHEET_ORDER,
)
from processors import (
    parse_erp_paste, parse_erp_excel, aggregate_by_distributor,
    process_monthly_file, most_common_date,
)
from init_db import (
    _ensure_v4_tables, load_mappings_from_db, get_distributor_list,
    import_from_excel, ensure_db, ensure_user,
    insert_sales_records, get_sales_by_month, get_available_months,
    delete_month_data, get_recent_imports,
    get_daily_trend, get_distributor_ranking, get_category_breakdown,
    get_overview_stats, get_category,
)
from styles import MAIN_CSS
import styles
import pandas as pd
from collections import defaultdict




def _flatten_for_storage(inv_records):
    """Convert aggregated invoice records to flat list for storage."""
    flat = []
    for inv_no, record in inv_records.items():
        cats = record.get("cats", {})
        if isinstance(cats, dict):
            for cat, amt in cats.items():
                flat.append({
                    "date": record.get("date", ""),
                    "invoice_no": inv_no,
                    "cust_code": record.get("cust_code", ""),
                    "cust_name": record.get("cust_name", ""),
                    "product_code": "",
                    "category": cat if cat else "其他",
                    "amount": amt if amt else 0.0,
                    "distributor": record.get("distributor", None),
                })
        else:
            flat.append({
                "date": record.get("date", ""),
                "invoice_no": inv_no,
                "cust_code": record.get("cust_code", ""),
                "cust_name": record.get("cust_name", ""),
                "product_code": "",
                "category": "其他",
                "amount": 0.0,
                "distributor": record.get("distributor", None),
            })
    return flat



def init_page():
    """設定頁面配置與會話狀態預設值。"""
    st.set_page_config(page_title=PAGE_TITLE, layout="wide", initial_sidebar_state="expanded")
    defaults = {
        'mapping_ready': False, 'code_to_dist': {}, 'dist_info': {},
        'last_result': None, 'current_page': '資料輸入',
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v
    if not st.session_state.mapping_ready and os.path.exists(DB_PATH):
        try:
            code_to_dist, dist_info = load_mappings_from_db()
            if code_to_dist:
                st.session_state.code_to_dist = code_to_dist
                st.session_state.dist_info = dist_info
                st.session_state.mapping_ready = True
        except Exception:
            pass


def get_db():
    """取得資料庫連線並確保資料表已建立。"""
    conn = sqlite3.connect(DB_PATH)
    ensure_db(conn)
    _ensure_v4_tables(conn)
    return conn



def _filter_future_months(raw_months):
    """過濾未來的月份。"""
    from datetime import datetime
    now = datetime.now()
    current_str = f"{now.year}-{now.month:02d}"
    return [m for m in raw_months if m <= current_str]


def render_sidebar():
    """側邊欄：系統狀態與操作指南。"""
    with st.sidebar:
        st.markdown('<div id="main-header"><h1>經銷商銷售對帳系統</h1><p>v4.3</p></div>', unsafe_allow_html=True)
        if st.session_state.mapping_ready:
            dc = len(st.session_state.dist_info)
            cc = len(st.session_state.code_to_dist)
            st.success("✅ 資料庫已就緒")
            st.caption(f"經銷商：{dc} 筆 ｜ 客戶代號：{cc} 筆")
        elif os.path.exists(DB_PATH):
            st.warning("⚠️ 資料庫存在但未載入對照資料")
        else:
            st.info("尚未建立資料庫。請先至「設定」上傳經銷商對照表。")
        st.divider()
        with st.container(border=True):
            st.markdown("**📋 使用步驟**")
            st.markdown("1. **資料輸入** \u2014 貼上 ERP 資料，預覽並儲存")
            st.markdown("2. **儀表板** \u2014 查看銷售分析圖表")
            st.markdown("3. **對帳單檢視** \u2014 直接在網頁上看各經銷商對帳單")
            st.markdown("4. **資料管理** \u2014 查看與刪除已儲存的資料")
            st.markdown("5. **報表匯出** \u2014 選擇月份，產生對帳單 Excel")
            st.markdown("6. **設定** \u2014 管理經銷商對照資料")
        st.divider()
        st.caption(f"v4.1 | {datetime.now():%Y-%m-%d}")


def render_home():
     """首頁：ERP 資料輸入與預覽。"""
     st.markdown(styles.header_html("資料輸入", "貼上或上傳 ERP 銷售日報資料"), unsafe_allow_html=True)

     if not st.session_state.mapping_ready:
         st.warning("⚠️ 尚未載入對照資料。請先至「設定」頁面上傳經銷商對照表檔案。")
         return

     col_input, col_preview = st.columns([1, 2], gap='large')

     with col_input:
         tab_paste, tab_upload = st.tabs(["貼上資料（建議）", "上傳 Excel"])
         paste_text = None
         uploaded_file = None

         with tab_paste:
             st.caption("從 ERP 匯出後全選複製 (Ctrl+A / Cmd+A)，貼到下方：")
             paste_text = st.text_area("請在此貼上 ERP 資料：", height=250,
                 placeholder="日期[TAB]發票號碼[TAB]客戶代號...", key="erp_paste")
         with tab_upload:
             uploaded_file = st.file_uploader("選擇 ERP 匯出的 Excel 檔案", type=["xlsx"], key="erp_upload")
             if uploaded_file is not None:
                 st.success(f"✅ 已選取：`{uploaded_file.name}`")

         has_input = (paste_text is not None and paste_text.strip()) or (uploaded_file is not None)
         col_btn1, col_btn2 = st.columns([2, 1])
         with col_btn1:
             preview_clicked = st.button("🔍 預覽資料", use_container_width=True, type="primary", disabled=not has_input)
         with col_btn2:
             clear_clicked = st.button("🗑️ 清除")

         if clear_clicked:
             st.session_state.pop("parsed_data", None)
             st.rerun()

         parsed = st.session_state.get('parsed_data')
         if preview_clicked and parsed is None:
             try:
                 if paste_text and paste_text.strip():
                     inv_records, _unmatched, _total, _dates = parse_erp_paste(paste_text, st.session_state.code_to_dist)
                 else:
                     inv_records, _unmatched, _total, _dates = parse_erp_excel(uploaded_file, st.session_state.code_to_dist)
                 st.session_state["parsed_data"] = _flatten_for_storage(inv_records)
                 st.rerun()
             except Exception as e:
                 st.error(f"❌ 解析失敗：{e}")

         with col_btn1:
             if parsed is not None:
                 save_clicked = st.button("💾 儲存資料", use_container_width=True, type="primary")
                 if save_clicked:
                     try:
                         conn = get_db()
                         result = insert_sales_records(parsed, "admin", {})
                         conn.commit()
                         conn.close()
                         st.session_state["last_result"] = result
                         del st.session_state["parsed_data"]
                         st.success(f"✅ 儲存成功！新增 {result['new']} 筆，跳過重複 {result['skipped']} 筆。")
                         st.rerun()
                     except Exception as e:
                         st.error(f"❌ 儲存失敗：{e}")

     with col_preview:
         parsed = st.session_state.get('parsed_data')
         if not parsed:
             st.info("👈 請先在左側貼上或上傳資料，然後點擊「預覽資料」。")
         else:
             st.subheader(f"📊 預覽結果：共 {len(parsed)} 筆記錄")
             # 摘要統計
             dists = set()
             total_amt = 0.0
             for r in parsed:
                 d = r.get("distributor", "")
                 if d: dists.add(d)
                 total_amt += r.get("amount", 0.0) or 0.0
             c1, c2, c3 = st.columns(3)
             with c1:
                 st.metric("明細筆數", f"{len(parsed):,}")
             with c2:
                 st.metric("銷售總額", f"NT${total_amt:,.0f}")
             with c3:
                 st.metric("經銷商數", f"{len(dists)}")

             # 明細表格
             rows = []
             for r in parsed:
                 rows.append({"日期": str(r.get("date","")), "發票號碼": r.get("invoice_no",""),
                     "客戶代號": r.get("cust_code",""), "產品代碼": r.get("product_code",""),
                     "類別": r.get("category",""), "金額": r.get("amount",0.0),
                     "經銷商": r.get("distributor","")})
             df = pd.DataFrame(rows)
             if "金額" in df.columns:
                 df["金額"] = df["金額"].apply(lambda x: f"{x:,.0f}" if isinstance(x, (int,float)) else "")
             st.dataframe(df, use_container_width=True, hide_index=True, height=450)


def render_dashboard():
      """儀表板：銷售分析圖表。"""
      st.markdown(styles.header_html("儀表板", "銷售趨勢與分析圖表"), unsafe_allow_html=True)

      # 取得可用月份清單
      try:
          conn_tmp = get_db()
          months = _filter_future_months(get_available_months(conn_tmp))
          conn_tmp.close()
      except Exception as e:
          st.error(f"讀取月份清單失敗：{e}")
          return

      if not months:
          st.info("目前尚無資料。請先至「資料輸入」貼上 ERP 資料並儲存。")
          return

      selected_month = st.selectbox("選擇月份：", options=months, index=0, key="dash_month_select")

      # 查詢統計資料
      try:
          conn2 = get_db()
          stats = get_overview_stats(conn2, selected_month)
          daily_data = get_daily_trend(conn2, selected_month)
          dist_data = get_distributor_ranking(conn2, selected_month)
          cat_data = get_category_breakdown(conn2, selected_month)
          conn2.close()
      except Exception as e:
          st.error(f"讀取資料失敗：{e}")
          import traceback; st.exception(e)
          return

      # 統計卡片
      c1, c2, c3, c4 = st.columns(4)
      tl = stats.get("total_lines", 0) if stats else 0
      tc = stats.get("invoice_count", 0) if stats else 0
      ta = stats.get("total_amount", 0) if stats else 0
      td = stats.get("distributor_count", 0) if stats else 0
      with c1: st.metric('明細筆數', f'{tl:,}')
      with c2: st.metric('發票張數', f'{tc}')
      with c3: st.metric('銷售總額', f'NT${ta:,.0f}')
      with c4: st.metric('經銷商數', f'{td}')

      st.subheader("每日銷售趨勢")
      if daily_data and len(daily_data) > 0:
          import plotly.express as px
          df_t = pd.DataFrame(daily_data, columns=["日期", "金額"])
          df_t["日期"] = pd.to_datetime(df_t["日期"], errors="coerce")
          fig1 = px.line(df_t, x="日期", y="金額", title=None)
          fig1.update_layout(yaxis_title="銷售金額 (NT$)", height=350, xaxis_title=None)
          st.plotly_chart(fig1, use_container_width=True)

          # 每日明細表格
          with st.expander('每日明細表格'):
              df_t_display = df_t.copy()
              df_t_display["金額"] = df_t_display["金額"].apply(lambda x: f"{x:,.0f}" if isinstance(x, (int,float)) and not pd.isna(x) else "")
              st.dataframe(df_t_display, use_container_width=True, hide_index=True)
      else:
          st.info(f"{selected_month} 尚無銷售資料。")

      st.subheader("經銷商銷售排名")
      if dist_data and len(dist_data) > 0:
          df_r = pd.DataFrame(dist_data, columns=["經銷商", "筆數", "金額"])
          top_n = 15
          if len(df_r) > top_n:
              other_total = df_r.iloc[top_n:]["金額"].sum()
              df_top = df_r.iloc[:top_n].copy()
              df_top = pd.concat([df_top, pd.DataFrame([{"經銷商": "其他", "金額": other_total}])], ignore_index=True)
          else:
              df_top = df_r.copy()
          fig2 = px.bar(df_top.sort_values("金額", ascending=True), y="經銷商", x="金額", color="金額", color_discrete_sequence=["#4a6741"], title=None)
          fig2.update_layout(height=min(400, len(df_top)*35+60), showlegend=False, xaxis_title="銷售金額 (NT$)")
          st.plotly_chart(fig2, use_container_width=True)

          with st.expander('完整排名表格'):
              df_r = df_r.copy()
              df_r.insert(0, "排名", range(1, len(df_r)+1))
              df_r["金額"] = df_r["金額"].apply(lambda x: f"{x:,.0f}" if isinstance(x, (int,float)) else "")
              st.dataframe(df_r, use_container_width=True, hide_index=True)
      else:
          st.info(f"{selected_month} 尚無經銷商資料。")

      st.subheader("產品類別佔比分析")
      if cat_data and len(cat_data) > 0:
          df_c = pd.DataFrame(cat_data, columns=["類別", "金額"])
          colors_list = ["#4a6741", "#6b8f5e", "#8fb572", "#c5d9b3", "#ddd", "#eee"]
          fig3 = px.pie(df_c, values="金額", names="類別", hole=0.5, color_discrete_sequence=colors_list, title=None)
          fig3.update_traces(textposition="outside", textinfo="percent+label")
          fig3.update_layout(height=400)
          st.plotly_chart(fig3, use_container_width=True)

          with st.expander('類別明細表格'):
              total_amt_c = df_c["金額"].sum()
              df_c_copy = df_c.copy()
              df_c_copy["佔比"] = df_c_copy["金額"].apply(lambda x: f"{x/total_amt_c*100:.1f}%" if total_amt_c else "0%")
              df_c_copy["金額"] = df_c_copy["金額"].apply(lambda x: f"{x:,.0f}" if isinstance(x, (int,float)) else "")
              st.dataframe(df_c_copy[["類別", "金額", "佔比"]], use_container_width=True, hide_index=True)
      else:
          st.info(f"{selected_month} 尚無產品類別資料。")


def render_data_management():
       """資料管理：查看、篩選與刪除已儲存的資料。"""
       st.markdown(styles.header_html("資料管理", "查看、篩選與刪除已儲存的資料"), unsafe_allow_html=True)

       try:
           conn = get_db()
           available_months = _filter_future_months(get_available_months(conn))
       except Exception as e:
           st.error(f"讀取資料失敗：{e}")
           return

       if not available_months:
           st.info("資料庫中尚無資料。請先至「資料輸入」貼上並儲存資料。")
           return

       col_a, col_b = st.columns([1, 3])
       with col_a:
           selected = st.selectbox("選擇月份", available_months, key="dm_month_select")

       try:
           cur = conn.cursor()
           cur.execute("SELECT sale_date, invoice_no, cust_code, cust_name, product_code, category, amount, distributor, input_by, created_at FROM sales_records WHERE accounting_month = ? ORDER BY sale_date, invoice_no", (selected,))
           rows = cur.fetchall()
       except Exception as e:
           st.error(f"查詢失敗：{e}")
           conn.close(); return

       if not rows:
           st.info(f"{selected} 尚無資料記錄。")
           conn.close(); return

       total_amt = sum(r[6] or 0.0 for r in rows)
       unique_inv = len(set(r[1] for r in rows))
       unique_dist = len(set(r[7] for r in rows if r[7]))
       c1, c2, c3, c4 = st.columns(4)
       c1.metric("明細筆數", len(rows))
       c2.metric("總銷售金額", f"{total_amt:,.0f}")
       c3.metric("發票張數", unique_inv)
       c4.metric("經銷商數", unique_dist)

       st.divider()

       df = pd.DataFrame(rows, columns=["日期","發票號碼","客戶代號","客戶名稱","產品代碼","類別","金額","經銷商","輸入者","建立時間"])
       df["金額"] = df["金額"].apply(lambda x: f"{x:,.0f}" if isinstance(x, (int,float)) else "")
       st.dataframe(df, use_container_width=True, hide_index=True, height=450)

       st.divider()
       col_del1, col_del2 = st.columns([3, 1])
       with col_del1:
           st.warning(f"⚠️ 刪除 {selected} 的全部資料 — 無法復原！")
       with col_del2:
           if st.button("刪除該月全部資料", type="primary", key=f"del_{selected}"):
               st.session_state[f"confirmed_{selected}"] = True

       confirmed_key = f"confirmed_{selected}"
       if st.session_state.get(confirmed_key):
           try:
               deleted = delete_month_data(conn, selected, "admin")
               conn.commit()
               st.success(f"✅ 已刪除 {deleted} 筆記錄（{selected}）。")
               del st.session_state[confirmed_key]
           except Exception as e:
               st.error(f"刪除失敗：{e}")

       conn.close()


def _build_dist_groups_from_db(records):
       """從資料庫記錄建立經銷商分組。"""
       inv_records = {}
       for rec in records:
           inv_no = rec.get("invoice_no", "") or ""
           if not inv_no: continue
           dist = rec.get("distributor", "") or "Unmatched"
           cat = rec.get("category", "") or "其他"
           amt = rec.get("amount", 0.0) or 0.0
           if inv_no not in inv_records:
               inv_records[inv_no] = {"date": rec.get("date"), "cust_name": rec.get("cust_name",""), "cust_code": rec.get("cust_code",""), "distributor": dist, "cats": {}}
           inv_records[inv_no]["cats"][cat] = inv_records[inv_no]["cats"].get(cat, 0.0) + amt
       dist_groups, unmatched = aggregate_by_distributor(inv_records)
       return dist_groups, unmatched


def render_report():
       """報表匯出：選擇月份，產生經銷商對帳單 Excel。"""
       st.markdown(styles.header_html("報表匯出", "選擇月份，產生經銷商對帳單 Excel 檔案"), unsafe_allow_html=True)

       try:
           conn = get_db()
           available_months = _filter_future_months(get_available_months(conn))
       except Exception as e:
           st.error(f"讀取資料失敗：{e}")
           return

       if not available_months:
           st.info("資料庫中尚無資料。請先至「資料輸入」貼上並儲存資料。")
           return

       selected = st.selectbox("選擇月份", available_months, key="rep_month_select")
       records = get_sales_by_month(conn, selected)
       if not records:
           st.info(f"{selected} 尚無資料。")
           conn.close(); return

       total_amt = sum(r.get("amount", 0.0) or 0.0 for r in records)
       unique_dists = len(set(r.get("distributor","") for r in records if r.get("distributor")))
       c1, c2, c3 = st.columns(3)
       c1.metric("明細筆數", len(records))
       c2.metric("銷售總額", f"NT${total_amt:,.0f}")
       c3.metric("經銷商數", unique_dists)

       st.divider()

       if st.button("產生 Excel 報表", type="primary", key="gen_report_btn"):
           try:
               dist_info = dict(st.session_state.dist_info)
               cur2 = conn.cursor()
               cur2.execute("SELECT DISTINCT_NAME, DISCOUNT_VALUE FROM DIST_DISCOUNTS")
               for row in cur2.fetchall():
                   if row[1] is not None: dist_info[row[0]]["discount"] = float(row[1])

               dist_groups, unmatched = _build_dist_groups_from_db(records)
               output_bytes, output_filename, stats_out, log_lines = process_monthly_file(dist_groups, dist_info, selected)

               st.download_button(label="下載 Excel 檔案", data=output_bytes, file_name=output_filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", key="dl_report_btn")

               with st.expander('匯出日誌'):
                   for line in log_lines: st.text(line)
           except Exception as e:
               st.error(f"產生報表失敗：{e}")
               import traceback; st.exception(e)

       conn.close()





def _render_settings_code_search(conn):
    """Customer code search with pagination."""
    st.subheader("客戶代號搜尋")
    search = st.text_input("輸入代號或名稱以搜尋", key="code_search_settings")
    try:
        cur3 = conn.cursor()
        cur3.execute("SELECT COUNT(*) FROM CODE_TO_DIST WHERE IS_ACTIVE > 0")
        total_count = cur3.fetchone()[0]
        if search:
            cur3.execute("SELECT CUSTOMER_CODE, CUSTOMER_NAME, DISTINCT_NAME FROM CODE_TO_DIST WHERE IS_ACTIVE > 0 AND (CUSTOMER_CODE LIKE ? OR CUSTOMER_NAME LIKE ?) ORDER BY CUSTOMER_CODE", ("%"+search+"%", "%"+search+"%"))
        else:
            cur3.execute("SELECT CUSTOMER_CODE, CUSTOMER_NAME, DISTINCT_NAME FROM CODE_TO_DIST WHERE IS_ACTIVE > 0 ORDER BY CUSTOMER_CODE")
        all_rows = cur3.fetchall()
        page_size = 25
        total_pages = max(1, (len(all_rows) + page_size - 1) // page_size)
        if "code_page" not in st.session_state:
            st.session_state.code_page = 1
        if search:
            st.session_state.code_page = 1
        start_idx = (st.session_state.code_page - 1) * page_size
        end_idx = min(start_idx + page_size, len(all_rows))
        page_data = all_rows[start_idx:end_idx]
        col_p_left, col_p_mid, col_p_right = st.columns([0.35, 0.3, 0.35], gap="small")
        with col_p_left:
            if st.button("\u2b05\uFE0F 上一頁", key="code_prev_btn", disabled=st.session_state.code_page <= 1):
                st.session_state.code_page -= 1
                st.rerun()
        with col_p_mid:
            pg_text = f"第 {st.session_state.code_page} / {total_pages} 頁 ｜ 共 {len(all_rows)} 筆"
            st.markdown(f'<div style="text-align:center;font-size:0.95rem;color:#374151;padding:8px 0;">{pg_text}</div>', unsafe_allow_html=True)
        with col_p_right:
            if st.button("下一頁 \u27a1\uFE0F", key="code_next_btn", disabled=st.session_state.code_page >= total_pages):
                st.session_state.code_page += 1
                st.rerun()
        if page_data:
            df_s = pd.DataFrame(page_data, columns=["客戶代號", "客戶名稱", "經銷商"])
            st.dataframe(df_s, use_container_width=True, hide_index=True, height=350)
        else:
            st.info("目前尚無資料。")
    except Exception as e:
        st.error(f"搜尋失敗：{e}")


def render_settings():
        """設定：管理經銷商對照資料與系統設定。"""
        conn = get_db()
        _ensure_v4_tables(conn)
        st.markdown(styles.header_html("設定", "管理經銷商對照資料與系統設定"), unsafe_allow_html=True)

        with st.container(border=True):
            st.subheader("上傳對照表檔案")
            up_file = st.file_uploader("上傳 02_經銷商對應與篩選對照表 (.xlsx / .xlsm)", type=["xlsx", "xlsm"], key="map_upload_settings")
            if up_file is not None:
                try:
                    result = import_from_excel(up_file)
                    if result.get("success"): 
                        st.success(f"匯入成功：經銷商 {result['distributors']} 筆、客戶代號 {result['codes']} 筆。")
                        code_to_dist2, dist_info2 = load_mappings_from_db()
                        st.session_state.code_to_dist = code_to_dist2
                        st.session_state.dist_info = dist_info2
                        st.session_state.mapping_ready = True
                    else:
                        msg = result.get("message", "未知錯誤")
                        st.error(f"匯入失敗：{msg}")
                except Exception as e:
                    st.error(f"匯入時發生錯誤：{e}")

        with st.container(border=True):
            st.subheader("經銷商清單與折數（各類藥品獨立）")
            try:
                cur2 = conn.cursor()
                cur2.execute("SELECT std_name, rate_dan_jin, rate_fu_jin, rate_dan_nong, rate_fu_nong, rate_otc FROM distributor_info ORDER BY std_name")
                dist_rows = cur2.fetchall()
                if dist_rows:
                    df_d = pd.DataFrame(dist_rows, columns=["經銷商名稱", "單斤折數", "複斤折數", "單濃折數", "複濃折數", "O.T.C.折數"])
                    for col in ["單斤折數", "複斤折數", "單濃折數", "複濃折數", "O.T.C.折數"]:
                        df_d[col] = df_d[col].apply(lambda x, c=col: f"{x*100:.1f}%" if x is not None and x > 0 else "\u2014")
                    st.dataframe(df_d, use_container_width=True, hide_index=True)
                else:
                    st.info("目前尚無經銷商資料。請先上傳對照表檔案。")
            except Exception as e:
                st.error(f"載入經銷商清單失敗：{e}")



        with st.container(border=True):
            _render_settings_code_search(conn)

        with st.container(border=True):
            st.subheader("使用者管理")
            col_u1, col_u2 = st.columns([3, 1])
            with col_u1:
                new_user = st.text_input("新增使用者名稱", key="new_user_settings")
            with col_u2:
                if st.button("新增", key="add_user_settings"): 
                    if new_user.strip():
                        ensure_user(conn, new_user.strip())
                        st.success(f"使用者 '{new_user.strip()}' 已新增。")
            try:
                cur4 = conn.cursor(); cur4.execute("SELECT username, full_name, is_active FROM user_profiles ORDER BY username")
                u_rows = cur4.fetchall()
                if u_rows:
                    df_u = pd.DataFrame(u_rows, columns=["使用者名稱", "全名", "啟用"])
                    st.dataframe(df_u, use_container_width=True, hide_index=True)
            except Exception as e:
                st.error(f"載入使用者清單失敗：{e}")

        conn.close()




def render_statement_view():
    """在網頁上模擬 Excel 04_客戶對帳單的排版，直接瀏覽各經銷商對帳單"""
    conn = get_db()
    try:
        raw_months = get_available_months(conn)
        available_months = _filter_future_months(raw_months)
    except Exception:
        available_months = []
    if not available_months:
        st.info("資料庫中尚無資料。請先至「資料輸入」貼上並儲存資料。")
        conn.close()
        return

    col_a, col_b = st.columns([1, 2])
    with col_a:
        selected_month = st.selectbox("選擇月份", available_months, key="stmt_month_select")
    try:
        cursor = conn.cursor()
        sql_distinct = "SELECT DISTINCT distributor FROM sales_records WHERE accounting_month = ? AND distributor IS NOT NULL AND distributor != '' GROUP BY distributor ORDER BY distributor"
        cursor.execute(sql_distinct, (selected_month,))
        active_dists = [r[0] for r in cursor.fetchall()]
    except Exception:
        active_dists = []
    if not active_dists:
        st.info(f"{selected_month} 尚無資料。")
        conn.close()
        return

    with col_b:
        selected_dist = st.selectbox("選擇經銷商", active_dists, key="stmt_dist_select")

    # Fetch records for this distributor and month
    try:
        cursor = conn.cursor()
        sql_records = "SELECT sale_date, invoice_no, cust_name, product_code, category, amount FROM sales_records WHERE accounting_month = ? AND distributor = ? ORDER BY sale_date, invoice_no"
        cursor.execute(sql_records, (selected_month, selected_dist))
        records = cursor.fetchall()
    except Exception as e:
        st.error(f"查詢失敗：{e}")
        conn.close()
        return

    # Get discount rates for this distributor
    dist_info = st.session_state.dist_info.get(selected_dist, {})
    rates = dist_info.get("discount_rates", {})
    r_dj = rates.get("單斤", 0.8)
    r_fj = rates.get("複斤", 0.5)
    r_dn = rates.get("單濃", 0.5)
    r_fn = rates.get("複濃", 0.45)
    r_otc = rates.get("O.T.C.", 0.5)

    # Group by invoice_no, aggregate by category
    inv_records = defaultdict(lambda: {"date": None, "cust_name": "", "cats": {}})
    for rec in records:
        inv_no = rec[1]
        cat = rec[4] or "其他"
        amt = rec[5] or 0.0
        r = inv_records[inv_no]
        if not r["date"]:
            r["date"] = rec[0]
            r["cust_name"] = rec[2] or ""
        r["cats"][cat] = r["cats"].get(cat, 0.0) + amt

    # Calculate totals
    cat_labels = ["單斤", "複斤", "單濃", "複濃", "O.T.C.", "其他"]
    rate_map = {"單斤": r_dj, "複斤": r_fj, "單濃": r_dn, "複濃": r_fn, "O.T.C.": r_otc, "其他": 1.0}
    total_brand = {cat: sum(inv_records[k]["cats"].get(cat, 0) for k in inv_records) for cat in cat_labels}
    total_actual = {}
    for cat in cat_labels:
        if cat != "其他":
            total_actual[cat] = round(total_brand[cat] * rate_map[cat])
        else:
            total_actual[cat] = total_brand[cat]
    grand_brand = sum(total_brand.values())
    grand_actual = sum(total_actual.values())

    # Parse month for display
    y_str, m_str = selected_month.split("-")
    rocy_year = int(y_str) - 1911
    month_title = f"{rocy_year}年{int(m_str):02d}月"

    # === Header card ===
    h = []
    h.append('<div style="background:linear-gradient(135deg,#4a6741 0%,#6b8f5e 100%);color:white;padding:1rem 1.5rem;border-radius:12px;margin-bottom:1rem;display:flex;justify-content:space-between;align-items:center;">')
    h.append('<div><div style="font-size:0.85rem;opacity:0.9;">' + month_title + ' 對帳單</div>')
    h.append('<div style="font-size:1.3rem;font-weight:700;margin-top:0.25rem;">' + selected_dist + '</div></div>')
    h.append('<div style="text-align:right;"><div style="font-size:0.85rem;opacity:0.9;">共 ' + str(len(inv_records)) + ' 筆銷貨</div></div>')
    h.append('</div>')
    st.markdown("".join(h), unsafe_allow_html=True)

    # === Discount rate cards ===
    _build_dist_rate_cards(r_dj, r_fj, r_dn, r_fn, r_otc)

    st.divider()

    # === Summary table ===
    st.subheader("類別合計對照")
    _render_summary_table(total_brand, total_actual, grand_brand, grand_actual, cat_labels)

    st.divider()

    # === Detail table ===
    st.subheader(f"明細列表（{len(inv_records)} 筆）")
    if inv_records:
        detail_rows = []
        for inv_no in inv_records:
            r = inv_records[inv_no]
            cats = r["cats"]
            date_str = str(r["date"]) if r["date"] else ""
            inv_total = sum(cats.values())
            detail_rows.append({
                "日期": date_str, "銷貨單號": inv_no, "客戶名稱": r["cust_name"],
                "單斤": f"{cats.get('單斤', 0):,.0f}", "複斤": f"{cats.get('複斤', 0):,.0f}",
                "單濃": f"{cats.get('單濃', 0):,.0f}", "複濃": f"{cats.get('複濃', 0):,.0f}",
                "O.T.C.": f"{cats.get('O.T.C.', 0):,.0f}", "其他": f"{cats.get('其他', 0):,.0f}",
                "小計": f"{inv_total:,.0f}",
            })
        _render_detail_table(detail_rows, total_brand, grand_brand, cat_labels)

    conn.close()


def _build_dist_rate_cards(r_dj, r_fj, r_dn, r_fn, r_otc):
    rate_items = [("單斤", r_dj), ("複斤", r_fj), ("單濃", r_dn), ("複濃", r_fn), ("O.T.C.", r_otc)]
    color_map_r = {"單斤": "#10b981", "複斤": "#3b82f6", "單濃": "#f59e0b", "複濃": "#ef4444", "O.T.C.": "#8b5cf6"}
    bg_map_r = {"單斤": "#d1fae5", "複斤": "#dbeafe", "單濃": "#fef3c7", "複濃": "#fee2e2", "O.T.C.": "#ede9fe"}
    cols_r = st.columns(len(rate_items))
    for ci, (label, rate_val) in enumerate(rate_items):
        with cols_r[ci]:
            pct_str = f"{rate_val*100:.0f}%"
            clr = color_map_r.get(label, '#6b7280')
            bgr = bg_map_r.get(label, '#f3f4f6')
            card_html = '<div style="background:' + bgr + ';padding:0.6rem 1rem;border-radius:8px;text-align:center;border-left:3px solid ' + clr + ';">'
            card_html += '<div style="font-size:0.75rem;color:#6b7280;">' + label + '</div>'
            card_html += '<div style="font-size:1.1rem;font-weight:700;color:' + clr + ';">' + pct_str + '</div></div>'
            st.markdown(card_html, unsafe_allow_html=True)


def _render_summary_table(total_brand, total_actual, grand_brand, grand_actual, cat_labels):
    brand_row = ["牌價"] + [f"{total_brand[c]:,.0f}" for c in cat_labels] + [f"{grand_brand:,.0f}"]
    actual_row = ["實價"] + [f"{total_actual[c]:,.0f}" for c in cat_labels] + [f"{grand_actual:,.0f}"]
    recv_row = ["應收帳款"] + [""] * (len(cat_labels) + 1)
    recv_row[-1] = f"{grand_actual:,.0f}"
    hdr = "<tr>" + "".join(f"<th style='padding:6px;text-align:center;background:#4a6741;color:white;'>{h}</th>" for h in ["項目"] + cat_labels + ["合計"]) + "</tr>"
    body = ""
    row_styles_s = ["", "background:#f0fdf4;", "background:#d1fae5;font-weight:700;color:#065f46;"]
    for ri, rv in enumerate([brand_row, actual_row, recv_row]):
        body += f"<tr style='{row_styles_s[ri]}'>" + "".join(f"<td style='padding:6px;text-align:right;'{'font-weight:bold;' if ri == 2 else ''}>{v}</td>" for v in rv) + "</tr>"
    html = '<div style="border-radius:8px;overflow:hidden;border:1px solid #e5e7eb;"><table style="width:100%;border-collapse:collapse;font-size:0.9rem;">' + hdr + body + '</table></div>'
    st.markdown(html, unsafe_allow_html=True)


def _render_detail_table(detail_rows, total_brand, grand_brand, cat_labels):
    d_hdr = "<tr>" + "".join(f"<th style='padding:6px;text-align:center;background:#4a6741;color:white;white-space:nowrap;'>{h}</th>" for h in ["日期","銷貨單號","客戶名稱","單斤","複斤","單濃","複濃","O.T.C.","其他","小計"]) + "</tr>"
    d_body = ""
    col_aligns = ["center","center","left"] + ["right"]*7
    for idx, dr in enumerate(detail_rows):
        bg = "#ffffff" if idx % 2 == 0 else "#f9fafb"
        vals = [dr[h] for h in ["日期","銷貨單號","客戶名稱","單斤","複斤","單濃","複濃","O.T.C.","其他","小計"]]
        d_body += f"<tr style='background:{bg};'>" + "".join(f"<td style='padding:5px;text-align:{a};'>{v}</td>" for v, a in zip(vals, col_aligns)) + "</tr>"
    t_vals = ["", "", "合計"] + [f"{total_brand[c]:,.0f}" for c in ["單斤","複斤","單濃","複濃","O.T.C."]] + [f"{total_brand['其他']:,.0f}", f"{grand_brand:,.0f}"]
    t_cells = ""
    for ti, tv in enumerate(t_vals):
        al = "center" if ti < 2 else ("left" if ti == 2 else "right")
        t_cells += f"<td style='padding:6px;text-align:{al};background:#d1fae5;color:#065f46;font-weight:700;'>{tv}</td>"
    d_body += f"<tr>{t_cells}</tr>"
    html = '<div style="border-radius:8px;overflow:hidden;border:1px solid #e5e7eb;"><table style="width:100%;border-collapse:collapse;font-size:0.85rem;">' + d_hdr + d_body + '</table></div>'
    st.markdown(html, unsafe_allow_html=True)



def render_import_history():
        """匯入紀錄：所有資料輸入的批次記錄。"""
        conn = get_db()
        st.markdown(styles.header_html("匯入紀錄", "所有資料輸入的批次記錄"), unsafe_allow_html=True)
        batches = get_recent_imports(conn, limit=50)
        if not batches:
            st.info("目前尚無匯入紀錄。")
            conn.close()
            return
        df_h = pd.DataFrame(batches, columns=["批次號", "時間", "輸入者", "總行數", "新增筆數", "跳過重複", "月份"])
        st.dataframe(df_h, use_container_width=True, hide_index=True)
        conn.close()


def main():
        """主程式：使用 radio 導航，避免 tabs 同時執行所有分頁的問題。"""
        init_page()
        st.markdown("<style>" + MAIN_CSS + "</style>", unsafe_allow_html=True)
        render_sidebar()

        nav_options = ["資料輸入", "儀表板", "對帳單檢視", "資料管理", "報表匯出", "設定"]
        page = st.radio("選擇頁面：", nav_options, horizontal=True, index=0, key="main_nav")

        try:
            if page == "資料輸入": render_home()
            elif page == "儀表板": render_dashboard()
            elif page == "對帳單檢視": render_statement_view()
            elif page == "資料管理": render_data_management()
            elif page == "報表匯出": render_report()
            elif page == "設定": render_settings()
        except Exception as e:
            st.error(f"頁面載入錯誤：{e}")
            import traceback; st.exception(e)


if __name__ == "__main__":
        main()

