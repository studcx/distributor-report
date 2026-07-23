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
    delete_month_data, get_recent_imports, get_dashboard_data,
    get_category, get_distributor_note, save_distributor_note,
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
                    "order_note": record.get("order_note", ""),
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
                "order_note": record.get("order_note", ""),
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



        # === 貼上資料區 ===

        with tab_paste:

            st.caption("從 ERP 匯出後全選複製 (Cmd+A)，貼到下方：")

            paste_text = st.text_area(

                "請在此貼上 ERP 資料：", height=250,

                placeholder="日期[TAB]發票號碼[TAB]客戶代號...",

                key="erp_paste_text")



            # 僅在此分頁顯示「預覽」按鈕（有解析結果時不顯示）

            if st.session_state.get('parsed_data') is None:

                has_text = (paste_text is not None and paste_text.strip())

                col_btn1, col_btn2 = st.columns([3, 1])

                with col_btn1:

                    if st.button("🔍 預覽資料", use_container_width=True, type="primary", disabled=not has_text):

                        try:

                            inv_records, _unmatched, _total, _dates = parse_erp_paste(paste_text, st.session_state.code_to_dist)

                            st.session_state["parsed_data"] = _flatten_for_storage(inv_records)

                            st.rerun()

                        except Exception as e:

                            st.error(f"❌ 解析失敗：{e}")

                with col_btn2:

                    if st.button("🗑️ 清除", use_container_width=True, key="clear_paste_btn"):

                        st.session_state.pop("parsed_data", None)

                        st.rerun()



        # === 上傳 Excel 區 ===

        with tab_upload:

            uploaded_file = st.file_uploader("選擇 ERP 匯出的 Excel 檔案", type=["xlsx"], key="erp_upload")

            if uploaded_file is not None:

                st.success(f"✅ 已選取：`{uploaded_file.name}`")



                # 僅在此分頁顯示「預覽」按鈕（有解析結果時不顯示）

                if st.session_state.get('parsed_data') is None:

                    col_btn_u1, col_btn_u2 = st.columns([3, 1])

                    with col_btn_u1:

                        if st.button("🔍 預覽資料", use_container_width=True, key="upload_preview_btn"):

                            try:

                                file_data = uploaded_file.read()

                                inv_records, _unmatched, _total, _dates = parse_erp_excel(file_data, st.session_state.code_to_dist)

                                st.session_state["parsed_data"] = _flatten_for_storage(inv_records)

                                st.rerun()

                            except Exception as e:

                                st.error(f"❌ 解析失敗：{e}")

                    with col_btn_u2:

                        if st.button("🗑️", use_container_width=True, key="upload_clear_btn"):

                            st.session_state.pop("parsed_data", None)

                            st.rerun()



        # === 操作按鈕區（在 tabs 之外，無論哪個分頁都可見）===

        parsed = st.session_state.get('parsed_data')

        if parsed is not None:

            st.markdown("---")

            col_btn1, col_btn2 = st.columns([3, 1])

            with col_btn1:

                if st.button("💾 儲存資料", use_container_width=True, type="primary", key="save_data_btn"):

                    try:

                        conn = get_db()

                        result = insert_sales_records(parsed, "admin", {})

                        conn.commit()

                        conn.close()

                        st.session_state["last_result"] = result

                        st.session_state.pop("parsed_data", None)

                        st.rerun()

                    except Exception as e:

                        st.error(f"❌ 儲存失敗：{e}")

            with col_btn2:

                if st.button("🔄 重貼", use_container_width=True, key="reinput_btn"):

                    st.session_state.pop("parsed_data", None)

                    st.rerun()



        # 顯示上次儲存結果（無論哪個分頁都可見）

        if st.session_state.get("last_result"):

            lr = st.session_state["last_result"]

            st.success(f"✅ 寫入完成！新增 {lr['new']} 筆，跳過重複 {lr['skipped']} 筆。")

            st.session_state.pop("last_result", None)



    # === 右側預覽區域 ===

    with col_preview:

        parsed = st.session_state.get('parsed_data')

        if not parsed:

            st.info("👈 貼上資料後點擊「預覽資料」查看解析結果。")

        else:

            st.subheader(f"📊 預覽結果：共 {len(parsed)} 筆記錄")



            # 摘要統計卡片

            dists = set()

            total_amt = 0.0

            for r in parsed:

                d = r.get("distributor", "")

                if d:

                    dists.add(d)

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

                    "經銷商": r.get("distributor",""), "備註": r.get("order_note","")})

            df = pd.DataFrame(rows)

            if "金額" in df.columns:

                df["金額"] = df["金額"].apply(lambda x: f"{x:,.0f}" if isinstance(x, (int,float)) else "")

            st.dataframe(df, use_container_width=True, hide_index=True, height=450)









def render_dashboard():
    """儀表板：銷售分析圖表，同時顯示牌價與實價。"""
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
    
    # 一次性查詢所有資料（同時包含牌價與實價）
    try:
        conn2 = get_db()
        dashboard_data = get_dashboard_data(conn2, selected_month)
        conn2.close()
    except Exception as e:
        st.error(f"讀取資料失敗：{e}")
        import traceback; st.exception(e)
        return
    
    daily_data = dashboard_data["daily"]
    dist_data = dashboard_data["distributors"]
    cat_data = dashboard_data["categories"]
    stats = dashboard_data["stats"]
    
    # 統計卡片
    c1, c2, c3, c4 = st.columns(4)
    tl = stats.get("total_lines", 0)
    tc = stats.get("invoice_count", 0)
    tb = stats.get("total_brand", 0)
    ta = stats.get("total_actual", 0)
    with c1: st.metric("明細筆數", f"{tl:,}")
    with c2: st.metric("銷貨單張數", f"{tc:,}")
    with c3: st.metric("牌價總額", f"NT${tb:,.0f}")
    with c4: st.metric("實價總額", f"NT${ta:,.0f}")
    
    # ====== 每日銷售趨勢：雙線圖 ======
    st.subheader("每日銷售趨勢")
    if daily_data:
        import plotly.graph_objects as go
        dates = [d["date"] for d in daily_data]
        brand_vals = [d["brand"] for d in daily_data]
        actual_vals = [d["actual"] for d in daily_data]
    
        fig1 = go.Figure()
        fig1.add_trace(go.Scatter(x=dates, y=brand_vals, mode="lines+markers",
            name='牌價', line=dict(color='#3b82f6', width=2.5), marker=dict(size=6)))
        fig1.add_trace(go.Scatter(x=dates, y=actual_vals, mode="lines+markers",
            name='實價', line=dict(color='#4a6741', width=2.5), marker=dict(size=6)))
        fig1.update_layout(
            yaxis_title="金額 (NT$)", xaxis_title=None,
            height=380, showlegend=True, legend=dict(yanchor='top', y=0.95, xanchor='left', x=0.02),
        )
        st.plotly_chart(fig1, use_container_width=True)
    
        with st.expander('每日明細表格'):
            rows = []
            for d in daily_data:
                rows.append({"日期": d["date"], "牌價": f'{d["brand"]:,.0f}', "實價": f'{d["actual"]:,.0f}'})
            st.dataframe(rows, use_container_width=True, hide_index=True)
    else:
        st.info(f"{selected_month} 尚無銷售資料。")
    
# ====== 經銷商銷售排名：水平柱狀圖（全部列出）======
    st.subheader("經銷商銷售排名")
    if dist_data:
        import plotly.graph_objects as go

             # 顯示全部經銷商，依牌價由大至小排列
        sorted_dist = list(reversed(dist_data))
        names = [d["distributor"] for d in sorted_dist]
        brand_bars = [d["brand"] for d in sorted_dist]
        actual_bars = [d["actual"] for d in sorted_dist]

        fig2 = go.Figure()
        fig2.add_trace(go.Bar(
             y=names, x=brand_bars, orientation='h',
            customdata=[f"牌價: NT${b:,.0f}<br>實價: NT${a:,.0f}" for b, a in zip(brand_bars, actual_bars)],
            hovertemplate='%{y}<br>%{customdata}<extra></extra>',
            marker_color='#4a6741',
           ))
        fig2.update_layout(
            yaxis_title=None, xaxis_title="銷售牌價 (NT$)",
            height=max(300, len(sorted_dist)*28+60), showlegend=False,
            margin=dict(l=90, r=10, t=20, b=40),
            yaxis=dict(tickfont=dict(size=12)),
           )
        st.plotly_chart(fig2, use_container_width=True)

        with st.expander('完整排名表格'):
            df_r = pd.DataFrame(dist_data)
            df_r.insert(0, "排名", range(1, len(df_r)+1))
            df_r["牌價"] = df_r["brand"].apply(lambda x: f"{x:,.0f}")
            df_r["實價"] = df_r["actual"].apply(lambda x: f"{x:,.0f}")
            df_r = df_r[["排名", "distributor", "牌價", "實價", "invoice_count"]]
            df_r.columns = ["排名", "經銷商", "牌價", "實價", "銷貨單張數"]
            st.dataframe(df_r, use_container_width=True, hide_index=True)
    else:
        st.info(f"{selected_month} 尚無經銷商資料。")
    
    # ====== 產品類別佔比：雙圓餅圖 ======
    st.subheader("產品類別佔比分析")
    if cat_data:
        import plotly.graph_objects as go
        from plotly.subplots import make_subplots
        colors_list = ["#4a6741", "#6b8f5e", "#8fb572", "#c5d9b3", "#ddd", "#eee"]
    
        total_brand_c = sum(d["brand"] for d in cat_data) or 1
        total_actual_c = sum(d["actual"] for d in cat_data) or 1
    
        cats = [d["category"] for d in cat_data]
        brand_vals = [d["brand"] for d in cat_data]
        actual_vals = [d["actual"] for d in cat_data]
    
        fig3 = make_subplots(rows=1, cols=2,
            subplot_titles=('牌價佔比', '實價佔比'),
            specs=[[{'type':'pie'}, {'type':'pie'}]])
    
        fig3.add_trace(go.Pie(
             labels=cats, values=brand_vals, hole=0.5,
            marker_colors=colors_list[:len(cats)],
            textinfo='percent+label', textposition='outside',
        ), row=1, col=1)
    
        fig3.add_trace(go.Pie(
             labels=cats, values=actual_vals, hole=0.5,
            marker_colors=colors_list[:len(cats)],
            textinfo='percent+label', textposition='outside',
        ), row=1, col=2)
    
        fig3.update_layout(height=420, showlegend=True)
        st.plotly_chart(fig3, use_container_width=True)
    
        with st.expander('類別明細表格'):
            rows = []
            for d in cat_data:
                rows.append({
                     "類別": d["category"],
                     '牌價': f'{d["brand"]:,.0f}',
                     '實價': f'{d["actual"]:,.0f}',
                     '牌價佔比': f'{d["brand"]/total_brand_c*100:.1f}%',
                     '實價佔比': f'{d["actual"]/total_actual_c*100:.1f}%',
                 })
            st.dataframe(rows, use_container_width=True, hide_index=True)
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

                     # Handle UploadedFile by saving to temp path first
                    import tempfile
                    tmp_path = os.path.join(tempfile.gettempdir(), up_file.name)
                    with open(tmp_path, 'wb') as tmp_f:
                         tmp_f.write(up_file.read())
                    result = import_from_excel(tmp_path)
                    os.remove(tmp_path)

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

        sql_records = "SELECT sale_date, invoice_no, cust_name, product_code, category, amount, COALESCE(order_note,'') FROM sales_records WHERE accounting_month = ? AND distributor = ? ORDER BY sale_date, invoice_no"

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

    inv_records = defaultdict(lambda: {"date": None, "cust_name": "", "order_note": "", "cats": {}})

    for rec in records:

        inv_no = rec[1]

        cat = rec[4] or "其他"

        amt = rec[5] or 0.0

        r = inv_records[inv_no]

        if not r["date"]:

            r["date"] = rec[0]

            r["cust_name"] = rec[2] or ""

            if not r["order_note"] and rec[6]:

                r["order_note"] = rec[6]

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



    # Build detail_rows early (needed for PDF export)

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

               "order_note": r["order_note"],
              "小計": f"{inv_total:,.0f}",

          })



    # Parse month for display

    y_str, m_str = selected_month.split("-")

    rocy_year = int(y_str) - 1911

    month_title = f"{rocy_year}年{int(m_str):02d}月"



# === Header card with PDF buttons on right side ===

    left_hdr, right_hdr = st.columns([3, 1])



    with left_hdr:

        h = []

        h.append('<div style="background:linear-gradient(135deg,#4a6741 0%,#6b8f5e 100%);color:white;padding:1rem 1.5rem;border-radius:12px;margin-bottom:0.5rem;">')

        h.append('<div style="display:flex;justify-content:space-between;align-items:center;">')

        h.append('<div><div style="font-size:0.85rem;opacity:0.9;">' + month_title + ' 對帳單</div>')

        h.append('<div style="font-size:1.3rem;font-weight:700;margin-top:0.25rem;">' + selected_dist + '</div></div>')

        h.append('<div style="text-align:right;"><div style="font-size:0.85rem;opacity:0.9;">共 ' + str(len(inv_records)) + ' 筆銷貨</div></div>')

        h.append('</div></div>')

        st.markdown("".join(h), unsafe_allow_html=True)



    with right_hdr:

        # 1. 單廠 PDF 匯出彈窗

        with st.popover("📄 匯出本經銷商", use_container_width=True):

            st.caption(f"廠商：{selected_dist}")

            if st.button("確認生成 PDF", key=f"btn_gen_s_{selected_dist}_{selected_month}", type="primary", use_container_width=True):

                with st.spinner("正在打包 PDF..."):

                    try:

                        pdf_bytes = _generate_single_statement_pdf(

                            selected_dist, selected_month, month_title,

                            rates, total_brand, total_actual, grand_brand, grand_actual,

                            cat_labels, detail_rows

                        )

                        st.download_button(

                            label="⬇️ 下載 PDF",

                            data=pdf_bytes,

                            file_name=f"{month_title}_{selected_dist}.pdf",

                            mime="application/pdf",

                            type="primary",

                            key=f"dl_s_{selected_dist}_{selected_month}",

                            use_container_width=True

                        )

                    except Exception as e:

                        st.error(f"生成失敗：{e}")



        # 2. 全廠 PDF 匯出彈窗

        with st.popover("📑 匯出全部經銷商", use_container_width=True):

            st.caption(f"月份：{month_title} 全部廠商")

            if st.button("確認生成全部", key=f"btn_gen_a_{selected_month}", type="primary", use_container_width=True):

                with st.spinner("正在處理全部廠商，請稍候..."):

                    try:

                        all_pdf_bytes = _generate_all_statements_pdf(selected_month, month_title)

                        st.download_button(

                            label="⬇️ 下載全部 PDF",

                            data=all_pdf_bytes,

                            file_name=f"{month_title}_全部經銷商.pdf",

                            mime="application/pdf",

                            type="primary",

                            key=f"dl_a_{selected_month}",

                            use_container_width=True

                        )

                    except Exception as e:

                        st.error(f"生成失敗：{e}")



    # Notes section below header row

    current_note = get_distributor_note(conn, selected_dist, selected_month)

    new_note = st.text_area('📝 廠商備註（僅內部使用，不會出現在輸出列印）',

        value=current_note, height=40,

        key=f'note_{selected_dist}_{selected_month}',

        label_visibility='collapsed')

    col_n1, col_n2 = st.columns([1, 5])

    with col_n1:

        if st.button('💾 儲存備註', key=f'save_note_{selected_dist}_{selected_month}'):

            save_distributor_note(conn, selected_dist, selected_month, new_note)

            st.success('✅ 已儲存')



    st.divider()



    # === Discount rate cards ===

    _build_dist_rate_cards(r_dj, r_fj, r_dn, r_fn, r_otc)



    st.divider()



    # === Summary table ===

    st.subheader("類別合計對照")

    _render_summary_table(total_brand, total_actual, grand_brand, grand_actual, cat_labels)



    st.divider()





    # === Detail table ===

    st.subheader(f"明細列表（{len(inv_records)} 筆）")

    if detail_rows:

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
    # Column headers including 備註
    all_cols = ["日期","銷貨單號","客戶名稱","單斤","複斤","單濃","複濃","O.T.C.","其他","備註","小計"]
    col_aligns = ["center","center","left","right","right","right","right","right","right","left","right"]
    col_widths = {"日期":"8%","銷貨單號":"10%","客戶名稱":"12%","單斤":"7%","複斤":"7%","單濃":"7%","複濃":"7%","O.T.C.":"7%","其他":"7%","備註":"14%","小計":"10%"}

    d_hdr = "<tr>" + "".join(f"<th style='padding:6px;text-align:center;background:#4a6741;color:white;white-space:nowrap;width:{col_widths.get(h,'auto')};'>{h}</th>" for h in all_cols) + "</tr>"

    d_body = ""

    for idx, dr in enumerate(detail_rows):

        bg = "#ffffff" if idx % 2 == 0 else "#f9fafb"

        vals = []
        for vc in all_cols:
            if vc == "備註":
                vals.append(dr.get("order_note", "") or "")
            else:
                vals.append(dr.get(vc, ""))

         # Color-code notes: show them with a subtle highlight
        note_val = dr.get("order_note", "") or ""

        d_row_html = "<tr style='background:" + bg + "'>"
        for ci, (v, a) in enumerate(zip(vals, col_aligns)):
            if ci == 9:  # 備註 column gets special styling
                note_color = "#f59e0b" if v else ""
                d_row_html += f"<td style='padding:5px;text-align:{a};color:{note_color};font-size:0.8rem;max-width:14%;overflow:hidden;text-overflow:ellipsis;'>{v if v else ''}</td>"
            else:
                d_row_html += f"<td style='padding:5px;text-align:{a};'>{v}</td>"
        d_row_html += "</tr>"
        d_body += d_row_html

    # Total row with blank for 備註
    t_vals = ["", "", "合 計"] + [f"{total_brand[c]:,.0f}" for c in ["單斤","複斤","單濃","複濃","O.T.C."]] + [f"{total_brand['其他']:,.0f}", "", f"{grand_brand:,.0f}"]

    t_cells = ""

    for ti, tv in enumerate(t_vals):

        al = col_aligns[ti] if ti < len(col_aligns) else "right"

        t_cells += f"<td style='padding:6px;text-align:{al};background:#d1fae5;color:#065f46;font-weight:700;'>{tv}</td>"

    d_body += f"<tr>{t_cells}</tr>"

    html = '<div style="border-radius:8px;overflow:hidden;border:1px solid #e5e7eb;"><table style="width:100%;border-collapse:collapse;font-size:0.85rem;table-layout:fixed;">' + d_hdr + d_body + '</table></div>'

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











def _build_statement_pdf_html(dist_name, month_title, rates, total_brand, total_actual, grand_brand, grand_actual, cat_labels, detail_rows):

    """Build HTML for distributor statement PDF with proper column widths."""

    # Determine font path

    proj_font = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts", "Arial Unicode.ttf")

    sys_font = "/Library/Fonts/Arial Unicode.ttf"

    fp = ""

    if os.path.exists(proj_font):

        fp = proj_font

    elif os.path.exists(sys_font):

        fp = sys_font



    font_css = ''

    if fp:

        font_css = f'@font-face {{ font-family: "CJKFont"; src: url("{fp}") format("truetype"); }}'



    from datetime import datetime as _dt_now

    today_str = _dt_now.now().strftime('%Y/%m/%d')



    html = []



    # CSS Styles - optimized for xhtml2pdf (no flexbox, explicit widths)

    html.append('<style>')

    html.append('@page { size: A4; margin: 15mm 12mm 18mm 12mm; }')

    html.append(font_css)

    html.append('body { font-family: "CJKFont", "Arial Unicode", sans-serif; font-size: 9pt; color: #1f2937; margin: 0; padding: 0; }')

    html.append('.company-header { text-align: center; margin-bottom: 4mm; padding-bottom: 3mm; border-bottom: 2px solid #4a6741; }')

    html.append('.co-name { font-size: 14pt; font-weight: bold; color: #4a6741; letter-spacing: 2px; }')

    html.append('.co-sub { font-size: 9pt; color: #6b7280; margin-top: 2pt; }')

    html.append('.dist-info-table { width: 100%; margin-bottom: 3mm; border-collapse: collapse; }')

    html.append('.dist-info-cell { padding: 4pt 6pt; vertical-align: middle; }')

    html.append('.dist-name { font-size: 12pt; font-weight: bold; color: #1f2937; }')

    html.append('.dist-date { text-align: right; font-size: 8.5pt; color: #6b7280; white-space: nowrap; }')

    html.append('.rate-table { width: 100%; margin-bottom: 3mm; border-collapse: separate; border-spacing: 0; }')

    html.append('.rate-cell { border: 0.5pt solid #d1d5db; padding: 6pt 4pt; background: #fafafa; font-size: 8.5pt; text-align: center; vertical-align: middle; box-sizing: border-box; }')

    html.append('.rate-label { color: #6b7280; font-size: 8pt; margin-bottom: 1pt; line-height: 1.3; }')

    html.append('.rate-value { font-weight: bold; font-size: 9.5pt; line-height: 1.3; }')

    html.append('.section-title { font-size: 10pt; font-weight: bold; color: #4a6741; padding: 2pt 0 2pt 5pt; border-left: 3px solid #4a6741; margin-bottom: 2mm; margin-top: 4mm; }')

    html.append('table.stmt-table { width: 100%; border-collapse: collapse; font-size: 8.5pt; margin-bottom: 3mm; empty-cells: show; }')

    html.append('table.stmt-table th { background-color: #4a6741; color: white; padding: 5pt 3pt; text-align: center; font-size: 8.5pt; font-weight: bold; border: 0.5pt solid #9ca3af; line-height: 1.2; }')

    html.append('table.stmt-table td { padding: 5pt 3pt; border: 0.5pt solid #d1d5db; font-size: 8.5pt; line-height: 1.2; vertical-align: middle; }')

    html.append('table.stmt-table tr:nth-child(even) td { background-color: #f9fafb; }')

    html.append('.right { text-align: right; white-space: nowrap; }')

    html.append('.center { text-align: center; white-space: nowrap; }')

    html.append('.left { text-align: left; white-space: nowrap; }')

    html.append('.total-row td { background-color: #d1fae5 !important; color: #065f46 !important; font-weight: bold !important; border-top: 2pt solid #4a6741 !important; padding: 5pt 3pt !important; }')

    html.append('.footer { margin-top: 8mm; padding-top: 3mm; border-top: 1.5pt solid #e5e7eb; font-size: 7pt; color: #9ca3af; text-align: center; }')

    html.append('</style>')



    # Company Header

    html.append('<div class="company-header">')

    html.append('        <div class="co-name">中草藥 GMP 製造廠</div>')

    html.append(f'        <div class="co-sub">{month_title} 經銷商銷售對帳單 | 列印日期：{today_str}</div>')

    html.append('</div>')



    # Distributor Info Row (table-based)

    html.append('<table class="dist-info-table"><tr>')

    html.append(f'        <td class="dist-info-cell dist-name" style="width:60%"> {dist_name}</td>')

    html.append(f'        <td class="dist-info-cell dist-date" style="width:40%"> 計算區間：{month_title} | 列印：{today_str}</td>')

    html.append('</tr></table>')



    # Discount Rates (table-based) - display in a single row

    rate_items = [("單斤", rates.get("單斤", 0.8)), ("複斤", rates.get("複斤", 0.5)),

                  ("單濃", rates.get("單濃", 0.5)), ("複濃", rates.get("複濃", 0.45)),

                  ("O.T.C.", rates.get("O.T.C.", 0.5))]

    cmap = {"單斤": "#10b981", "複斤": "#3b82f6", "單濃": "#f59e0b", "複濃": "#ef4444", "O.T.C.": "#8b5cf6"}

    html.append('<table class="rate-table"><tr>')

    for label, rval in rate_items:

        color = cmap.get(label, "#6b7280")

        pct_str = f"{rval*100:.0f}%"

        html.append(f'        <td class="rate-cell" style="width:20%"><div class="rate-label">{label}</div><div class="rate-value" style="color:{color}">{pct_str}</div></td>')

    html.append('</tr></table>')



    # Summary Table (類別合計對照) - explicit widths per cell

    cats = ["單斤", "複斤", "單濃", "複濃", "O.T.C.", "其他"]

    # Use mm-based widths for summary table: 項目=22mm, each category ~16mm

    sum_widths_pct = [18, 14, 14, 14, 15, 15, 14]  # totals 100%



    html.append('<div class="section-title">類別合計對照</div>')

    html.append('<table class="stmt-table" style="table-layout:fixed;"><tr>')

    for ci, h in enumerate(["項目"] + cats):

        html.append(f'            <th style="width:{sum_widths_pct[ci]}%"> {h}</th>')

    html.append('</tr>')



    # Brand price row

    html.append('<tr>')

    html.append(f'            <td class="left" style="width:{sum_widths_pct[0]}%;font-weight:bold">牌價</td>')

    for ci, c in enumerate(cats):

        html.append(f'            <td class="right" style="width:{sum_widths_pct[ci+1]}%">{total_brand.get(c, 0):,.0f}</td>')

    html.append('</tr>')



    # Actual price row

    html.append('<tr style="background:#f0fdf4">')

    html.append(f'            <td class="left" style="width:{sum_widths_pct[0]}%;font-weight:bold">實價</td>')

    for ci, c in enumerate(cats):

        html.append(f'            <td class="right" style="width:{sum_widths_pct[ci+1]}%">{total_actual.get(c, 0):,.0f}</td>')

    html.append('</tr>')



    # Receivable row (only last column has value)

    html.append('<tr class="total-row">')

    html.append(f'            <td class="left" style="width:{sum_widths_pct[0]}%">應收帳款</td>')

    for ci in range(len(cats)):

        if ci == len(cats) - 1:

            html.append(f'            <td class="right" style="width:{sum_widths_pct[ci+1]}%">{grand_actual:,.0f}</td>')

        else:

            html.append(f'            <td class="right" style="width:{sum_widths_pct[ci+1]}%">-</td>')

    html.append('</tr>')

    html.append('</table>')



    # Detail Table (明細列表) - explicit widths per cell, wider for text columns

    if detail_rows:

        n_detail = len(detail_rows)

        html.append(f'<div class="section-title">明細列表（共 {n_detail} 筆）</div>')

        d_cols = ["日期", "銷貨單號", "客戶名稱", "單斤", "複斤", "單濃", "複濃", "O.T.C.", "其他", "小計"]

        # Widths: date=10%, invoice=18%, customer=20%, each cat=6%, subtotal=12% => 10+18+20+36+12=96, adjust: 10+18+20+7*6+12=100

        d_widths = [10, 18, 20, 6, 6, 6, 6, 6, 6, 12]

        d_aligns = ['center', 'center', 'left'] + ['right'] * 7



        html.append('<table class="stmt-table" style="table-layout:fixed;"><tr>')

        for dw, h in zip(d_widths, d_cols):

            html.append(f'            <th style="width:{dw}%"> {h}</th>')

        html.append('</tr>')



        for dr in detail_rows:

            vals = [str(dr.get("日期", "")), str(dr.get("銷貨單號", "")), str(dr.get("客戶名稱", ""))]

            vals += [dr.get(c, "0") for c in ["單斤", "複斤", "單濃", "複濃", "O.T.C.", "其他"]]

            vals.append(dr.get("小計", "0"))

            html.append('<tr>')

            for vi, v in enumerate(vals):

                w = d_widths[vi]

                cl = d_aligns[vi]

                html.append(f'            <td class="{cl}" style="width:{w}%">{v}</td>')

            html.append('</tr>')



        # Total row

        t_vals = ['', '', '合  計'] + [f'{total_brand.get(c, 0):,.0f}' for c in cats]

        d_total = sum(float(total_brand.get(c, 0)) for c in cats)

        t_vals.append(f'{d_total:,.0f}')

        html.append('<tr class="total-row">')

        for vi, v in enumerate(t_vals):

            w = d_widths[vi]

            cl = d_aligns[vi] if vi < len(d_aligns) else 'right'

            html.append(f'            <td class="{cl}" style="width:{w}%">{v}</td>')

        html.append('</tr>')

        html.append('</table>')



    # Footer

    html.append(f'<div class="footer">本對帳單僅供內部核對使用 | 產生日期：{today_str} | 系統自動產生</div>')



    return "\n".join(html)





def _generate_single_statement_pdf(dist_name, month_key, month_title, rates, total_brand, total_actual, grand_brand, grand_actual, cat_labels, detail_rows):

    """Generate PDF for a single distributor statement."""

    from xhtml2pdf import pisa

    html_str = _build_statement_pdf_html(dist_name, month_title, rates, total_brand, total_actual, grand_brand, grand_actual, cat_labels, detail_rows)

    full = '<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>' + html_str + '</body></html>'

    import io

    pdf_file = io.BytesIO()

    pisa.CreatePDF(full.encode('utf-8'), dest=pdf_file)

    return pdf_file.getvalue()





def _generate_all_statements_pdf(month_key, month_title):

    """Generate PDF for all distributors in a given month."""

    from xhtml2pdf import pisa

    conn = get_db()

    cursor = conn.cursor()



    active_dists = []

    cursor.execute("SELECT DISTINCT distributor FROM sales_records WHERE accounting_month = ? AND distributor IS NOT NULL AND distributor != '' GROUP BY distributor ORDER BY distributor", (month_key,))

    for r in cursor.fetchall():

        active_dists.append(r[0])



    if not active_dists:

        conn.close()

        raise ValueError(f"該月份({month_key})尚無經銷商資料可匯出")



    parts = []

    for dist_name in active_dists:

        cursor.execute("SELECT sale_date, invoice_no, cust_name, product_code, category, amount FROM sales_records WHERE accounting_month = ? AND distributor = ? ORDER BY sale_date, invoice_no", (month_key, dist_name))

        records = cursor.fetchall()



        di = st.session_state.dist_info.get(dist_name, {})

        rats = di.get("discount_rates", {})

        inv_recs = {}

        for rec in records:

            inv_no = rec[1]

            cat = rec[4] or "其他"

            amt = rec[5] or 0.0

            if inv_no not in inv_recs:

                inv_recs[inv_no] = {"date": rec[0], "cust_name": rec[2] or "", "cats": {}}

            inv_recs[inv_no]["cats"][cat] = inv_recs[inv_no]["cats"].get(cat, 0.0) + amt



        cats_all = ["單斤", "複斤", "單濃", "複濃", "O.T.C.", "其他"]

        rmx = {"單斤": rats.get("單斤", 0.8), "複斤": rats.get("複斤", 0.5),

               "單濃": rats.get("單濃", 0.5), "複濃": rats.get("複濃", 0.45),

               "O.T.C.": rats.get("O.T.C.", 0.5), "其他": 1.0}

        tb = {cat: sum(inv_recs[k]["cats"].get(cat, 0) for k in inv_recs) for cat in cats_all}

        ta = {}

        for cat in cats_all:

            if cat != "其他":

                ta[cat] = round(tb[cat] * rmx[cat])

            else:

                ta[cat] = tb[cat]

        gb = sum(tb.values())

        ga = sum(ta.values())



        d_rows = []

        for inv_no in inv_recs:

            r2 = inv_recs[inv_no]

            cats2 = r2["cats"]

            itotal = sum(cats2.values())

            d_rows.append({"日期": str(r2["date"]) if r2["date"] else "", "銷貨單號": inv_no,

                          "客戶名稱": r2["cust_name"],

                          "單斤": f"{cats2.get('單斤', 0):,.0f}", "複斤": f"{cats2.get('複斤', 0):,.0f}",

                          "單濃": f"{cats2.get('單濃', 0):,.0f}", "複濃": f"{cats2.get('複濃', 0):,.0f}",

                          "O.T.C.": f"{cats2.get('O.T.C.', 0):,.0f}", "其他": f"{cats2.get('其他', 0):,.0f}",

                          "小計": f"{itotal:,.0f}"})



        html_str = _build_statement_pdf_html(dist_name, month_title, rats, tb, ta, gb, ga, cats_all, d_rows)

        parts.append(html_str)



    conn.close()



    combined = '<div style="page-break-after:always;">' + '</div><div style="page-break-after:always;">'.join(parts) + '</div>'

    full_html = '<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>' + combined + '</body></html>'

    import io

    pdf_file = io.BytesIO()

    pisa.CreatePDF(full_html.encode('utf-8'), dest=pdf_file)

    return pdf_file.getvalue()

if __name__ == "__main__":

    main()
