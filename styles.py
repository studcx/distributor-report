# -*- coding: utf-8 -*-
"""
styles.py — CSS 樣式定義與 UI 工具函式
==========================================
所有視覺相關的 CSS 集中在這裡。
未來要調整配色、字型、間距等，只需修改本檔案。

設計理念：
   - 卡片式佈局，資訊分層清晰
   - 溫暖專業色調（適合傳統中藥企業形象）
   - 進度狀態明確，操作回饋即時
"""


# ============================================================
# 主要 CSS 樣式 — 注入整個頁面
# ============================================================
MAIN_CSS = """
/* ===== 全域設定 ===== */
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@300;400;500;700&display=swap');

.main .block-container {
    padding-top: 1rem;
}

/* ===== 標題區域 ===== */
#main-header {
    text-align: center;
    padding: 2rem 2rem 1.5rem;
    margin-bottom: 2rem;
    background: linear-gradient(135deg, #4a6741 0%, #6b8f5e 50%, #8fb572 100%);
    border-radius: 16px;
    color: white;
    box-shadow: 0 4px 20px rgba(74, 103, 65, 0.3);
}

#main-header h1 {
    font-family: 'Noto Sans TC', sans-serif;
    font-weight: 700;
    font-size: 1.8rem;
    margin-bottom: 0.3rem;
    letter-spacing: 2px;
}

#main-header p {
    font-family: 'Noto Sans TC', sans-serif;
    font-size: 0.95rem;
    opacity: 0.9;
    margin-top: 0.2rem;
}

/* ===== 卡片容器 ===== */
.step-card {
    background: #ffffff;
    border-radius: 12px;
    padding: 1.5rem;
    margin-bottom: 1rem;
    box-shadow: 0 2px 12px rgba(0, 0, 0, 0.06);
    border-left: 4px solid #4a6741;
}

.step-card-info {
    background: #f0f7ff;
    border-left-color: #3b82f6;
}

.step-card-success {
    background: #f0fdf4;
    border-left-color: #22c55e;
}

.step-card-warning {
    background: #fffbeb;
    border-left-color: #f59e0b;
}

.step-card-error {
    background: #fef2f2;
    border-left-color: #ef4444;
}

/* ===== 步驟標題 ===== */
.step-title {
    font-family: 'Noto Sans TC', sans-serif;
    font-size: 1.05rem;
    font-weight: 600;
    color: #374151;
    margin-bottom: 0.8rem;
}

.step-title .step-num {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 28px;
    height: 28px;
    background: #4a6741;
    color: white;
    border-radius: 50%;
    margin-right: 10px;
    font-size: 0.9rem;
}

/* ===== 貼上區域 ===== */
textarea {
    border-radius: 8px !important;
    border: 2px solid #e5e7eb !important;
    font-family: 'Menlo', monospace !important;
    font-size: 0.85rem !important;
}

textarea:focus {
    border-color: #4a6741 !important;
    box-shadow: 0 0 0 3px rgba(74, 103, 65, 0.1) !important;
}

/* ===== 統計數字卡片 ===== */
.metric-card {
    text-align: center;
    padding: 1rem;
    border-radius: 10px;
    background: #f9fafb;
}

.metric-value {
    font-size: 1.5rem;
    font-weight: 700;
    color: #4a6741;
}

.metric-label {
    font-size: 0.8rem;
    color: #6b7280;
    margin-top: 0.3rem;
}

/* ===== 表格優化 ===== */
table {
    border-radius: 8px !important;
    overflow: hidden !important;
}

thead tr th {
    background-color: #4a6741 !important;
    color: white !important;
    font-family: 'Noto Sans TC', sans-serif;
    padding: 0.75rem !important;
}

tbody tr:nth-child(even) {
    background-color: #f9fafb !important;
}

tbody tr:hover {
    background-color: #ecfdf5 !important;
}

/* ===== 按鈕美化 ===== */
.stButton > button {
    border-radius: 8px !important;
    font-family: 'Noto Sans TC', sans-serif;
    font-weight: 500;
    transition: all 0.2s ease !important;
}

/* ===== 分隔線 ===== */
.divider {
    height: 2px;
    background: linear-gradient(to right, #4a6741, transparent);
    border: none;
    margin: 1.5rem 0;
}

/* ===== 狀態徽章 ===== */
.status-badge {
    display: inline-flex;
    align-items: center;
    padding: 4px 12px;
    border-radius: 20px;
    font-size: 0.8rem;
    font-weight: 500;
}

.status-active {
    background: #dcfce7;
    color: #166534;
}

.status-inactive {
    background: #f3f4f6;
    color: #9ca3af;
}

/* ===== 隱藏 Streamlit 預設元素 ===== */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
.header {visibility: hidden;}
"""


# ============================================================
# 工具函式 — 產生 HTML 卡片區塊
# ============================================================

def step_card(title_text, body_html, card_type="default"):
    """產出帶樣式的 HTML 卡片。"""
    type_class = "step-card-{}".format(card_type) if card_type != "default" else ""
    return """
    <div class="step-card {}">
        <div class="step-title">{}</div>
        {}
    </div>
    """.format(type_class, title_text, body_html)


def stat_card(label, value):
    """產出統計數字卡片。"""
    return """
    <div class="metric-card" style="flex:1; min-width: 120px;">
        <div class="metric-value">{}</div>
        <div class="metric-label">{}</div>
    </div>
    """.format(value, label)


def status_badge(text, active=True):
    """產出狀態徽章。"""
    cls = "status-active" if active else "status-inactive"
    return '<span class="status-badge {}">{}</span>'.format(cls, text)


def header_html(title="", subtitle=""):
    """產出頁面頂端標題區 HTML。"""
    return """<div id="main-header"><h1>{}</h1><p>{}</p></div>""".format(
        title, subtitle if subtitle else "")


def stats_row(labels_values):
    """產出橫排統計列。"""
    cards = "".join(stat_card(l, v) for l, v in labels_values)
    return '<div style="display:flex; gap:1rem;">{}</div>'.format(cards)
