import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from io import BytesIO
import os

# --- 1. 頁面配置 ---
st.set_page_config(page_title="馬尼通訊 企劃提案系統 v14.3.0", page_icon="🐎", layout="wide")

# CSS 強制修正：左欄背景馬尼藍(#003f7e)，標題馬尼橘(#ef8200)
st.markdown("""
    <style>
    .main { background-color: #F0F2F6; color: #1E2D4A; }
    ::placeholder { color: #888888 !important; opacity: 0.7 !important; }
    
    /* 左側側邊欄視覺修正 */
    [data-testid="stSidebar"] { background-color: #003f7e !important; }
    [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 { color: #ef8200 !important; }
    [data-testid="stSidebar"] .stMarkdown p, [data-testid="stSidebar"] label, [data-testid="stSidebar"] span {
        color: #FFFFFF !important;
    }
    
    /* AI 按鈕樣式 */
    .stButton>button { border-radius: 8px; font-weight: bold; width: 100%; }
    .ai-btn>div>button { background-color: #6200EA !important; color: white !important; border: 1px solid #ef8200 !important; }
    .stDownloadButton>button { background-color: #27AE60; color: white; border-radius: 8px; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 分章節 AI 優化邏輯 ---
def section_ai_logic(field_id, text):
    if not text or len(text) < 2: return text
    
    # 分章節屬性需求配置 [cite: 37]
    if field_id == "p_purpose": # 營運目的邏輯 [cite: 3, 5]
        return f"【營運目的優化】本活動核心在於{text}。透過精準時機切入與誘因設計，旨在提升客流並強化品牌高性價比形象 [cite: 4, 5]。"
    
    elif field_id == "p_core": # 賣點強化 [cite: 7, 10]
        return f"【核心內容優化】名稱為{text}。鎖定目標族群需求，透過差異化服務與核心商品配置建立市場絕對優勢 [cite: 7, 8, 10]。"
    
    elif field_id == "p_schedule": # 執行重點建議 [cite: 11, 12]
        return f"{text}\n\n💡 AI 執行重點：請特別注意宣傳期與銷售期的銜接，確保人員在1/12前完成所有文宣物佈置 [cite: 12, 18]。"
    
    elif field_id == "p_prizes": # 配置用意與賣點 [cite: 13, 17]
        return f"{text}\n\n💡 AI 配置建議：吸睛大獎用於創造流量與話題，小額購物金則用於強制官網引流產生二次消費 [cite: 15, 17, 46]。"
    
    elif field_id == "p_sop": # SOP 注意事項建議 [cite: 18, 19]
        return f"{text}\n\n💡 AI SOP 建議：執行過程應強調『卸下武裝』話術，先聊需求不推產品，並嚴格執行限量管理 [cite: 189, 234]。"
    
    elif field_id == "p_marketing": # 建議管道與潤稿 [cite: 23, 24]
        return f"🚀【整合行銷】{text}。建議同步佈署 FB 區域廣告與 LINE 官方帳號通知，確保觸及最大化 [cite: 45, 58]。"
    
    elif field_id == "p_risk": # 風險建議 [cite: 27, 28]
        return f"{text}\n\n💡 AI 風險提示：務必注意稅務申報門檻(>$1000)與中獎序號的防偽蓋章核對流程 [cite: 28, 31, 40]。"
    
    elif field_id == "p_effect": # 效益面建議 [cite: 33, 34]
        return f"【預期效益優化】{text}。除短期業績外，預計可累積超過2,000筆潛在客戶名單作為未來行銷受眾 [cite: 34, 45]。"
    
    return text

# --- 3. 初始化 Session State ---
FIELDS = ["p_name", "p_proposer", "p_purpose", "p_core", "p_schedule", "p_prizes", "p_sop", "p_marketing", "p_risk", "p_effect"]
for field in FIELDS:
    if field not in st.session_state: st.session_state[field] = ""

if 'templates_store' not in st.session_state:
    st.session_state.templates_store = {
        "🐎 馬年慶：百倍奉還 (官方)": {
            "p_name": "2026 馬尼通訊「馬年慶：百倍奉還」企劃案",
            "p_purpose": "迎接馬年，透過 $100 元低門檻吸引新舊客戶進店，增加官網流量 [cite: 4, 5]。",
            "p_core": "對象：全門市消費者；核心產品：$100 新年禮包 [cite: 8, 10]。",
            "p_schedule": "115/01/12 宣傳、01/19 販售 [cite: 12]。",
            "p_prizes": "PS5 | 1名 | 售價 $100 包裝 [cite: 15]。",
            "p_sop": "確認限購3包、引導加官方 LINE [cite: 19, 22]。",
            "p_marketing": "FB/IG 限動倒數、門市完售海報 [cite: 21, 25]。",
            "p_risk": "稅金申報規範、序號防偽處理 [cite: 28, 31]。",
            "p_effect": "預期 2,000+ 人流、官網互動提升 [cite: 34, 35]。"
        }
    }

# --- 4. 側邊欄 ---
with st.sidebar:
    st.header("📋 企劃範本庫")
    selected_tpl = st.selectbox("選擇範本", options=list(st.session_state.templates_store.keys()))
    
    c1, c2 = st.columns(2)
    with c1:
        if st.button("📥 載入"):
            for k, v in st.session_state.templates_store[selected_tpl].items(): st.session_state[k] = v
            st.rerun()
    with c2:
        if st.button("💾 儲存"):
            new_key = f"💾 自訂：{st.session_state.p_name[:5]}..." if st.session_state.p_name else f"💾 {datetime.now().strftime('%m%d%H%M')}"
            st.session_state.templates_store[new_key] = {f: st.session_state[f] for f in FIELDS}
            st.success("已存入庫")

    st.divider()
    if st.button("🗑️ 清空編輯區"):
        for f in FIELDS: st.session_state[f] = ""
        st.rerun()

    st.markdown("<br>"*5, unsafe_allow_html=True)
    with st.expander("🛠️ 系統資訊 v14.3.0", expanded=False):
        st.caption("馬尼門活動企劃系統 © 2025 Money MKT\n1. AI 分章節獨立配置\n2. 馬尼藍橘品牌色優化")

# --- 5. 主要編輯區 ---
st.title("📱 馬尼通訊 行銷企劃提案系統")

t1, t2, t3 = st.columns([2, 1, 1])
with t1: st.text_input("一、 活動名稱", key="p_name", placeholder="例如: 馬年慶：百倍奉還")
with t2: st.text_input("提案人", key="p_proposer")
with t3: st.date_input("提案日期", value=datetime.now(), key="p_date")

st.divider()

# 定義章節配置與對應建議
sections = [
    ("p_purpose", "活動時機與目的", "營運目的邏輯建議", "迎接馬年話題，解決連假後人流痛點 [cite: 4, 5, 41]。"),
    ("p_core", "二、 活動核心內容", "核心賣點配置建議", "產品具備衝動購買力($100)，適合快速成交 [cite: 10, 52]。"),
    ("p_schedule", "三、 活動時程安排", "執行重點建議", "宣傳期需於除夕前完成，開獎設定於開工後引流回訪 [cite: 11, 12, 39]。"),
    ("p_prizes", "四、 贈品結構與預算", "商品配置用意建議", "PS5 創造話題，購物金強制客戶登入官網產生二次消費 [cite: 15, 17, 46]。"),
    ("p_sop", "五、 門市執行 SOP", "執行環節注意事項", "務必強調『序號正本』為兌獎唯一憑證，先卸下武裝不推產品 [cite: 31, 189, 190]。"),
    ("p_marketing", "六、 行銷流程與策略", "建議管道與潤稿", "利用紅包色視覺，社群任務可設計分享好運抽購物金 [cite: 26, 47, 58]。"),
    ("p_risk", "七、 風險管理與注意事項", "規範與注意建議", "每店配額管理避免跨區落空，務必收齊身分證影本報稅 [cite: 28, 40, 42]。"),
    ("p_effect", "八、 預估成效", "效益面建議", "重點指標：門市進店率、官網註冊數、二次轉化率 [cite: 34, 35, 46]。")
]

col_a, col_b = st.columns(2)
for i, (fid, title, tip_title, tip_content) in enumerate(sections):
    target_col = col_a if i < 4 else col_b
    with target_col:
        st.text_area(title, key=fid, height=120)
        # 佈置 AI 潤稿與建議按鈕
        st.markdown('<div class="ai-btn">', unsafe_allow_html=True)
        if st.button(f"🪄 執行 {title} AI 優化", key=f"btn_{fid}"):
            st.session_state[fid] = section_ai_logic(fid, st.session_state[fid])
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
        with st.expander(f"💡 {tip_title} (馬年慶背景)", expanded=False):
            st.write(tip_content)
        st.write("")

# --- 6. Word 下載 ---
def generate_word():
    doc = Document()
    h = doc.add_heading('行銷企劃執行提案書', 0); h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_heading(st.session_state.p_name, level=1)
    for fid, title, _, _ in sections:
        doc.add_heading(title, level=2)
        doc.add_paragraph(st.session_state[fid] if st.session_state[fid] else "（未填寫）")
    word_io = BytesIO(); doc.save(word_io)
    return word_io.getvalue()

st.divider()
if st.session_state.p_name:
    if st.button("✅ 完成企劃並產生文檔"):
        data = generate_word()
        st.download_button(label=f"📥 下載 {st.session_state.p_name} 企劃書", data=data, file_name=f"MoneyMKT_{st.session_state.p_name}.docx")
