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
st.set_page_config(page_title="馬尼通訊 企劃提案系統 v14.3.4", page_icon="🐎", layout="wide")

# CSS 強制美化：左欄背景馬尼藍(#003f7e)，標題馬尼橘(#ef8200)
st.markdown("""
    <style>
    .main { background-color: #F0F2F6; color: #1E2D4A; }
    
    /* 修正引導文字顏色 */
    textarea::placeholder { color: #888888 !important; opacity: 1 !important; }
    
    /* 左側側邊欄視覺 */
    [data-testid="stSidebar"] { background-color: #003f7e !important; border-right: 2px solid #ef8200; }
    [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 { color: #ef8200 !important; font-weight: bold; }
    [data-testid="stSidebar"] .stMarkdown p, [data-testid="stSidebar"] label { color: #FFFFFF !important; }
    
    /* 下拉選單美化 */
    div[data-baseweb="select"] > div { background-color: #FFFFFF !important; color: #003f7e !important; }
    
    /* 章節標題強化 (明顯標題感) */
    .section-header { 
        font-size: 20px !important; 
        color: #003f7e !important; 
        font-weight: 800 !important; 
        margin-top: 10px !important;
        margin-bottom: 5px !important;
        border-left: 5px solid #ef8200;
        padding-left: 10px;
    }
    
    /* AI 按鈕樣式：字體縮小且緊湊 */
    .ai-btn-small>div>button { 
        background-color: #6200EA !important; 
        color: white !important; 
        border: 1px solid #ef8200 !important;
        font-size: 13px !important;
        padding: 2px 10px !important;
        height: auto !important;
        min-height: 30px !important;
    }
    
    /* 建議按鈕字體縮小 */
    .stExpander label p { font-size: 13px !important; color: #666 !important; }
    .stExpander div p { font-size: 13px !important; }
    
    .stButton>button { border-radius: 6px; font-weight: bold; width: 100%; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 分章節 AI 優化邏輯 ---
def section_ai_logic(field_id, text):
    if not text or len(text) < 2: return text
    if field_id == "p_purpose":
        return f"【營運目的優化】本活動核心在於{text}。透過精準時機切入與誘因設計，旨在提升客流並強化品牌高性價比形象。"
    elif field_id == "p_core":
        return f"【核心內容優化】本活動名稱為「{st.session_state.p_name}」，鎖定目標族群需求，建立市場競爭優勢。"
    elif field_id == "p_schedule":
        return f"{text}\n\n💡 AI 執行建議：確保宣傳與銷售期銜接，文宣佈置需提前完成。"
    elif field_id == "p_prizes":
        return f"{text}\n\n💡 AI 配置建議：大獎造勢，小額購物金驅動官網二次轉化。"
    elif field_id == "p_sop":
        return f"{text}\n\n💡 AI SOP 建議：強調「卸下武裝」話術，先聊需求，落實限量管理。"
    elif field_id == "p_marketing":
        return f"🚀【整合行銷】{text}。整合區域廣告與 LINE 官方帳號通知。"
    elif field_id == "p_risk":
        return f"{text}\n\n💡 AI 風險提示：注意稅務申報門檻與序號防偽核對。"
    elif field_id == "p_effect":
        return f"【預期效益優化】{text}。累積潛在客戶名單並提升品牌活躍度。"
    return text

# --- 3. 初始化數據 ---
FIELDS = ["p_name", "p_proposer", "p_purpose", "p_core", "p_schedule", "p_prizes", "p_sop", "p_marketing", "p_risk", "p_effect"]
for field in FIELDS:
    if field not in st.session_state: st.session_state[field] = ""

if 'templates_store' not in st.session_state:
    st.session_state.templates_store = {
        "🐎 馬年慶：百倍奉還 (官方)": {
            "p_name": "2026 馬尼通訊「馬年慶：百倍奉還」企劃案",
            "p_purpose": "迎接馬年，透過 $100 元低門檻吸引新舊客戶進店，增加官網流量。",
            "p_core": "對象：全門市消費者；核心產品：$100 新年禮包。",
            "p_schedule": "115/01/12 宣傳、01/19 販售。",
            "p_prizes": "PS5 | 1名 | 售價 $100 包裝。",
            "p_sop": "確認限購3包、引導加官方 LINE。",
            "p_marketing": "FB/IG 限動倒數、門市完售海報。",
            "p_risk": "稅金申報規範、序號防偽處理。",
            "p_effect": "預期 2,000+ 人流、官網互動提升。"
        }
    }

# --- 4. 側邊欄 ---
with st.sidebar:
    st.header("📋 企劃範本庫")
    selected_tpl = st.selectbox("選擇既有範本", options=list(st.session_state.templates_store.keys()))
    
    c1, c2 = st.columns(2)
    with c1:
        if st.button("📥 載入範本"):
            for k, v in st.session_state.templates_store[selected_tpl].items(): st.session_state[k] = v
            st.rerun()
    with c2:
        if st.button("💾 儲存為範本"):
            name_snip = st.session_state.p_name[:5] if st.session_state.p_name else datetime.now().strftime('%H%M')
            st.session_state.templates_store[f"💾 自訂：{name_snip}..."] = {f: st.session_state[f] for f in FIELDS}
            st.success("已存入庫")

    st.divider()
    if st.button("🗑️ 清空編輯區"):
        for f in FIELDS: st.session_state[f] = ""
        st.rerun()

    st.markdown("<br>"*5, unsafe_allow_html=True)
    with st.expander("🛠️ 系統資訊 v14.3.4", expanded=False):
        st.caption("修正：標題視覺強化、按鈕縮小優化\n馬尼門活動企劃系統 © 2025 Money MKT")

# --- 5. 主要編輯區 ---
st.title("📱 馬尼通訊 行銷企劃提案系統")

t1, t2, t3 = st.columns([2, 1, 1])
with t1: st.text_input("一、 活動名稱", key="p_name", placeholder="例如: 馬年慶：百倍奉還")
with t2: st.text_input("提案人", key="p_proposer", placeholder="行銷部 / 您的姓名")
with t3: st.date_input("提案日期", value=datetime.now(), key="p_date")

st.divider()

# 章節配置
sections = [
    ("p_purpose", "一、 活動時機與目的", "營運目的邏輯建議", "解決連假後人流痛點。", "請輸入活動背景與目的..."),
    ("p_core", "二、 活動核心內容", "賣點配置建議", "產品具備衝動購買力($100)。", "請輸入執行單位、主要商品賣點..."),
    ("p_schedule", "三、 活動時程安排", "執行重點建議", "宣傳期需於除夕前完成。", "115/01/12: 宣傳啟動..."),
    ("p_prizes", "四、 贈品結構與預算", "配置用意建議", "PS5 話題 + 購物金轉化。", "品項 | 數量 | 備註..."),
    ("p_sop", "五、 門市執行 SOP", "執行注意事項", "先卸下武裝不推產品。", "請輸入銷售環節、限量管理與話術..."),
    ("p_marketing", "六、 行銷流程與策略", "建議管道與潤稿", "社群分享好運抽購物金。", "請輸入宣傳管道與標語策略..."),
    ("p_risk", "七、 風險管理與注意事項", "規範與注意建議", "務必收齊身分證影本報稅。", "請輸入稅務、防偽與退場機制..."),
    ("p_effect", "八、 預估成效", "效益面建議", "重點指標：官網註冊數提升。", "預期帶動的人流量或轉化比例...")
]

col_a, col_b = st.columns(2)
for i, (fid, title, tip_title, tip_content, ph_text) in enumerate(sections):
    target_col = col_a if i < 4 else col_b
    with target_col:
        # 1. 章節標題 (強化版標題感)
        st.markdown(f'<p class="section-header">{title}</p>', unsafe_allow_html=True)
        
        # 2. AI 按鈕 (縮小並置於標題與輸入框之間)
        st.markdown('<div class="ai-btn-small">', unsafe_allow_html=True)
        if st.button(f"🪄 AI 優化 {title[:4]}...", key=f"btn_{fid}"):
            st.session_state[fid] = section_ai_logic(fid, st.session_state[fid])
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
        
        # 3. 輸入框 (含引導文)
        st.text_area("", key=fid, height=120, placeholder=ph_text, label_visibility="collapsed")
        
        # 4. 建議區 (縮小感)
        with st.expander(f"💡 查看建議", expanded=False):
            st.caption(f"**{tip_title}:** {tip_content}")
        st.write("")

# --- 6. Word 下載 ---
def generate_word():
    doc = Document()
    doc.add_heading('行銷企劃執行提案書', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_heading(st.session_state.p_name if st.session_state.p_name else "未命名活動", level=1)
    for fid, title, _, _, _ in sections:
        doc.add_heading(title, level=2)
        doc.add_paragraph(st.session_state[fid] if st.session_state[fid] else "（未填寫）")
    word_io = BytesIO(); doc.save(word_io)
    return word_io.getvalue()

st.divider()
if st.session_state.p_name:
    if st.button("✅ 完成企劃並產生文檔"):
        data = generate_word()
        st.download_button(label=f"📥 下載企劃書", data=data, file_name=f"MoneyMKT_{st.session_state.p_name}.docx")
