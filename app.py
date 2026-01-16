import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from io import BytesIO
import os

# 嘗試匯入 AI 套件，若失敗則進入模擬模式
try:
    import google.generativeai as genai
    HAS_AI_SDK = True
except ImportError:
    HAS_AI_SDK = False

# --- 1. 頁面配置與清新感 UI ---
st.set_page_config(page_title="馬尼通訊 企劃提案系統 v14.3.7", page_icon="🐎", layout="centered")

st.markdown("""
    <style>
    .main { background-color: #F8FAFC; color: #1E293B; }
    
    /* 左側側邊欄：清新白底 */
    [data-testid="stSidebar"] { 
        background-color: #FFFFFF !important; 
        border-right: 1px solid #E2E8F0 !important;
    }
    [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 { color: #003f7e !important; font-weight: 700; }
    [data-testid="stSidebar"] .stMarkdown p, [data-testid="stSidebar"] label { color: #475569 !important; }
    
    /* 章節標題強化 */
    .section-header { 
        font-size: 20px !important; color: #003f7e !important; font-weight: 700 !important; 
        margin-top: 35px !important; margin-bottom: 10px !important;
        display: flex; align-items: center;
    }
    .section-header::before {
        content: ""; display: inline-block; width: 5px; height: 24px; 
        background-color: #ef8200; margin-right: 12px; border-radius: 2px;
    }
    
    /* AI 按鈕精簡化 */
    .ai-btn-small>div>button { 
        background-color: #F5F3FF !important; color: #6D28D9 !important; 
        border: 1px solid #DDD6FE !important; font-size: 12px !important;
        width: auto !important; padding: 2px 15px !important;
    }
    
    textarea::placeholder { color: #94A3B8 !important; font-style: italic; }
    
    /* 分隔線美化 */
    hr { margin-top: 2rem !important; margin-bottom: 2rem !important; border-top: 1px solid #E2E8F0 !important; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 初始化 Session State 與範本庫 ---
FIELDS = ["p_name", "p_proposer", "p_purpose", "p_core", "p_schedule", "p_prizes", "p_sop", "p_marketing", "p_risk", "p_effect"]

if 'templates_store' not in st.session_state:
    st.session_state.templates_store = {
        "請選擇範本": {f: "" for f in FIELDS},
        "🐎 馬年慶：百倍奉還": {
            "p_name": "2026「馬年慶：百倍奉還」",
            "p_purpose": "解決連假後人流痛點，去化新年禮包庫存，達成數據增長目標。",
            "p_core": "對象：全門市。賣點：$100 低門檻試手氣。",
            "p_sop": "核心話術：先聊新年願望。SOP：限購3包、引導加官方LINE。",
            "p_marketing": "FB/IG 視覺紅包標語。",
            "p_risk": "每店配額管理。",
            "p_effect": "預計進店人次 +20%。"
        }
    }

for field in FIELDS:
    if field not in st.session_state: st.session_state[field] = ""

# --- 3. 側邊欄：動態管理 ---
with st.sidebar:
    st.header("📋 企劃管理")
    selected_tpl_key = st.selectbox("選擇既有範本", options=list(st.session_state.templates_store.keys()))
    
    c1, c2 = st.columns(2)
    with c1:
        if st.button("📥 載入範本"):
            data = st.session_state.templates_store[selected_tpl_key]
            for k, v in data.items(): st.session_state[k] = v
            st.rerun()
    with c2:
        if st.button("💾 儲存範本"):
            if st.session_state.p_name:
                new_key = f"💾 {st.session_state.p_name[:10]}"
                st.session_state.templates_store[new_key] = {f: st.session_state[f] for f in FIELDS}
                st.success("自訂範本已儲存")
                st.rerun()
            else:
                st.error("請輸入活動名稱")

    st.divider()
    if st.button("🗑️ 清空編輯區"):
        for f in FIELDS: st.session_state[f] = ""
        st.rerun()

# --- 4. 配置直列章節數據 ---
sections_config = [
    ("p_purpose", "一、 活動時機與目的", "營運目的邏輯：強化解決痛點（如降低購買門檻）與數據增長，增加目標商品銷售或是去化高壓商品。", "核心：春節紅包議題，解決人流痛點。目標：引導消耗紅包財。"),
    ("p_core", "二、 活動核心內容", "賣點配置建議：依據名稱、對象、執行單位、主要賣點，建立「低門檻、零風險」誘因。", "機制：購買禮包獲得序號。定價：$100 具備衝動購買力。"),
    ("p_schedule", "三、 活動時程安排", "執行重點建議：規劃宣傳、銷售、結案期的資源分配。", "時程：1月中旬啟動，確保除夕前銷售完畢。"),
    ("p_prizes", "四、 贈品結構與預算", "配置用意與賣點：平衡大獎話題與小獎導流。", "配置：PS5 (話題) + 現金。購物金用於官網引流。"),
    ("p_sop", "五、 門市執行流程 (SOP)", "執行環節注意事項建議：注入「卸下武裝」策略。", "話術：先聊願望再推「試手氣」。SOP：強調序號正本。"),
    ("p_marketing", "六、 行銷流程與策略", "行銷策略：自動推薦適合管道並生成行銷標語。", "宣傳：紅包視覺，社群任務設計分享好運。"),
    ("p_risk", "七、 風險管理與注意事項", "風險管理建議：針對法務、稅務及產品損毀進行規範。", "風險：每店配額管理。法規：中獎者身份證影本蒐集。"),
    ("p_effect", "八、 預估成效", "成效效益面建議：分析 O2O 轉換、名單累積與問卷數據價值。", "指標：門市進店率、官網註冊數、轉化率。")
]

# --- 5. 主要編輯區 (直列版面) ---
st.title("📱 馬尼通訊 模組化企劃系統 v14.3.7")

# 基本資訊區
st.markdown('<p class="section-header">基本提案資訊</p>', unsafe_allow_html=True)
b1, b2, b3 = st.columns([2, 1, 1])
with b1: st.text_input("活動名稱", key="p_name", placeholder="請輸入本案活動名稱")
with b2: st.text_input("提案人", key="p_proposer")
with b3: st.date_input("提案日期", value=datetime.now(), key="p_date")

st.divider()

# 直列渲染各章節
for fid, title, logic_guide, real_tip in sections_config:
    # 標題
    st.markdown(f'<p class="section-header">{title}</p>', unsafe_allow_html=True)
    
    # 輸入框
    st.text_area("", key=fid, height=150, placeholder=logic_guide, label_visibility="collapsed")
    
    # 功能按鈕區
    c_ai, c_tip = st.columns([1, 4])
    with c_ai:
        if fid in ["p_purpose", "p_core", "p_marketing", "p_risk", "p_effect"]:
            st.markdown('<div class="ai-btn-small">', unsafe_allow_html=True)
            if st.button(f"🪄 AI 優化", key=f"btn_{fid}"):
                st.session_state[fid] = f"【AI 優化中】{st.session_state[fid]}"
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
    with c_tip:
        with st.expander("💡 查看實戰建議", expanded=False):
            st.caption(real_tip)
    st.write("") # 增加章節間距

# --- 6. Word 產出 ---
def generate_word():
    doc = Document()
    doc.add_heading('行銷企劃執行提案書 v14.3.7', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_heading(st.session_state.p_name if st.session_state.p_name else "活動企劃書", level=1)
    for fid, title, _, _ in sections_config:
        doc.add_heading(title, level=2)
        doc.add_paragraph(st.session_state[fid] if st.session_state[fid] else "（未填寫）")
    word_io = BytesIO(); doc.save(word_io); return word_io.getvalue()

st.divider()
if st.session_state.p_name:
    if st.button("✅ 完成企劃並產生文檔"):
        doc_data = generate_word()
        st.download_button(label="📥 下載企劃書", data=doc_data, file_name=f"MoneyMKT_{st.session_state.p_name}.docx")
