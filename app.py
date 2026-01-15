import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from io import BytesIO
import os
import google.generativeai as genai

# --- 1. 頁面配置與品牌色彩 ---
st.set_page_config(page_title="馬尼通訊 模組化企劃系統 v14.3.5", page_icon="🐎", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #F0F2F6; color: #1E2D4A; }
    textarea::placeholder { color: #888888 !important; opacity: 1 !important; }
    
    /* 左側側邊欄視覺：馬尼藍(#003f7e)與馬尼橘(#ef8200) */
    [data-testid="stSidebar"] { background-color: #003f7e !important; border-right: 2px solid #ef8200; }
    [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 { color: #ef8200 !important; font-weight: bold; }
    [data-testid="stSidebar"] .stMarkdown p, [data-testid="stSidebar"] label { color: #FFFFFF !important; }
    div[data-baseweb="select"] > div { background-color: #FFFFFF !important; color: #003f7e !important; }
    
    /* 章節標題強化 */
    .section-header { 
        font-size: 20px !important; color: #003f7e !important; font-weight: 800 !important; 
        margin-top: 20px !important; margin-bottom: 5px !important;
        border-left: 5px solid #ef8200; padding-left: 10px;
    }
    
    /* AI 按鈕精簡化 */
    .ai-btn-small>div>button { 
        background-color: #6200EA !important; color: white !important; 
        border: 1px solid #ef8200 !important; font-size: 12px !important;
        padding: 2px 8px !important; height: auto !important;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 安全 API 串接與 AI 邏輯 ---
# 支援 GitHub 部署與本地環境安全讀取
api_key = st.secrets.get("GEMINI_API_KEY") or os.getenv("GEMINI_API_KEY")

def call_ai_optimize(field_id, user_text):
    if not api_key or not user_text:
        return f"【模擬優化】{user_text} (請設定 API 金鑰以啟用真實 AI)"
    
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-1.5-flash') # 使用最新的 Flash 模型提升速度
    
    # 針對章節屬性配置 Prompt
    prompts = {
        "p_purpose": f"請以營運邏輯優化以下內容，強調解決痛點(如降低購買門檻)及數據增長，並加入去化商品之目標：{user_text}",
        "p_core": f"請優化此核心內容，強調產品唯一賣點與對象契合度：{user_text}",
        "p_sop": f"請針對此門市 SOP 加入「卸下武裝」話術建議與執行細節：{user_text}",
        "p_effect": f"請將以下成效轉化為具備 O2O 轉換與 UGC 口碑累積的效益描述：{user_text}"
    }
    prompt = prompts.get(field_id, f"請潤色並專業化以下行銷企劃內容：{user_text}")
    
    try:
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        return f"連線錯誤：{str(e)}"

# --- 3. 初始化數據與範本 ---
FIELDS = ["p_name", "p_proposer", "p_purpose", "p_core", "p_schedule", "p_prizes", "p_sop", "p_marketing", "p_risk", "p_effect"]
for field in FIELDS:
    if field not in st.session_state: st.session_state[field] = ""

# --- 4. 側邊欄：範本管理 ---
with st.sidebar:
    st.header("📋 企劃範本管理")
    # 預設範本
    tpl_options = ["請選擇範本", "🐎 馬年慶：百倍奉還", "⌚ 7日智慧手錶試戴"]
    selected_tpl = st.selectbox("載入預設模組", tpl_options)
    
    if st.button("📥 確認載入"):
        if "馬年慶" in selected_tpl:
            st.session_state.p_name = "2026「馬年慶：百倍奉還」"
            st.session_state.p_purpose = "解決連假後人流痛點，透過 $100 門檻去化高壓新年禮包庫存。"
            st.session_state.p_sop = "話術：先聊新年願望。SOP：限購3包、引導加官方LINE。"
        elif "試戴" in selected_tpl:
            st.session_state.p_name = "「先體驗再入手」7日試戴專案"
            st.session_state.p_purpose = "降低高單價智慧手錶購買門檻，解決消費者不適配的擔憂。"
            st.session_state.p_sop = "話術：建議先不要買，戴過才知道。SOP：支付押金、簽署同意書。"
        st.rerun()

    st.divider()
    if st.button("🗑️ 清空所有草稿"):
        for f in FIELDS: st.session_state[f] = ""
        st.rerun()

# --- 5. 主要編輯區 ---
st.title("📱 馬尼通訊 模組化企劃提案系統 v14.3.5")

t1, t2, t3 = st.columns([2, 1, 1])
with t1: st.text_input("一、 活動名稱", key="p_name", placeholder="例如: 某商品銷售目的或是去化高壓商品專案")
with t2: st.text_input("提案人", key="p_proposer", placeholder="行銷部 / 您的姓名")
with t3: st.date_input("提案日期", value=datetime.now(), key="p_date")

st.divider()

# 模組化章節配置
sections = [
    ("p_purpose", "一、 活動時機與目的", "營運目的邏輯", "解決消費痛點、數據增長、增加目標商品銷售或去化高壓商品 。", "請輸入背景，例如：欲去化特定庫存或解決價格門檻..."),
    ("p_core", "二、 活動核心內容", "賣點配置建議", "「低門檻、零風險」誘因，將銷售轉為體驗 [cite: 3, 131]。", "請定義對象、執行單位與唯一賣點..."),
    ("p_schedule", "三、 活動時程安排", "執行重點建議", "確保宣傳期與銷售期銜接，文宣提前佈置 [cite: 13, 178]。", "格式：1/12 宣傳、1/19 銷售..."),
    ("p_prizes", "四、 贈品結構與預算", "配置用意建議", "大獎造話題，小獎/優惠券驅動二次回流 [cite: 41, 55]。", "品項 | 數量 | 預算..."),
    ("p_sop", "五、 門市執行 SOP", "心理戰話術建議", "卸下武裝：「建議不要直接買」，先戴再決定 。", "請輸入起手式、引導路徑與銷售禁語..."),
    ("p_marketing", "六、 行銷宣傳策略", "建議管道與潤稿", "累積真實 UGC 心得作為後續素材 [cite: 7, 46]。", "請輸入宣傳管道與社群分享任務..."),
    ("p_risk", "七、 風險管理與規範", "規範與注意建議", "明確扣款標準（如受損、無法開機）與稅法規範 [cite: 74, 111]。", "請輸入損壞界定、退場機制..."),
    ("p_effect", "八、 預估成效", "效益面建議", "重點指標：O2O 轉換率、潛在名單累積 [cite: 80, 83]。", "預期帶動人流、成交筆數、問卷回流量...")
]

col_a, col_b = st.columns(2)
for i, (fid, title, tip_title, tip_content, ph_text) in enumerate(sections):
    target_col = col_a if i < 4 else col_b
    with target_col:
        st.markdown(f'<p class="section-header">{title}</p>', unsafe_allow_html=True)
        # 輸入框
        st.text_area("", key=fid, height=140, placeholder=ph_text, label_visibility="collapsed")
        # 輔助工具區
        c_ai, c_tip = st.columns([1, 1])
        with c_ai:
            st.markdown('<div class="ai-btn-small">', unsafe_allow_html=True)
            if st.button(f"🪄 AI 優化此模組", key=f"btn_{fid}"):
                st.session_state[fid] = call_ai_optimize(fid, st.session_state[fid])
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
        with c_tip:
            with st.expander("💡 邏輯參考"):
                st.caption(f"**{tip_title}:**\n{tip_content}")

# --- 6. Word 產出 ---
def generate_pro_word():
    doc = Document()
    doc.add_heading('行銷企劃執行提案書', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_heading(st.session_state.p_name if st.session_state.p_name else "未命名活動", level=1)
    for fid, title, _, _, _ in sections:
        doc.add_heading(title, level=2)
        doc.add_paragraph(st.session_state[fid] if st.session_state[fid] else "（未填寫內容）")
    word_io = BytesIO(); doc.save(word_io); return word_io.getvalue()

st.divider()
if st.session_state.p_name:
    if st.button("✅ 完成企劃並產生文檔"):
        doc_data = generate_pro_word()
        st.download_button(label="📥 下載模組化企劃書", data=doc_data, file_name=f"MoneyMKT_{st.session_state.p_name}.docx")
