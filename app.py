import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
from io import BytesIO
import os

# --- 1. 頁面配置 ---
st.set_page_config(page_title="馬尼通訊 戰略發想系統 v14.6.1", page_icon="🐎", layout="centered")

st.markdown("""
    <style>
    .main { background-color: #F8FAFC; color: #1E293B; }
    [data-testid="stSidebar"] { background-color: #FFFFFF !important; border-right: 1px solid #E2E8F0 !important; }
    .section-header { 
        font-size: 20px !important; color: #003f7e !important; font-weight: 700 !important; 
        margin-top: 35px !important; margin-bottom: 12px !important;
        display: flex; align-items: center;
    }
    .section-header::before {
        content: ""; display: inline-block; width: 5px; height: 24px; 
        background-color: #ef8200; margin-right: 12px; border-radius: 2px;
    }
    .stButton>button { border-radius: 8px !important; font-weight: bold !important; }
    .ai-btn-small>div>button { 
        background-color: #6D28D9 !important; color: white !important; 
        font-size: 13px !important; height: 42px !important;
    }
    .stExpander { border: 1px solid #E2E8F0 !important; border-radius: 8px !important; background-color: white !important; }
    textarea::placeholder { color: #94A3B8 !important; font-style: italic; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 核心邏輯配置 (修復後的 Key 結構) ---
MODULES = [
    ("p_purpose", "一、 活動時機與目的", "【增長目標】去化高壓商品/增加數據資產。為何而戰？解決痛點還是出清庫存？"),
    ("p_core", "二、 活動核心內容", "【百倍誘餌】心理帳戶槓桿（$100換>$500價值）、大獎勾子（百倍價值感）。"),
    ("p_schedule", "三、 活動時程安排", "【執行節奏】含宣傳、銷售、結案期。針對弱勢店面是否有額外即時誘因？"),
    ("p_prizes", "四、 贈品結構與預算", "【價值槓桿】誘餌價值是否大於門檻金額？物資、獎項配置預算平衡。"),
    ("p_sop", "五、 門市執行流程 (SOP)", "【轉化路徑】破冰第一句話要說什麼？加購埋伏、二訪勾子（下次領的贈品）。"),
    ("p_marketing", "六、 行銷流程與策略", "【病毒傳播】霸氣/親民型標題、社群短文案、宣傳力道分配。"),
    ("p_risk", "七、 風險管理與注意事項", "【資產保護】庫存動態策略、損壞界定、稅務與退場機制。"),
    ("p_effect", "八、 預估成效", "【數據漏斗】進店>參與>成交。名單資產(LINE)累積與質化問卷指標。")
]

FIELDS = [m[0] for m in MODULES] + ["p_name", "p_proposer", "p_date"]

DEFAULT_TIPS = {
    "p_purpose": "核心邏輯：春節紅包議題，解決人流痛點。目標：引導消耗紅包財。",
    "p_core": "實戰建議：定價 $100 具備衝動購買力。機制：買禮包獲得百倍大獎序號。",
    "p_sop": "卸下武裝：『建議先試戴不要買』。破冰：『過年試手氣，中獎直接帶走。』",
    "p_effect": "成效檢核：1.數據漏斗(進店>成交) 2.LINE增粉 3.購買原因調查。"
}

# 確保 Session State 初始化正確，不發生 KeyError
if 'logic_state' not in st.session_state:
    st.session_state.logic_state = {fid: guide for fid, _, guide in MODULES}
if 'tips_state' not in st.session_state:
    st.session_state.tips_state = DEFAULT_TIPS.copy()
if 'templates_store' not in st.session_state:
    st.session_state.templates_store = {"馬尼百倍奉還範本": {f: "" for f in FIELDS}}

for f in FIELDS:
    if f not in st.session_state:
        if f == 'p_date': st.session_state[f] = datetime.now()
        else: st.session_state[f] = ""

# --- 3. 側邊欄 ---
with st.sidebar:
    st.header("📋 戰略管理中心")
    selected_tpl = st.selectbox("載入企劃範本", options=list(st.session_state.templates_store.keys()))
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("📥 載入"):
            data = st.session_state.templates_store[selected_tpl]
            for k, v in data.items():
                if k in st.session_state: st.session_state[k] = v
            st.rerun()
    with col2:
        if st.button("💾 儲存"):
            if st.session_state.p_name:
                st.session_state.templates_store[f"💾 {st.session_state.p_name[:10]}"] = {f: st.session_state[f] for f in FIELDS}
                st.success("儲存成功")

    st.markdown("<br>"*15, unsafe_allow_html=True)
    with st.expander("ℹ️ 系統版本資訊"):
        st.caption("v14.6.1: 修復項目配對錯誤 (KeyError Fix)")
        edit_mode = st.toggle("🔓 開啟邏輯編輯模式", value=False)

# --- 4. 主要編輯區 ---
st.title("📱 馬尼通訊：雙重戰略發想系統")

b1, b2, b3 = st.columns([2, 1, 1])
with b1: st.text_input("活動名稱", key="p_name", placeholder="例如：2026馬年慶百倍奉還")
with b2: st.text_input("提案人", key="p_proposer")
with b3: st.date_input("提案日期", key="p_date")

st.divider()

for fid, title, guide in MODULES:
    st.markdown(f'<p class="section-header">{title}</p>', unsafe_allow_html=True)
    
    if edit_mode:
        st.session_state.logic_state[fid] = st.text_input(f"修改「{title}」引導邏輯", value=st.session_state.logic_state[fid], key=f"edit_logic_{fid}")
    
    # 這裡現在確保了 st.session_state.logic_state[fid] 一定存在
    st.text_area("", key=fid, height=160, placeholder=st.session_state.logic_state[fid], label_visibility="collapsed")
    
    # 按鈕對齊優化
    c_ai, c_tip = st.columns([1, 2.5], vertical_alignment="center") 
    with c_ai:
        st.markdown('<div class="ai-btn-small">', unsafe_allow_html=True)
        if st.button(f"🔥 戰略優化", key=f"btn_{fid}"):
            # 專屬雙重引擎輸出 (侵略性+創意)
            st.session_state[fid] = f"【🔥 戰略摧毀與重建】\n- 侵略性挑戰：分析此項目的邏輯漏洞...\n- 創意新玩法：提供基於馬尼資源的非典型方案...\n---\n{st.session_state[fid]}"
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    
    with c_tip:
        with st.expander("💡 顧問實戰建議", expanded=False):
            if edit_mode:
                st.session_state.tips_state[fid] = st.text_area("編輯建議內容", value=st.session_state.tips_state.get(fid, ""), key=f"edit_tip_{fid}")
            else:
                st.caption(st.session_state.tips_state.get(fid, "點擊戰略優化獲得更多靈感"))
    st.write("")

# --- 5. 文檔產出 ---
def generate_word():
    doc = Document()
    doc.add_heading('馬尼通訊 戰略執行提案書 v14.6.1', 0)
    doc.add_heading(st.session_state.p_name if st.session_state.p_name else "企劃案", level=1)
    for fid, title, _ in MODULES:
        doc.add_heading(title, level=2)
        doc.add_paragraph(st.session_state[fid] if st.session_state[fid] else "（內容待填寫）")
    word_io = BytesIO(); doc.save(word_io); return word_io.getvalue()

st.divider()
if st.session_state.p_name:
    if st.button("✅ 完成企劃並產生文檔"):
        doc_data = generate_word()
        st.download_button(label="📥 下載標準企劃書 (docx)", data=doc_data, file_name=f"MoneyMKT_{st.session_state.p_name}.docx")


