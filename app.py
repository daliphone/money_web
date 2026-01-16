import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
from io import BytesIO

# --- 1. 頁面配置與 UI (維持清新視覺) ---
st.set_page_config(page_title="馬尼通訊 戰略發想系統 v14.6.0", page_icon="🐎", layout="centered")

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
    .ai-btn-small>div>button { 
        background-color: #6D28D9 !important; color: white !important; 
        font-weight: 800 !important; border-radius: 8px !important;
        height: 42px !important; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 戰略融合配置 ---
MODULES = [
    ("p_purpose", "一、 活動時機與目的", "【戰略目的】去化高壓商品/增加數據資產。AI 會質疑你的目標是否夠具侵略性。"),
    ("p_core", "二、 活動核心內容", "【百倍誘餌】$100換>$500價值？AI 會挑戰你的誘餌吸引力。"),
    ("p_schedule", "三、 活動時程安排", "【執行節奏】含弱勢店面加碼啟動時機。"),
    ("p_prizes", "四、 贈品結構與預算", "【價值槓桿】AI 會提供沒想過的獎項配置（如：無形服務、專屬權利）。"),
    ("p_sop", "五、 門市執行流程 (SOP)", "【心理攻防】破冰第一句話、加購埋伏。AI 會提供反直覺的話術。"),
    ("p_marketing", "六、 行銷流程與策略", "【病毒傳播】霸氣/親民標題。AI 會生成讓顧客忍不住拍照分享的視覺點。"),
    ("p_risk", "七、 風險管理與注意事項", "【資產保護】庫存動態策略與退場機制。"),
    ("p_effect", "八、 預估成效", "【數據漏斗】進店>參與>成交。AI 會檢核數據邏輯是否嚴謹。")
]

FIELDS = [m[0] for m in MODULES] + ["p_name", "p_proposer", "p_date"]

if 'logic_state' not in st.session_state: st.session_state.logic_state = {fid: guide for fid, _, guide in MODULES}
if 'templates_store' not in st.session_state: st.session_state.templates_store = {"馬尼百倍奉還範本": {f: "" for f in FIELDS}}

for f in FIELDS:
    if f not in st.session_state:
        if f == 'p_date': st.session_state[f] = datetime.now()
        else: st.session_state[f] = ""

# --- 3. 側邊欄與版本管理 ---
with st.sidebar:
    st.header("📋 戰略管理中心")
    selected_tpl = st.selectbox("載入企劃範本", options=list(st.session_state.templates_store.keys()))
    if st.button("📥 載入並重置"):
        for k, v in st.session_state.templates_store[selected_tpl].items():
            if k in st.session_state: st.session_state[k] = v
        st.rerun()
    
    st.divider()
    with st.expander("ℹ️ 系統版本資訊"):
        st.caption("v14.6.0: 雙重戰略引擎 (侵略性+創意)")
        edit_mode = st.toggle("🔓 編輯引導邏輯", value=False)

# --- 4. 主要編輯區 ---
st.title("📱 馬尼通訊：雙重戰略發想系統")

b1, b2, b3 = st.columns([2, 1, 1])
with b1: st.text_input("活動名稱", key="p_name")
with b2: st.text_input("提案人", key="p_proposer")
with b3: st.date_input("提案日期", key="p_date")

st.divider()

for fid, title, guide in MODULES:
    st.markdown(f'<p class="section-header">{title}</p>', unsafe_allow_html=True)
    if edit_mode:
        st.session_state.logic_state[fid] = st.text_input(f"編輯邏輯", value=st.session_state.logic_state[fid], key=f"le_{fid}")
    
    st.text_area("", key=fid, height=160, placeholder=st.session_state.logic_state[fid], label_visibility="collapsed")
    
    # 戰略優化按鈕 (對齊調整)
    c_ai, c_tip = st.columns([1, 2.5], vertical_alignment="center") 
    with c_ai:
        st.markdown('<div class="ai-btn-small">', unsafe_allow_html=True)
        if st.button(f"🔥 戰略優化", key=f"btn_{fid}"):
            # 此處未來串接上述雙重引擎 Prompt
            st.session_state[fid] = f"【🔥 戰略摧毀與重建】\n1. 侵略性挑戰：你目前的目標太保守了...\n2. 創意新玩法：考慮結合數位刮刮樂與門市實體任務...\n---\n{st.session_state[fid]}"
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    with c_tip:
        with st.expander("💡 顧問實戰建議", expanded=False):
            st.caption("點擊戰略優化可獲得針對馬尼資源的進階玩法。")

# --- 5. 文檔導出 ---
def generate_word():
    doc = Document()
    doc.add_heading('馬尼通訊 戰略執行提案書 v14.6.0', 0)
    for fid, title, _ in MODULES:
        doc.add_heading(title, level=2)
        doc.add_paragraph(st.session_state[fid] if st.session_state[fid] else "待填寫")
    word_io = BytesIO(); doc.save(word_io); return word_io.getvalue()

st.divider()
if st.session_state.p_name and st.button("✅ 生成戰略文檔"):
    st.download_button(label="📥 下載 docx", data=generate_word(), file_name=f"Strategy_{st.session_state.p_name}.docx")
