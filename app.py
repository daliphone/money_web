import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
from io import BytesIO
import os

# --- 1. 頁面配置與 UI ---
st.set_page_config(page_title="馬尼通訊 營銷發想系統 v14.4.2", page_icon="🐎", layout="centered")

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
    
    /* 強制按鈕與摺疊區塊垂直居中對齊 */
    .stColumn { display: flex; align-items: center; }
    
    /* AI 按鈕精緻化 */
    .stButton>button { 
        width: 100% !important; 
        border-radius: 8px !important;
        height: 45px !important; 
        font-weight: bold !important;
    }
    .ai-btn-small>div>button { 
        background-color: #F5F3FF !important; color: #6D28D9 !important; 
        border: 1px solid #DDD6FE !important; font-size: 13px !important;
    }
    
    /* 摺疊區塊樣式對齊 */
    .stExpander { border: 1px solid #E2E8F0 !important; border-radius: 8px !important; }
    
    textarea::placeholder { color: #94A3B8 !important; font-style: italic; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 初始化 Session State ---
MODULES = [
    ("step1_goal", "第一步：增長目標區（確定為何而戰）", "活動類型：流量型/轉化型/去化型。核心 KPI：預計售出數、毛利額、會員增長數。"),
    ("step2_bait", "第二步：誘餌設計（確定如何引流）", "心理帳戶槓桿：$100能換>$500價值？大獎勾子：如何營造夢幻百倍價值感？"),
    ("step3_path", "第三步：轉化路徑 (Path Optimization)", "破冰第一句話？二訪機制（下次領的贈品）？加購埋伏（推銷哪類高庫存商品）？"),
    ("step4_inventory", "第四步：庫存動態策略 (Inventory Strategy)", "主推庫存商品清單？弱勢店加碼方案（額外的即時誘因）？"),
    ("step5_headline", "第五步：溝通標題（確定宣傳力道）", "霸氣型標題（大獎價值）、親民型標題（低門檻）、社群短文案。"),
    ("step6_metrics", "第六步：資源預算與成效（漏斗化指標）", "人力配置、物資、漏斗轉換預估(進店>參與>成交)、數據資產(LINE好友)、質化指標。")
]

FIELDS = [m[0] for m in MODULES] + ["p_name", "p_proposer", "p_date"]

# 預設建議
DEFAULT_TIPS = {
    "step1_goal": "核心邏輯：若是為了去化，KPI 應設定為『庫存周轉率』而非單純業績。",
    "step2_bait": "實戰建議：利用『紅包感』降低支付痛苦，提升參與率。",
    "step3_path": "破冰話術：『這張抽獎券是送您的，要不要試試手氣？』",
    "step6_metrics": "成效檢核：務必包含『數據資產累積』，例如蒐集到的問卷數量。"
}

if 'p_date' not in st.session_state: st.session_state.p_date = datetime.now()
if 'logic_state' not in st.session_state: st.session_state.logic_state = {m[0]: m[ guide] for m, _, guide in zip(MODULES, [None]*6, [m[2] for m in MODULES])}
if 'tips_state' not in st.session_state: st.session_state.tips_state = DEFAULT_TIPS.copy()
if 'templates_store' not in st.session_state: st.session_state.templates_store = {"請選擇範本": {f: "" for f in FIELDS}}

for f in FIELDS:
    if f not in st.session_state and f != 'p_date': st.session_state[f] = ""

# --- 3. 側邊欄 (遵循 v14.3.9 佈局) ---
with st.sidebar:
    st.header("📋 企劃管理")
    selected_tpl = st.selectbox("選擇既有範本", options=list(st.session_state.templates_store.keys()))
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("📥 載入範本"):
            data = st.session_state.templates_store[selected_tpl]
            for k, v in data.items():
                if k in st.session_state: st.session_state[k] = v
            st.rerun()
    with col2:
        if st.button("💾 儲存範本"):
            if st.session_state.p_name:
                st.session_state.templates_store[f"💾 {st.session_state.p_name[:10]}"] = {f: st.session_state[f] for f in FIELDS}
                st.success("儲存成功")
                st.rerun()

    st.markdown("<br>"*15, unsafe_allow_html=True)
    with st.expander("ℹ️ 系統版本資訊", expanded=False):
        st.caption("v14.4.2: 修復 Widget 衝突與按鈕對齊")
        edit_mode = st.toggle("🔓 開啟引導詞編輯模式", value=False)
        st.write("---")
        st.caption("v14.4.1: 六步發想與 AI 成效檢核")

# --- 4. 主要編輯區 ---
st.title("📱 馬尼通訊 營銷發想系統 v14.4.2")

st.markdown('<p class="section-header">基本提案資訊</p>', unsafe_allow_html=True)
b1, b2, b3 = st.columns([2, 1, 1])
with b1: st.text_input("活動名稱", key="p_name", placeholder="例如：2026馬年慶百倍奉還")
with b2: st.text_input("提案人", key="p_proposer")
with b3: 
    # 修復 Widget 衝突點
    st.date_input("提案日期", key="p_date")

st.divider()

# 直列渲染與水平對齊修復
for fid, title, guide in MODULES:
    st.markdown(f'<p class="section-header">{title}</p>', unsafe_allow_html=True)
    
    if edit_mode:
        st.session_state.logic_state[fid] = st.text_input(f"修改「{title}」提示詞", value=st.session_state.logic_state[fid], key=f"logic_edit_{fid}")
    
    st.text_area("", key=fid, height=160, placeholder=st.session_state.logic_state[fid], label_visibility="collapsed")
    
    # 使用 columns 並設定垂直對齊
    c_ai, c_tip = st.columns([1, 2.5]) 
    with c_ai:
        st.markdown('<div class="ai-btn-small" style="margin-top: 5px;">', unsafe_allow_html=True)
        if st.button(f"🪄 AI 優化檢核", key=f"btn_{fid}"):
            if fid == "step6_metrics":
                st.session_state[fid] = f"【AI 成效診斷】：需包含進店量、轉化率與 LINE 增粉指標。\n---\n{st.session_state[fid]}"
            else:
                st.session_state[fid] = f"【AI 優化建議】針對{title}：\n{st.session_state[fid]}"
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    
    with c_tip:
        # Expander 預設會有一些 margin，我們透過容器對齊
        with st.expander("💡 查看/編輯實戰建議", expanded=False):
            if edit_mode:
                st.session_state.tips_state[fid] = st.text_area("編輯建議", value=st.session_state.tips_state.get(fid, ""), key=f"tip_edit_{fid}")
            else:
                st.caption(st.session_state.tips_state.get(fid, "暫無建議內容"))
    st.write("")

# --- 5. Word 產出 ---
def generate_word():
    doc = Document()
    doc.add_heading('馬尼通訊 營銷執行提案書 v14.4.2', 0)
    doc.add_heading(st.session_state.p_name, level=1)
    for fid, title, _ in MODULES:
        doc.add_heading(title, level=2)
        doc.add_paragraph(st.session_state[fid] if st.session_state[fid] else "（未填寫）")
    word_io = BytesIO(); doc.save(word_io); return word_io.getvalue()

st.divider()
if st.session_state.p_name:
    if st.button("✅ 完成企劃並產生文檔"):
        doc_data = generate_word()
        st.download_button(label="📥 下載企劃書 (docx)", data=doc_data, file_name=f"MoneyMKT_{st.session_state.p_name}.docx")
