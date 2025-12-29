import streamlit as st
import pandas as pd
from icalendar import Calendar, Event
from datetime import datetime, timedelta
from docx import Document
from io import BytesIO

# --- 頁面配置 ---
st.set_page_config(page_title="馬尼通訊 行銷排程系統", page_icon="📱", layout="wide")

# 強制馬尼品牌色風格
st.markdown("""
    <style>
    .main { background-color: #0B1C3F; }
    h1, h2, h3 { color: #FFD700 !important; }
    .stButton>button { background-color: #F39C12; color: white; border-radius: 8px; font-weight: bold; }
    .stDownloadButton>button { background-color: #27AE60; color: white; border-radius: 8px; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

st.title("馬尼通訊 行銷排程系統 v10.0")

# --- 初始化狀態 ---
if 'activity_list' not in st.session_state:
    st.session_state.activity_list = []
if 'modules' not in st.session_state:
    # 預設 8 組空白模組
    st.session_state.modules = [{"name": f"模組 {i+1}", "platform": "門市活動", "s": "", "e": "", "note": "", "spec": ""} for i in range(8)]

# --- 側邊欄：8 組快速模組管理 ---
with st.sidebar:
    st.header("🛠️ 快速模組設定")
    mod_idx = st.selectbox("選擇編輯/載入模組", range(8), format_func=lambda x: st.session_state.modules[x]["name"])
    
    if st.button("💾 將下方編輯區存入此模組"):
        st.session_state.modules[mod_idx] = {
            "name": st.session_state.cur_name,
            "platform": st.session_state.cur_plat,
            "s": st.session_state.cur_s,
            "e": st.session_state.cur_e,
            "note": st.session_state.cur_note,
            "spec": st.session_state.cur_spec
        }
        st.success(f"已儲存至：{st.session_state.cur_name}")

# --- 主要活動編輯區 ---
with st.container():
    col1, col2, col3, col4 = st.columns([2, 1, 1, 1])
    m_data = st.session_state.modules[mod_idx]

    with col1:
        name = st.text_input("活動名稱", value=m_data["name"], key="cur_name")
    with col2:
        platform = st.selectbox("發布平台", ["公司活動(各社群平台)", "門市活動", "自訂"], index=0, key="cur_plat")
    with col3:
        start_date = st.text_input("開始 (MM/DD)", value=m_data["s"], key="cur_s")
    with col4:
        end_date = st.text_input("結束 (MM/DD)", value=m_data["e"], key="cur_e")

    note = st.text_area("活動內容 (支援條列編輯)", height=100, value=m_data["note"], key="cur_note")
    spec = st.text_area("內容規範 (支援條列編輯)", height=150, value=m_data["spec"], key="cur_spec")

    if st.button("➕ 新增至發布清單"):
        if name and start_date and end_date:
            st.session_state.activity_list.append({
                "名稱": name, "平台": platform, "開始": start_date, "結束": end_date, "內容": note, "規範": spec
            })
            st.rerun()

# --- 清單預覽 ---
st.divider()
st.subheader("📋 待匯出活動清單")
if st.session_state.activity_list:
    df = pd.DataFrame(st.session_state.activity_list)
    st.dataframe(df[["名稱", "平台", "開始", "結束"]], use_container_width=True)
    
    if st.button("🗑️ 清空所有清單"):
        st.session_state.activity_list = []
        st.rerun()

# --- 匯出功能 ---
if st.session_state.activity_list:
    st.subheader("📥 產出檔案")
    c_ics, c_word = st.columns(2)

    # 1. 生成 ICS
    cal = Calendar()
    for act in st.session_state.activity_list:
        e = Event()
        e.add('summary', f"[{act['平台']}] {act['名稱']}")
        e.add('description', f"【內容】\n{act['內容']}\n\n【規範】\n{act['規範']}")
        try:
            m1, d1 = map(int, act['開始'].split('/'))
            m2, d2 = map(int, act['結束'].split('/'))
            e.add('dtstart', datetime(2025, m1, d1))
            e.add('dtend', datetime(2025, m2, d2) + timedelta(days=1))
            cal.add_component(e)
        except: continue
    
    with c_ics:
        st.download_button("📅 匯出手機行事曆 (.ics)", data=cal.to_ical(), file_name="馬尼行銷排程.ics", mime="text/calendar")

    # 2. 生成 Word
    doc = Document()
    doc.add_heading('馬尼通訊 行銷活動執行公告', 0)
    for act in st.session_state.activity_list:
        doc.add_heading(act['名稱'], level=1)
        p = doc.add_paragraph()
        p.add_run(f"📍 平台：{act['平台']} | 📅 期間：{act['開始']} - {act['結束']}").bold = True
        doc.add_heading('📝 活動內容', level=2)
        doc.add_paragraph(act['內容'])
        doc.add_heading('📌 執行規範', level=2)
        for s in act['規範'].split('\n'):
            if s.strip(): doc.add_paragraph(s.strip(), style='List Bullet')
        doc.add_page_break()
    
    word_io = BytesIO()
    doc.save(word_io)
    with c_word:
        st.download_button("📄 匯出活動企劃書 (.docx)", data=word_io.getvalue(), file_name="馬尼公告.docx")