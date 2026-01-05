import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from io import BytesIO
import os

# --- 1. 頁面配置與顏色調整 ---
st.set_page_config(page_title="馬尼通訊 企劃排程系統 v12.5", page_icon="🐎", layout="wide")

# 修改標題顏色為深藍色，其餘保持易讀性
st.markdown("""
    <style>
    .main { background-color: #F0F2F6; color: #1E2D4A; } /* 改為淺灰底深藍字，提升商務質感 */
    .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 { color: #0B1C3F !important; } /* 標題改為深藍色 */
    .stButton>button { background-color: #0B1C3F; color: white; border-radius: 8px; font-weight: bold; }
    .stDownloadButton>button { background-color: #27AE60; color: white; border-radius: 8px; font-weight: bold; }
    /* 側邊欄範本區樣式 */
    section[data-testid="stSidebar"] { background-color: #0B1C3F; color: white; }
    section[data-testid="stSidebar"] .stMarkdown h2 { color: #FFD700 !important; } /* 側邊欄標題保持金黃色 */
    </style>
    """, unsafe_allow_html=True)

# --- 2. 範本數據 ---
TEMPLATES = {
    "🐎 馬年慶：百倍奉還": {
        "name": "2026 馬尼通訊「馬年慶：百倍奉還」",
        "purpose": "迎接馬年，透過 $100 低門檻吸引新舊客，增加會員登錄與官網流量。",
        "core": "對象：全體消費者；範圍：全台門市；產品：「百倍奉還」禮包 ($100)。",
        "schedule": "01/12-01/18: 宣傳期 (FB/IG/脆前導)\n01/19-02/08: 販售期 (門市現場銷售)\n02/11: 開獎日 (官網公布)\n02/12-02/28: 兌獎期 (中獎核對)",
        "prizes": "Sony PS5 | 1 名 | 吸睛大獎\n現金 $6,666 | 1 名 | 百倍奉還獎\n官網購物金 $1,500 | 115 名 | 二次轉化關鍵",
        "sop": "1.限購3包。 2.告知序號重要性。 3.引導加入LINE。",
        "marketing": "倒數計時限動；弱勢分店區域廣告投遞。",
        "risk": "稅務申報流程；序號防偽蓋章；滯銷調度機制。",
        "effect": "預估帶動 2,000+ 進店人次；強化品牌高 CP 值形象。"
    },
    "📱 範本：新機上市": {"name": "新品發表企劃", "purpose": "", "core": "", "schedule": "", "prizes": "", "sop": "", "marketing": "", "risk": "", "effect": ""},
    "🎁 範本：品牌週年": {"name": "十週年盛典", "purpose": "", "core": "", "schedule": "", "prizes": "", "sop": "", "marketing": "", "risk": "", "effect": ""},
    "🛍️ 範本：門市振興": {"name": "弱勢門市支援方案", "purpose": "", "core": "", "schedule": "", "prizes": "", "sop": "", "marketing": "", "risk": "", "effect": ""}
}

# --- 3. 側邊欄佈局 ---
with st.sidebar:
    st.header("📋 快速範本")
    for t_name, t_data in TEMPLATES.items():
        if st.button(t_name):
            for key in t_data: st.session_state[f"p_{key}"] = t_data[key]
            st.rerun()

    st.divider()
    if st.button("🗑️ 清空所有草稿"):
        for key in list(st.session_state.keys()):
            if key.startswith("p_"): st.session_state[key] = ""
        st.rerun()

    # 系統資訊移至側邊欄底部，預設閉合
    with st.expander("🛠️ 系統資訊", expanded=False):
        st.caption("""
        **版本**: v12.5 (Professional)  
        **更新**: 
        - 輸出文件字體統一為微軟正黑體
        - 時程表自動生成 Word 時間軸表格
        - UI 顏色切換為商務深藍
        
        馬尼行銷規劃提案 © 2025 Money MKT
        """)

# --- 4. 主要編輯區 ---
st.title("📱 馬尼通訊 行銷企劃提案系統")

col_info1, col_info2, col_info3 = st.columns([2, 1, 1])
with col_info1: p_name = st.text_input("一、 活動名稱", key="p_name", placeholder="請輸入完整活動標題")
with col_info2: proposer = st.text_input("提案人", key="p_proposer", value="行銷部")
with col_info3: p_date = st.date_input("提案日期", value=datetime.now(), key="p_date")

st.divider()
c1, c2 = st.columns(2)
with c1:
    p_purpose = st.text_area("活動時機與目的", key="p_purpose", height=100)
    p_core = st.text_area("二、 活動核心內容", key="p_core", height=100)
    st.caption("時程建議格式：MM/DD-MM/DD: 內容描述")
    p_schedule = st.text_area("三、 活動時程安排", key="p_schedule", height=120)
    st.caption("贈品格式：品項 | 數量 | 備註")
    p_prizes = st.text_area("四、 贈品結構與預算", key="p_prizes", height=120)

with c2:
    p_sop = st.text_area("五、 門市執行流程 (SOP)", key="p_sop", height=100)
    p_marketing = st.text_area("六、 行銷流程與策略", key="p_marketing", height=100)
    p_risk = st.text_area("七、 風險管理與注意事項", key="p_risk", height=100)
    p_effect = st.text_area("八、 預估成效", key="p_effect", height=100)

# --- 5. Word 輸出美化 (正黑體 & 時間軸表格) ---
def set_font_msjh(run):
    """設置字體為微軟正黑體"""
    run.font.name = 'Microsoft JhengHei'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft JhengHei')

def generate_pro_word():
    doc = Document()
    
    # 字體預設設定
    style = doc.styles['Normal']
    set_font_msjh(style.node)

    # A. 代入 Logo
    if os.path.exists("logo.png"):
        doc.add_picture("logo.png", width=Inches(1.2))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # B. 標題區
    h = doc.add_heading('行銷企劃執行提案書', 0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    info = doc.add_paragraph()
    info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_info = info.add_run(f"提案人：{st.session_state.get('p_proposer')}  |  日期：{st.session_state.get('p_date')}")
    set_font_msjh(run_info)

    doc.add_heading(st.session_state.get('p_name', '未命名企劃'), level=1)

    # C. 章節邏輯
    sections = [
        ("一、 活動時機與目的", st.session_state.p_purpose),
        ("二、 活動核心內容", st.session_state.p_core),
        ("三、 活動時程安排 (Timeline)", st.session_state.p_schedule),
        ("四、 贈品結構與預算", st.session_state.p_prizes),
        ("五、 門市執行流程 (SOP)", st.session_state.p_sop),
        ("六、 行銷流程與策略", st.session_state.p_marketing),
        ("七、 風險管理與注意事項", st.session_state.p_risk),
        ("八、 預估成效", st.session_state.p_effect)
    ]

    for title_text, content in sections:
        h2 = doc.add_heading(title_text, level=2)
        h2.runs[0].font.color.rgb = RGBColor(11, 28, 63) # 深藍色章節
        
        # 1. 時間軸表格化 (針對第三點)
        if "時程安排" in title_text:
            t = doc.add_table(rows=1, cols=2)
            t.style = 'Light Shading Accent 1'
            t.rows[0].cells[0].text = "階段/日期"
            t.rows[0].cells[1].text = "執行細節"
            for line in content.split('\n'):
                if ":" in line or "-" in line:
                    parts = line.split(':') if ":" in line else line.split(' ')
                    row = t.add_row().cells
                    row[0].text = parts[0].strip()
                    row[1].text = parts[1].strip() if len(parts)>1 else ""
        
        # 2. 贈品表格化 (針對第四點)
        elif "贈品結構" in title_text and "|" in content:
            t = doc.add_table(rows=1, cols=3)
            t.style = 'Table Grid'
            hdr = t.rows[0].cells
            hdr[0].text, hdr[1].text, hdr[2].text = "品項", "數量", "備註"
            for line in content.split('\n'):
                if "|" in line:
                    parts = line.split('|')
                    row = t.add_row().cells
                    for i in range(min(len(parts), 3)): row[i].text = parts[i].strip()
        
        # 3. 一般文字
        else:
            p = doc.add_paragraph(content)
            set_font_msjh(p.add_run(""))

    word_io = BytesIO()
    doc.save(word_io)
    return word_io.getvalue()

# --- 6. 輸出按鈕 ---
st.divider()
if p_name:
    if st.button("✅ 完成企劃並產生文檔"):
        doc_bytes = generate_pro_word()
        st.download_button(
            label="📥 下載馬尼行銷企劃書 (微軟正黑體版)",
            data=doc_bytes,
            file_name=f"MoneyMKT_{p_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
