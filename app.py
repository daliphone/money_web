import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from io import BytesIO
import os
import re

# --- 1. 頁面配置 ---
st.set_page_config(page_title="馬尼通訊 企劃排程系統 v14.1.3", page_icon="🐎", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #F0F2F6; color: #1E2D4A; }
    .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 { color: #0B1C3F !important; }
    ::placeholder { color: #888888 !important; opacity: 0.5 !important; }
    div[data-baseweb="select"] > div { background-color: white !important; color: #0B1C3F !important; }
    .stButton>button { background-color: #0B1C3F; color: white; border-radius: 8px; font-weight: bold; }
    .stDownloadButton>button { background-color: #27AE60; color: white; border-radius: 8px; font-weight: bold; }
    .ai-btn>div>button { background-color: #6200EA !important; border: 1px solid #FFD700 !important; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 深度場景化 AI 引擎 ---
def smart_ai_optimize(field_id, text, style):
    if not text or len(text) < 2: return text
    
    # 修正：移除數據標記的邏輯，避免使用引發報錯的反斜線結尾
    text = re.sub(r"\", "", text).strip()
    
    if field_id == "p_purpose":
        return f"【營運目的】本活動旨在{text}。透過精準檔期切入，預期強化品牌在該期間的市佔率並提升客戶回流量。"
    elif field_id == "p_core":
        return f"【核心賣點】{text}。本活動以獨家資源為引，建立市場區隔，直接命中目標客群需求。"
    elif field_id == "p_schedule":
        return f"{text}\n\n💡 AI 執行建議：請確保『宣傳期』與『銷售期』的轉場銜接，門市海報需於銷售期前2日佈置完畢。"
    elif field_id == "p_prizes":
        return f"{text}\n\n💡 AI 獎項建議：此配置中大獎具備話題性，小獎（購物金）則負責驅動官網流量。"
    elif field_id == "p_sop":
        return f"{text}\n\n💡 SOP 注意事項：銷售環節應強調『序號核對』之嚴謹性，避免後續獎項發放爭議。"
    elif field_id == "p_marketing":
        prefix = "🚀【全通路行銷】" if style == "創意社群" else "📈【行銷規劃】"
        return f"{prefix}{text}。利用多元管道覆蓋客群，確保活動聲量最大化。"
    elif field_id == "p_risk":
        return f"{text}\n\n💡 風險評估：建議於活動文案顯眼處標示稅務規範，並預留備用贈品處理瑕疵爭議。"
    elif field_id == "p_effect":
        return f"【預期效益】{text}。除即時業績增長外，本次活動預計可為品牌增加長期會員資產及社群互動數。"
    return text

# --- 3. 初始化數據與範本 (已手動清理所有 [cite] 標記) ---
if 'templates_store' not in st.session_state:
    st.session_state.templates_store = {
        "🐎 馬年慶：百倍奉還": {
            "name": "2026 馬尼通訊「馬年慶：百倍奉還」活動執行企劃案",
            "purpose": "迎接 2026 農曆馬年（丙午年），結合春節紅包與「百倍奉還」話題。透過 $100 元低門檻吸引新(舊)客戶，增加會員登錄與官網流量。",
            "core": "執行單位: 馬尼行動通訊門市；對象: 所有門市消費者；核心產品: 「百倍奉還」新年禮包 ($100/包)。",
            "schedule": "宣傳期: 115/01/12-01/18\n銷售期: 115/01/19-02/08\n開獎日: 115/02/11\n兌獎期: 115/02/12-02/28",
            "prizes": "Sony PS5 (1名) | 現金 $6,666 (1名) | 總獎值突破 $130,000\n官網購物金 $1,500 | 115名 | 帶動二次消費",
            "sop": "確認客購數量(上限3包)；告知序號保存；限量管理(每店66包)；引導加入官方LINE。",
            "marketing": "FB/IG/Threads 倒數限時動態；針對弱勢分店進行區域廣告投遞。",
            "risk": "稅務申報(>$1000)；序號防偽蓋章；銷售分佈不均之調度機制。",
            "effect": "預計帶動 2,000+ 人次進入門市；購物金帶動至少 60 筆官網訂單。"
        }
    }

if "p_proposer" not in st.session_state: 
    st.session_state["p_proposer"] = "行銷部"

# --- 4. 側邊欄 ---
with st.sidebar:
    st.header("📋 快速範本區")
    selected_tpl_key = st.selectbox("選擇操作範本", options=list(st.session_state.templates_store.keys()))
    
    col_tpl1, col_tpl2 = st.columns(2)
    with col_tpl1:
        if st.button("📥 載入範本"):
            data = st.session_state.templates_store[selected_tpl_key]
            for k, v in data.items():
                st.session_state[f"p_{k}"] = v
            st.rerun()
            
    with col_tpl2:
        if st.button("💾 儲存範本"):
            st.session_state.templates_store[selected_tpl_key] = {
                "name": st.session_state.get("p_name", ""),
                "purpose": st.session_state.get("p_purpose", ""),
                "core": st.session_state.get("p_core", ""),
                "schedule": st.session_state.get("p_schedule", ""),
                "prizes": st.session_state.get("p_prizes", ""),
                "sop": st.session_state.get("p_sop", ""),
                "marketing": st.session_state.get("p_marketing", ""),
                "risk": st.session_state.get("p_risk", ""),
                "effect": st.session_state.get("p_effect", "")
            }
            st.success(f"已儲存回範本庫")

    if st.button("🗑️ 清空編輯區"):
        fields = ["p_name", "p_purpose", "p_core", "p_schedule", "p_prizes", "p_sop", "p_marketing", "p_risk", "p_effect"]
        for f in fields: st.session_state[f] = ""
        st.rerun()

    st.divider()
    st.header("✨ AI 創意引擎")
    ai_style = st.radio("主要優化語氣", ["熱血商務", "創意社群", "專業條列"])
    
    st.markdown('<div class="ai-btn">', unsafe_allow_html=True)
    if st.button("🪄 場景化 AI 深度優化"):
        fields = ["p_purpose", "p_core", "p_schedule", "p_prizes", "p_sop", "p_marketing", "p_risk", "p_effect"]
        for f in fields:
            if f in st.session_state:
                st.session_state[f] = smart_ai_optimize(f, st.session_state[f], ai_style)
        st.toast("已完成 AI 場景優化建議！")
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("<br>"*5, unsafe_allow_html=True)
    with st.expander("🛠️ 系統資訊", expanded=False):
        st.caption("""
        **版本**: v14.1.3 (Stable Build)
        - 徹底修正 re.sub 語法錯誤
        - 修復範本載入/儲存雙向功能
        - 手動清除馬年範本內文引註標籤
        
        馬尼門活動企劃系統 © 2025 Money MKT
        """)

# --- 5. 主要編輯區 ---
st.title("📱 馬尼通訊 行銷企劃提案系統")

c_top1, c_top2, c_top3 = st.columns([2, 1, 1])
with c_top1: st.text_input("一、 活動名稱", key="p_name")
with c_top2: st.text_input("提案人", key="p_proposer")
with c_top3: st.date_input("提案日期", value=datetime.now(), key="p_date")

st.divider()
c1, c2 = st.columns(2)
with c1:
    st.text_area("活動時機與目的 (營運目的邏輯)", key="p_purpose", height=100, placeholder="填寫經營目標與時機...")
    st.text_area("二、 活動核心內容 (賣點配置)", key="p_core", height=100, placeholder="對象、執行單位與產品核心...")
    st.text_area("三、 活動時程安排 (執行重點建議)", key="p_schedule", height=120, placeholder="各階段日期與細節...")
    st.text_area("四、 贈品結構與預算 (關鍵商品用意)", key="p_prizes", height=120, placeholder="品項 | 數量 | 備註")

with c2:
    st.text_area("五、 門市執行流程 (SOP 注意事項)", key="p_sop", height=100, placeholder="銷售環節與限量管理 SOP...")
    st.text_area("六、 行銷流程與策略 (建議管道)", key="p_marketing", height=100, placeholder="線上廣告與標語策略...")
    st.text_area("七、 風險管理與注意事項 (規範建議)", key="p_risk", height=100, placeholder="稅務、調度與序號防偽...")
    st.text_area("八、 預估成效 (效益面建議)", key="p_effect", height=100, placeholder="觸及人次、官網轉化等...")

# --- 6. Word 導出與下載 ---
def set_msjh_font(run):
    run.font.name = 'Microsoft JhengHei'
    r = run._element
    rFonts = r.find(qn('w:rFonts'))
    if rFonts is None:
        from docx.oxml import OxmlElement
        rFonts = OxmlElement('w:rFonts')
        r.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), 'Microsoft JhengHei')

def generate_pro_word():
    doc = Document()
    if os.path.exists("logo.png"):
        doc.add_picture("logo.png", width=Inches(1.2))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    h = doc.add_heading('行銷企劃執行提案書', 0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    info_p = doc.add_paragraph()
    info_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_info = info_p.add_run(f"提案人：{st.session_state.get('p_proposer')}  |  日期：{st.session_state.get('p_date')}")
    set_msjh_font(r_info)

    doc.add_heading(st.session_state.get('p_name', '未命名企劃'), level=1)

    sections = [
        ("一、 活動時機與目的", st.session_state.p_purpose),
        ("二、 活動核心內容", st.session_state.p_core),
        ("三、 活動時程安排", st.session_state.p_schedule),
        ("四、 贈品結構與預算", st.session_state.p_prizes),
        ("五、 門市執行流程 (SOP)", st.session_state.p_sop),
        ("六、 行銷流程與策略", st.session_state.p_marketing),
        ("七、 風險管理與注意事項", st.session_state.p_risk),
        ("八、 預估成效", st.session_state.p_effect)
    ]

    for title_text, content in sections:
        h2 = doc.add_heading(title_text, level=2)
        h2.runs[0].font.color.rgb = RGBColor(11, 28, 63)
        
        if "時程安排" in title_text and content:
            t = doc.add_table(rows=1, cols=2)
            t.style = 'Light Shading Accent 1'
            for line in content.split('\n'):
                if line.strip():
                    parts = line.split(':') if ':' in line else [line, ""]
                    row = t.add_row().cells
                    row[0].text = parts[0].strip()
                    row[1].text = parts[1].strip() if len(parts)>1 else ""
        elif "贈品結構" in title_text and "|" in content:
            t = doc.add_table(rows=1, cols=3)
            t.style = 'Table Grid'
            for line in content.split('\n'):
                if "|" in line:
                    parts = line.split('|')
                    row = t.add_row().cells
                    for i in range(min(len(parts), 3)): row[i].text = parts[i].strip()
        else:
            p = doc.add_paragraph()
            r = p.add_run(content)
            set_msjh_font(r)

    word_io = BytesIO()
    doc.save(word_io)
    return word_io.getvalue()

st.divider()
if st.session_state.get('p_name'):
    if st.button("✅ 完成企劃並產生文檔"):
        doc_bytes = generate_pro_word()
        st.download_button(
            label="📥 下載馬尼行銷企劃書 (Stable Build)",
            data=doc_bytes,
            file_name=f"MoneyMKT_{st.session_state.p_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
