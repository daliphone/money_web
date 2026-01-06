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
st.set_page_config(page_title="馬尼通訊 企劃排程系統 v14.2.0", page_icon="🐎", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #F0F2F6; color: #1E2D4A; }
    .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 { color: #0B1C3F !important; }
    .stButton>button { background-color: #0B1C3F; color: white; border-radius: 8px; font-weight: bold; }
    .stDownloadButton>button { background-color: #27AE60; color: white; border-radius: 8px; font-weight: bold; }
    .ai-btn>div>button { background-color: #6200EA !important; border: 1px solid #FFD700 !important; }
    /* 調整提示標籤顏色 */
    .stTooltipIcon { color: #0B1C3F !important; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 深度場景化 AI 引擎 ---
def smart_ai_optimize(field_id, text, style):
    if not text or len(text) < 2: return text
    text = text.replace("", "").strip()
    
    if field_id == "p_purpose":
        return f"【營運目的】本活動旨在{text}。透過精準檔期切入，預期強化品牌在該期間的市佔率並提升客戶回流量。"
    elif field_id == "p_core":
        return f"【核心賣點】{text}。本活動以獨家資源為引，建立市場區隔，直接命中目標客群需求。"
    elif field_id == "p_schedule":
        return f"{text}\n\n💡 AI 執行建議：請確保『宣傳期』與『銷售期』的轉場銜接，門市海報需於銷售期前2日佈置完畢。"
    elif field_id == "p_prizes":
        return f"{text}\n\n💡 AI 獎項建議：此配置中大獎具備話題性，小獎則負責驅動官網流量。"
    elif field_id == "p_sop":
        return f"{text}\n\n💡 SOP 注意事項：應包含「卸下武裝」話術，先詢問需求而非直接推產品，提升客戶信任感。"
    elif field_id == "p_marketing":
        prefix = "🚀【全通路行銷】" if style == "創意社群" else "📈【行銷規劃】"
        return f"{prefix}{text}。利用多元管道覆蓋客群，確保活動聲量最大化。"
    elif field_id == "p_risk":
        return f"{text}\n\n💡 風險評估：需明確定義「損壞界定」與「退場機制」，標示稅務規範以避免爭議。"
    elif field_id == "p_effect":
        return f"【預期效益】{text}。除業績外，應蒐集真實使用回饋(UGC)，優化未來銷售策略。"
    return text

# --- 3. 初始化數據與範本 ---
if 'templates_store' not in st.session_state:
    st.session_state.templates_store = {
        "🐎 馬年慶：百倍奉還": {
            "name": "2026 馬尼通訊「馬年慶：百倍奉還」活動執行企劃案",
            "purpose": "迎接 2026 農曆馬年，結合春節紅包與「百倍奉還」話題。吸引新舊客戶，增加會員登錄與官網流量。",
            "core": "執行單位: 馬尼行動通訊門市；對象: 所有門市消費者；核心產品: 「百倍奉還」新年禮包 ($100/包)。",
            "schedule": "宣傳期: 115/01/12-01/18\\n銷售期: 115/01/19-02/08\\n開獎日: 115/02/11\\n兌獎期: 115/02/12-02/28",
            "prizes": "Sony PS5 (1名) | 現金 $6,666 (1名) | 總獎值突破 $130,000\\n官網購物金 $1,500 | 115名 | 帶動二次消費",
            "sop": "確認客購數量；告知序號保存；引導加入官方LINE。話術建議：先聊過年需求，再帶出禮包價值。",
            "marketing": "FB/IG/Threads 倒數計時限時動態；針對弱勢分店進行區域廣告投遞。",
            "risk": "稅務申報(>$1000)；序號防偽蓋章；明確定義中獎者領取期限與流程。",
            "effect": "預計帶動 2,000+ 人次進入門市；購物金帶動至少 60 筆官網訂單。"
        }
    }

# --- 4. 側邊欄 ---
with st.sidebar:
    st.header("📋 快速範本區")
    selected_tpl_key = st.selectbox("選擇操作範本", options=list(st.session_state.templates_store.keys()))
    
    col_tpl1, col_tpl2 = st.columns(2)
    with col_tpl1:
        if st.button("📥 載入範本"):
            data = st.session_state.templates_store[selected_tpl_key]
            for k, v in data.items(): st.session_state[f"p_{k}"] = v
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
            st.success("已儲存回範本庫")

    st.divider()
    st.header("✨ AI 顧問引擎")
    ai_style = st.radio("主要優化語氣", ["熱血商務", "創意社群", "專業條列"])
    
    st.markdown('<div class="ai-btn">', unsafe_allow_html=True)
    if st.button("🪄 執行場景化 AI 優化"):
        fields = ["p_purpose", "p_core", "p_schedule", "p_prizes", "p_sop", "p_marketing", "p_risk", "p_effect"]
        for f in fields:
            if f in st.session_state:
                st.session_state[f] = smart_ai_optimize(f, st.session_state[f], ai_style)
        st.toast("已完成 AI 顧問優化！")
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("<br>"*5, unsafe_allow_html=True)
    with st.expander("🛠️ 系統資訊 v14.2.0", expanded=False):
        st.caption("新增功能：\n1. 欄位提示視窗(Tooltip)\n2. 整合「試戴專案」邏輯建議\n3. 強化 SOP 心理戰話術建議")

# --- 5. 主要編輯區 ---
st.title("📱 馬尼通訊 行銷企劃提案系統")

c_top1, c_top2, c_top3 = st.columns([2, 1, 1])
with c_top1: st.text_input("一、 活動名稱", key="p_name")
with c_top2: st.text_input("提案人", key="p_proposer", value=st.session_state.get("p_proposer", "行銷部"))
with c_top3: st.date_input("提案日期", value=datetime.now(), key="p_date")

st.divider()
c1, c2 = st.columns(2)

with c1:
    st.text_area("活動時機與目的", key="p_purpose", height=100, 
                 help="【建議順序 1】核心價值：定義活動是為了解決什麼痛點？量化目標：除了銷售額，是否包含蒐集真實數據或社群素材(UGC)？")
    
    st.text_area("二、 活動核心內容 (賣點配置)", key="p_core", height=100,
                 help="【建議順序 2】活動機制設計：分階段說明申請/開始、體驗期間及結束後選擇。透明化表格：列出租借成本、售價、活動價及押金。")
    
    st.text_area("三、 活動時程安排", key="p_schedule", height=120,
                 help="【建議順序 3】包含宣傳期、執行期、結案期。確保第一線人員在每個時間點都知道要做什麼。")
    
    st.text_area("四、 贈品結構與預算", key="p_prizes", height=120,
                 help="【建議順序 4】誘因機制：任務化獎勵（如完成分享即贈小禮）。區分購買與否：即使未成交，只要有回饋也給予小贈品建立長期信任。")

with c2:
    st.text_area("五、 門市執行 SOP (含實戰話術)", key="p_sop", height=100,
                 help="【建議順序 7】實戰話術：1. 卸下武裝：不要一開始推產品。2. 反向推銷：建議客人「先體驗不要直接買」。3. 禁語列表：避開「今天不買會沒了」。")
    
    st.text_area("六、 行銷宣傳與策略", key="p_marketing", height=100,
                 help="【建議順序 4】擴散機制：社群任務設計、FB/IG/Threads 倒數限時動態，增加緊張感與話題。")
    
    st.text_area("七、 風險管理與退場機制", key="p_risk", height=100,
                 help="【建議順序 6】控管機制：明確定義損壞界定、押金退還條件、稅務法規申報、及銷售不均的內部調度。")
    
    st.text_area("八、 預估成效與數據蒐集", key="p_effect", height=100,
                 help="【建議順序 5】數據蒐集：問卷設計，詢問「影響購買的主要原因」與「體驗是否幫助決策」，作為優化話術的指標。")

# --- 6. Word 導出與下載 (保持穩定邏輯) ---
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
            t = doc.add_table(rows=1, cols=2); t.style = 'Light Shading Accent 1'
            for line in content.split('\\n'):
                if line.strip():
                    parts = line.split(':') if ':' in line else [line, ""]
                    row = t.add_row().cells
                    row[0].text = parts[0].strip(); row[1].text = parts[1].strip() if len(parts)>1 else ""
        elif "贈品結構" in title_text and "|" in content:
            t = doc.add_table(rows=1, cols=3); t.style = 'Table Grid'
            for line in content.split('\\n'):
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
            label=f"📥 下載 {st.session_state.p_name} 企劃書",
            data=doc_bytes,
            file_name=f"MoneyMKT_{st.session_state.p_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
