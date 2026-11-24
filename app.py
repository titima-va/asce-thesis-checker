import io
import re
from pathlib import Path
import docx
import fitz
import pandas as pd
import streamlit as st

st.set_page_config(page_title="ASCE Reference Checker | ระบบตรวจสอบเอกสารอ้างอิง", page_icon="🎓", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Thai:wght@400;600;700&family=Inter:wght@400;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'IBM Plex Sans Thai', 'Inter', sans-serif; }
    .stApp, [data-testid="stAppViewContainer"] { background-color: #d9ecf0 !important; }
    [data-testid="stHeader"] { background-color: rgba(217, 236, 240, 0.85) !important; }
    .hero-banner {
        background: linear-gradient(135deg, #b52324 0%, #283848 54%, #33c0ca 100%);
        border-bottom: 5px solid #33c0ca; padding: 2rem 2.2rem; border-radius: 14px;
        color: #ffffff; margin-bottom: 1.5rem; box-shadow: 0 4px 16px rgba(181, 35, 36, 0.16);
    }
    .hero-badge {
        display: inline-block; background-color: rgba(255, 255, 255, 0.16);
        border: 1px solid rgba(51, 192, 202, 0.85); color: #ffffff;
        padding: 0.28rem 0.85rem; border-radius: 20px; font-size: 0.85rem; margin-bottom: 0.65rem;
    }
    .hero-title { font-size: 1.85rem; font-weight: 700; margin: 0 0 0.5rem 0; color: #ffffff; }
    .hero-subtitle { font-size: 1rem; color: #f1f5f9; margin: 0; line-height: 1.6; }
    div[data-testid="stExpander"] {
        background-color: #ffffff !important; border: 2px solid #33c0ca !important;
        border-radius: 12px !important; overflow: hidden; margin-bottom: 1.2rem;
    }
    div[data-testid="stExpander"] > details > summary {
        background-color: #33c0ca !important; color: #0f2942 !important;
        font-weight: 700 !important; padding: 0.75rem 1rem !important;
    }
    div[data-testid="stExpander"] > details > summary:hover { background-color: #29aeb8 !important; }
    div[data-testid="stExpander"] > details > summary p,
    div[data-testid="stExpander"] > details > summary span {
        font-weight: 700 !important; font-size: 1.02rem !important; color: #0f2942 !important;
    }
    div[data-testid="stExpander"] > details > summary svg { fill: #0f2942 !important; color: #0f2942 !important; }
    section[data-testid="stSidebar"] { border-right: 3px solid #33c0ca; }
    div[data-testid="stMetric"] {
        background-color: #ffffff; border: 1px solid #cbd5e1;
        border-left: 5px solid #33c0ca; padding: 1rem 1.2rem; border-radius: 10px;
    }
    div[data-testid="column"]:nth-child(3) div[data-testid="stMetric"],
    div[data-testid="column"]:nth-child(4) div[data-testid="stMetric"] { border-left: 5px solid #b52324; }
    div[data-testid="stFileUploader"] label[data-testid="stWidgetLabel"] p {
        font-weight: 700 !important; font-size: 1.06rem !important; color: #0f2942 !important;
    }
    section[data-testid="stFileUploader"] {
        background-color: #ffffff; border: 2px dashed #33c0ca; border-radius: 12px; padding: 1rem;
    }
    section[data-testid="stFileUploader"]:hover { border-color: #b52324; }
    div[data-testid="stDownloadButton"] > button {
        background-color: #b52324 !important; color: #ffffff !important;
        border: none !important; border-radius: 8px; font-weight: 600; padding: 0.55rem 1.2rem;
    }
    div[data-testid="stDownloadButton"] > button:hover { background-color: #33c0ca !important; color: #0f2942 !important; }
    button[data-baseweb="tab"][aria-selected="true"] {
        color: #b52324 !important; border-bottom-color: #b52324 !important; font-weight: 600;
    }
    div[data-testid="stAlert"] {
        background-color: #ffffff !important; border: 1.5px solid #33c0ca !important; border-radius: 10px !important;
    }
    .bottom-divider { margin-top: 3rem; margin-bottom: 1.2rem; border: none; border-top: 2px solid #33c0ca; }
</style>
""", unsafe_allow_html=True)

TEXTS = {
    "th": {
        "org": "บัณฑิตศึกษา วิทยาลัยการคอมพิวเตอร์ | ระบบบริการตรวจสอบรูปแบบวิทยานิพนธ์",
        "title": "ระบบตรวจสอบเอกสารอ้างอิงวิทยานิพนธ์ (มาตรฐาน ASCE)",
        "subtitle": "ตรวจสอบความสอดคล้องระหว่างการอ้างอิงในเนื้อหา (In-text Citations) กับบรรณานุกรมท้ายเล่ม (References / Bibliography) พร้อมตรวจรูปแบบการพิมพ์ตามเกณฑ์ American Society of Civil Engineers อัตโนมัติ",
        "sleep_notice": '💡 **หมายเหตุการใช้งาน:** หากเปิดลิงก์แล้วพบหน้าจอแจ้งเตือนพร้อมปุ่มสีฟ้า **"Yes, get this app back up!"** ให้กดปุ่มดังกล่าวและรอประมาณ 15 วินาที ระบบจะเปิดใช้งานได้ตามปกติ',
        "guide_title": "📌 **คำแนะนำการเตรียมไฟล์และขอบเขตการตรวจสอบ (คลิกเพื่ออ่าน)**",
        "guide_body": """**1. การเตรียมไฟล์ก่อนอัปโหลด:**
* รองรับไฟล์นามสกุล **`.docx` (แนะนำเพื่อความแม่นยำสูงสุด)** และ **`.pdf`** ขนาดไม่เกิน 200 MB
* ในเอกสารต้องมีหัวข้อ **`References`**, **`Bibliography`**, **`เอกสารอ้างอิง`** หรือ **`บรรณานุกรม`** แยกขึ้นบรรทัดใหม่อย่างชัดเจนก่อนเริ่มรายการอ้างอิงท้ายเล่ม

**2. สิ่งที่ระบบตรวจสอบให้:**
* **Cross-Matching:** เช็กว่าทุกจุดที่อ้างอิงในเนื้อหามีรายชื่อท้ายเล่มครบหรือไม่ มีรายการเกินท้ายเล่มหรือไม่ และปี พ.ศ./ค.ศ. ตรงกันหรือไม่
* **ASCE Formatting:** ตรวจสอบการใช้คำว่า `and` (ห้ามใช้ `&`), การไม่ใส่เครื่องหมายจุลภาค `,` คั่นระหว่างชื่อผู้แต่งกับปีในเนื้อหา, การใส่จุดหลัง `et al.`, การเรียงลำดับ A–Z และเครื่องหมายคำพูดครอบชื่อบทความวารสาร""",
        "upload_label": "📂 **อัปโหลดไฟล์ร่างวิทยานิพนธ์ (.docx หรือ .pdf)**",
        "spinner": "กำลังวิเคราะห์โครงสร้างเอกสารและตรวจสอบรายการอ้างอิง...",
        "success": "ตรวจสอบไฟล์ **{filename}** เสร็จสมบูรณ์!",
        "m_intext": "การอ้างอิงในเนื้อหา (จุด)",
        "m_refs": "บรรณานุกรมท้ายเล่ม (รายการ)",
        "m_cross": "จุดไม่สอดคล้องกัน (รายการ)",
        "m_format": "จุดผิดรูปแบบ ASCE (รายการ)",
        "dl_btn": "📥 ดาวน์โหลดรายงานผลการตรวจสอบฉบับเต็ม (Excel)",
        "tab1": "❌ ความคลาดเคลื่อนเนื้อหา–ท้ายเล่ม",
        "tab2": "⚠️ จุดผิดรูปแบบมาตรฐาน ASCE",
        "tab3": "📋 รายชื่อบรรณานุกรมที่ตรวจพบ",
        "no_cross": "ยอดเยี่ยม! ไม่พบความคลาดเคลื่อน การอ้างอิงในเนื้อหาและท้ายเล่มสอดคล้องกันครบถ้วน",
        "no_format": "ยอดเยี่ยม! ไม่พบข้อผิดพลาดด้านรูปแบบตามเกณฑ์ ASCE เบื้องต้น",
        "no_refs": "ไม่พบรายการบรรณานุกรมในเอกสาร",
        "err_prefix": "ไม่สามารถประมวลผลเอกสารได้:",
        "err_ext": "รองรับเฉพาะไฟล์ .docx และ .pdf เท่านั้น",
        "err_no_heading": "ไม่พบหัวข้อ 'References', 'Bibliography' หรือ 'เอกสารอ้างอิง' ในไฟล์ กรุณาตรวจสอบว่ามีหัวข้อดังกล่าวขึ้นบรรทัดใหม่ในเอกสาร",
        "sb_header": "📖 ตัวอย่างรูปแบบ ASCE ที่ถูกต้อง",
        "sb_intext_title": "1. การอ้างอิงในเนื้อหา (In-text)",
        "sb_intext_ex": """* **ผู้แต่ง 1 คน:** `(Smith 2020)` หรือ `Smith (2020)`
* **ผู้แต่ง 2 คน:** `(Smith and Jones 2020)` *(ใช้ and ห้ามใช้ &)*
* **ผู้แต่ง 3 คนขึ้นไป:**
