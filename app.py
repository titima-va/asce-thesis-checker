import io
import re
from pathlib import Path
import docx
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st

# =========================================================
# 1. ตั้งค่าชื่อหน่วยงาน (สามารถแก้ไขข้อความในเครื่องหมายคำพูดได้ตามต้องการ)
# =========================================================
ORG_NAME_TH = "บัณฑิตศึกษา วิทยาลัยการคอมพิวเตอร์ | ระบบบริการตรวจสอบรูปแบบวิทยานิพนธ์"
ORG_NAME_EN = "Graduate Studies, College of Computing | Thesis Formatting & Reference Service"
CONTACT_TH = "หากพบข้อสงสัยเกี่ยวกับรูปแบบการอ้างอิง กรุณาติดต่อเจ้าหน้าที่งานบัณฑิตศึกษา"
CONTACT_EN = "For formatting inquiries, please contact the Graduate Studies Office."

# =========================================================
# 2. ตั้งค่าหน้าเว็บไซต์และ CSS ธีมสีประจำวิทยาลัย (#33c0ca และ #b52324)
# =========================================================
st.set_page_config(
    page_title="ASCE Reference Checker | ระบบตรวจสอบเอกสารอ้างอิง",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Thai:wght@300;400;500;600;700&family=Inter:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'IBM Plex Sans Thai', 'Inter', sans-serif;
    }

    /* แบนเนอร์ส่วนหัว (ไล่เฉดสี #b52324 และ #33c0ca) */
    .hero-banner {
        background: linear-gradient(135deg, #b52324 0%, #283848 54%, #33c0ca 100%);
        border-bottom: 5px solid #33c0ca;
        padding: 2rem 2.2rem;
        border-radius: 14px;
        color: #ffffff;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 16px rgba(181, 35, 36, 0.16);
    }
    .hero-badge {
        display: inline-block;
        background-color: rgba(255, 255, 255, 0.16);
        border: 1px solid rgba(51, 192, 202, 0.85);
        color: #ffffff;
        padding: 0.28rem 0.85rem;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 500;
        margin-bottom: 0.65rem;
    }
    .hero-title {
        font-size: 1.85rem;
        font-weight: 700;
        margin: 0 0 0.5rem 0;
        line-height: 1.3;
        color: #ffffff;
    }
    .hero-subtitle {
        font-size: 1rem;
        color: #f1f5f9;
        margin: 0;
        line-height: 1.6;
    }

    /* กล่องคำแนะนำการเตรียมไฟล์ (เน้นสี #33c0ca และตัวหนา) */
    div[data-testid="stExpander"] {
        border: 2px solid #33c0ca !important;
        border-radius: 12px !important;
        overflow: hidden;
        margin-bottom: 1.2rem;
        box-shadow: 0 2px 8px rgba(51, 192, 202, 0.15);
    }
    div[data-testid="stExpander"] > details > summary {
        background-color: #33c0ca !important;
        color: #0f2942 !important;
        font-weight: 700 !important;
        padding: 0.75rem 1rem !important;
        transition: background-color 0.2s ease;
    }
    div[data-testid="stExpander"] > details > summary:hover {
        background-color: #29aeb8 !important;
    }
    div[data-testid="stExpander"] > details > summary p,
    div[data-testid="stExpander"] > details > summary span {
        font-weight: 700 !important;
        font-size: 1.02rem !important;
        color: #0f2942 !important;
    }
    div[data-testid="stExpander"] > details > summary svg {
        fill: #0f2942 !important;
        color: #0f2942 !important;
    }

    /* แถบด้านข้าง (Sidebar) */
    section[data-testid="stSidebar"] {
        border-right: 3px solid #33c0ca;
    }

    /* กล่องการ์ดแสดงตัวเลขสถิติ */
    div[data-testid="stMetric"] {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-left: 5px solid #33c0ca;
        padding: 1rem 1.2rem;
        border-radius: 10px;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.04);
    }
    div[data-testid="column"]:nth-child(3) div[data-testid="stMetric"],
    div[data-testid="column"]:nth-child(4) div[data-testid="stMetric"] {
        border-left: 5px solid #b52324;
    }

    /* หัวข้อและกล่องอัปโหลดไฟล์ */
    div[data-testid="stFileUploader"] label[data-testid="stWidgetLabel"] p {
        font-weight: 700 !important;
        font-size: 1.06rem !important;
        color: #0f2942 !important;
    }
    section[data-testid="stFileUploader"] {
        background-color: #f3fbfc;
        border: 2px dashed #33c0ca;
        border-radius: 12px;
        padding: 1rem;
        transition: border-color 0.2s ease;
    }
    section[data-testid="stFileUploader"]:hover {
        border-color: #b52324;
    }

    /* ปุ่มดาวน์โหลดรายงาน Excel */
    div[data-testid="stDownloadButton"] > button {
        background-color: #b52324 !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 8px;
        font-weight: 600;
        padding: 0.55rem 1.2rem;
        transition: all 0.2s ease;
    }
    div[data-testid="stDownloadButton"] > button:hover {
        background-color: #33c0ca !important;
        color: #0f2942 !important;
    }

    /* แถบแท็บแสดงผล */
    button[data-baseweb="tab"][aria-selected="true"] {
        color: #b52324 !important;
        border-bottom-color: #b52324 !important;
        font-weight: 600;
    }

    /* ส่วนท้ายเว็บ */
    .footer-box {
        margin-top: 3rem;
        padding-top: 1rem;
        padding-bottom: 0.8rem;
        border-top: 2px solid #33c0ca;
        text-align: center;
        color: #475569;
        font-size: 0.88rem;
    }
</style>
""", unsafe_allow_html=True)

# =========================================================
# 3. คลังข้อความ 2 ภาษา (TH / EN)
# =========================================================
TEXTS = {
    "th": {
        "org": ORG_NAME_TH,
        "title": "ระบบตรวจสอบเอกสารอ้างอิงวิทยานิพนธ์ (มาตรฐาน ASCE)",
        "subtitle": "ตรวจสอบความสอดคล้องระหว่างการอ้างอิงในเนื้อหา (In-text Citations) กับบรรณานุกรมท้ายเล่ม (References) พร้อมตรวจรูปแบบการพิมพ์ตามเกณฑ์ American Society of Civil Engineers อัตโนมัติ",
        "sleep_notice": "💡 **หมายเหตุการใช้งาน:** หากเปิดลิงก์แล้วพบหน้าจอแจ้งเตือนพร้อมปุ่มสีฟ้า **\"Yes, get this app back up!\"** ให้กดปุ่มดังกล่าวและรอประมาณ 15 วินาที ระบบจะเปิดใช้งานได้ตามปกติ",
        "guide_title": "📌 **คำแนะนำการเตรียมไฟล์และขอบเขตการตรวจสอบ (คลิกเพื่ออ่าน)**",
        "guide_body": """
**1. การเตรียมไฟล์ก่อนอัปโหลด:**
* รองรับไฟล์นามสกุล **`.docx` (แนะนำเพื่อความแม่นยำสูงสุด)** และ **`.pdf`** ขนาดไม่เกิน 200 MB
* ในเอกสารต้องมีหัวข้อ **`References`**, **`เอกสารอ้างอิง`** หรือ **`บรรณานุกรม`** แยกขึ้นบรรทัดใหม่อย่างชัดเจนก่อนเริ่มรายการอ้างอิงท้ายเล่ม

**2. สิ่งที่ระบบตรวจสอบให้:**
* **Cross-Matching:** เช็กว่าทุกจุดที่อ้างอิงในเนื้อหามีรายชื่อท้ายเล่มครบหรือไม่, มีรายการเกินท้ายเล่มที่ไม่ได้อ้างอิงหรือไม่ และปี พ.ศ./ค.ศ. ตรงกันหรือไม่
* **ASCE Formatting:** ตรวจสอบการใช้คำว่า `and` (ห้ามใช้ `&`), การไม่ใส่เครื่องหมายจุลภาค `,` คั่นระหว่างชื่อผู้แต่งกับปีในเนื้อหา, การใส่จุดหลัง `et al.`, การเรียงลำดับตัวอักษร A–Z และเครื่องหมายคำพูดครอบชื่อบทความวารสาร
        """,
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
        "no_cross": "ยอดเยี่ยม! ไม่พบความคลาดเคลื่อน การอ้างอิงในเนื้อหาและบรรณานุกรมท้ายเล่มสอดคล้องกันครบถ้วน",
        "no_format": "ยอดเยี่ยม! ไม่พบข้อผิดพลาดด้านรูปแบบตามเกณฑ์ ASCE เบื้องต้น",
        "no_refs": "ไม่พบรายการบรรณานุกรมในเอกสาร",
        "err_prefix": "ไม่สามารถประมวลผลเอกสารได้:",
        "err_ext": "รองรับเฉพาะไฟล์ .docx และ .pdf เท่านั้น",
        "err_no_heading": "ไม่พบหัวข้อ 'References' หรือ 'เอกสารอ้างอิง' ในไฟล์ กรุณาตรวจสอบว่ามีหัวข้อดังกล่าวขึ้นบรรทัดใหม่ในเอกสาร",
        "sb_header": "📖 ตัวอย่างรูปแบบ ASCE ที่ถูกต้อง",
        "sb_intext_title": "1. การอ้างอิงในเนื้อหา (In-text)",
        "sb_intext_ex": """
* **ผู้แต่ง 1 คน:** `(Smith 2020)` หรือ `Smith (2020)`
* **ผู้แต่ง 2 คน:** `(Smith and Jones 2020)` *(ใช้ and ห้ามใช้ &)*
* **ผู้แต่ง 3 คนขึ้นไป:** `(Smith et al. 2020)` *(มีจุดหลัง al
