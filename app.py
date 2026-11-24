import io
import re
from pathlib import Path
import docx
import fitz
import pandas as pd
import streamlit as st

ORG_NAME_TH = "บัณฑิตศึกษา วิทยาลัยการคอมพิวเตอร์ | ระบบบริการตรวจสอบรูปแบบวิทยานิพนธ์"
ORG_NAME_EN = "Graduate Studies, College of Computing | Thesis Formatting & Reference Service"

st.set_page_config(
    page_title="ASCE Reference Checker | ระบบตรวจสอบเอกสารอ้างอิง",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Thai:wght@300;400;500;600;700&family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'IBM Plex Sans Thai', 'Inter', sans-serif; }
    .stApp, [data-testid="stAppViewContainer"] { background-color: #d9ecf0 !important; }
    [data-testid="stHeader"] { background-color: rgba(217, 236, 240, 0.85) !important; }
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
    .hero-title { font-size: 1.85rem; font-weight: 700; margin: 0 0 0.5rem 0; line-height: 1.3; color: #ffffff; }
    .hero-subtitle { font-size: 1rem; color: #f1f5f9; margin: 0; line-height: 1.6; }
    div[data-testid="stExpander"] {
        background-color: #ffffff !important;
        border: 2px solid #33c0ca !important;
        border-radius: 12px !important;
        overflow: hidden;
        margin-bottom: 1.2rem;
        box-shadow: 0 2px 8px rgba(15, 41, 66, 0.06);
    }
    div[data-testid="stExpander"] > details > summary {
        background-color: #33c0ca !important;
        color: #0f2942 !important;
        font-weight: 700 !important;
        padding: 0.75rem 1rem !important;
    }
    div[data-testid="stExpander"] > details > summary:hover { background-color: #29aeb8 !important; }
    div[data-testid="stExpander"] > details > summary p,
    div[data-testid="stExpander"] > details > summary span {
        font-weight: 700 !important;
        font-size: 1.02rem !important;
        color: #0f2942 !important;
    }
    div[data-testid="stExpander"] > details > summary svg { fill: #0f2942 !important; color: #0f2942 !important; }
    section[data-testid="stSidebar"] { border-right: 3px solid #33c0ca; }
    div[data-testid="stMetric"] {
        background-color: #ffffff;
        border: 1px solid #cbd5e1;
        border-left: 5px solid #33c0ca;
        padding: 1rem 1.2rem;
        border-radius: 10px;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.05);
    }
    div[data-testid="column"]:nth-child(3) div[data-testid="stMetric"],
    div[data-testid="column"]:nth-child(4) div[data-testid="stMetric"] { border-left: 5px solid #b52324; }
    div[data-testid="stFileUploader"] label[data-testid="stWidgetLabel"] p {
        font-weight: 700 !important;
        font-size: 1.06rem !important;
        color: #0f2942 !important;
    }
    section[data-testid="stFileUploader"] {
        background-color: #ffffff;
        border: 2px dashed #33c0ca;
        border-radius: 12px;
        padding: 1rem;
    }
    section[data-testid="stFileUploader"]:hover { border-color: #b52324; }
    div[data-testid="stDownloadButton"] > button {
        background-color: #b52324 !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 8px;
        font-weight: 600;
        padding: 0.55rem 1.2rem;
    }
    div[data-testid="stDownloadButton"] > button:hover { background-color: #33c0ca !important; color: #0f2942 !important; }
    button[data-baseweb="tab"][aria-selected="true"] {
        color: #b52324 !important;
        border-bottom-color: #b52324 !important;
        font-weight: 600;
    }
    div[data-testid="stAlert"] {
        background-color: #ffffff !important;
        border: 1.5px solid #33c0ca !important;
        border-radius: 10px !important;
    }
    .bottom-divider { margin-top: 3rem; margin-bottom: 1.2rem; border: none; border-top: 2px solid #33c0ca; }
</style>
""", unsafe_allow_html=True)

TEXTS = {
    "th": {
        "org": ORG_NAME_TH,
        "title": "ระบบตรวจสอบเอกสารอ้างอิงวิทยานิพนธ์ (มาตรฐาน ASCE)",
        "subtitle": "ตรวจสอบความสอดคล้องระหว่างการอ้างอิงในเนื้อหา (In-text Citations) กับบรรณานุกรมท้ายเล่ม (References / Bibliography) พร้อมตรวจรูปแบบการพิมพ์ตามเกณฑ์ American
