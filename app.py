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
ORG_NAME_EN = "Graduate Students, College of Computing | Thesis Formatting & Reference Service"
CONTACT_TH = "หากพบข้อสงสัยเกี่ยวกับรูปแบบการอ้างอิง กรุณาติดต่อเจ้าหน้าที่งานบัณฑิตศึกษา"
CONTACT_EN = "For formatting inquiries, please contact the Graduate Studies Office."

# =========================================================
# 2. ตั้งค่าหน้าเว็บไซต์และ CSS ตกแต่งความสวยงาม
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

    /* แบนเนอร์ส่วนหัว */
    .hero-banner {
        background: linear-gradient(135deg, #0f2942 0%, #1d4e89 100%);
        padding: 2rem 2.2rem;
        border-radius: 14px;
        color: #ffffff;
        margin-bottom: 1.2rem;
        box-shadow: 0 4px 14px rgba(15, 41, 66, 0.15);
    }
    .hero-badge {
        display: inline-block;
        background-color: rgba(255, 255, 255, 0.18);
        color: #e0f2fe;
        padding: 0.25rem 0.75rem;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 500;
        margin-bottom: 0.6rem;
    }
    .hero-title {
        font-size: 1.85rem;
        font-weight: 700;
        margin: 0 0 0.5rem 0;
        line-height: 1.3;
    }
    .hero-subtitle {
        font-size: 1rem;
        color: #dbeafe;
        margin: 0;
        line-height: 1.6;
    }

    /* กล่องการ์ดแสดงตัวเลขสถิติ */
    div[data-testid="stMetric"] {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        padding: 1rem 1.2rem;
        border-radius: 10px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    }

    /* กล่องอัปโหลดไฟล์ */
    section[data-testid="stFileUploader"] {
        background-color: #f8fafc;
        border: 1.5px dashed #94a3b8;
        border-radius: 12px;
        padding: 1rem;
    }

    /* ส่วนท้ายเว็บ */
    .footer-box {
        margin-top: 3rem;
        padding-top: 1rem;
        border-top: 1px solid #e2e8f0;
        text-align: center;
        color: #64748b;
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
        "guide_title": "📌 คำแนะนำการเตรียมไฟล์และขอบเขตการตรวจสอบ (คลิกเพื่ออ่าน)",
        "guide_body": """
**1. การเตรียมไฟล์ก่อนอัปโหลด:**
* รองรับไฟล์นามสกุล **`.docx` (แนะนำเพื่อความแม่นยำสูงสุด)** และ **`.pdf`** ขนาดไม่เกิน 200 MB
* ในเอกสารต้องมีหัวข้อ **`References`**, **`เอกสารอ้างอิง`** หรือ **`บรรณานุกรม`** แยกขึ้นบรรทัดใหม่อย่างชัดเจนก่อนเริ่มรายการอ้างอิงท้ายเล่ม

**2. สิ่งที่ระบบตรวจสอบให้:**
* **Cross-Matching:** เช็กว่าทุกจุดที่อ้างอิงในเนื้อหามีรายชื่อท้ายเล่มครบหรือไม่, มีรายการเกินท้ายเล่มที่ไม่ได้อ้างอิงหรือไม่ และปี พ.ศ./ค.ศ. ตรงกันหรือไม่
* **ASCE Formatting:** ตรวจสอบการใช้คำว่า `and` (ห้ามใช้ `&`), การไม่ใส่เครื่องหมายจุลภาค `,` คั่นระหว่างชื่อผู้แต่งกับปีในเนื้อหา, การใส่จุดหลัง `et al.`, การเรียงลำดับตัวอักษร A–Z และเครื่องหมายคำพูดครอบชื่อบทความวารสาร
        """,
        "upload_label": "📂 อัปโหลดไฟล์ร่างวิทยานิพนธ์ (.docx หรือ .pdf)",
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
* **ผู้แต่ง 3 คนขึ้นไป:** `(Smith et al. 2020)` *(มีจุดหลัง al.)*
* **หลายเรื่องพร้อมกัน:** `(Lee 2019; Smith 2020)`
* **ข้อควรระวัง:** ไม่มีเครื่องหมายจุลภาค `,` คั่นระหว่างชื่อผู้แต่งและปี
        """,
        "sb_ref_title": "2. บรรณานุกรมท้ายเล่ม (References)",
        "sb_ref_ex": """
* **บทความวารสาร:**
  `Smith, J. A., and Jones, B. C. (2020). "Title of the paper." J. Struct. Eng., 146(4), 04020035.`
* **หนังสือ:**
  `Brown, R. D. (2018). Structural analysis, 3rd Ed., Wiley, New York.`
* **การเรียงลำดับ:** เรียงตามตัวอักษรตัวแรกของนามสกุลผู้แต่งคนแรก (A–Z)
        """,
        "sb_privacy_title": "🔒 ความปลอดภัยของข้อมูล",
        "sb_privacy_text": "ไฟล์วิทยานิพนธ์จะถูกประมวลผลในหน่วยความจำชั่วคราวขณะใช้งานเท่านั้น ระบบไม่มีการบันทึกหรือเก็บสำเนาไฟล์ไว้บนเซิร์ฟเวอร์",
        "footer": CONTACT_TH,
        "col_refs": ["ลำดับ", "ผู้แต่ง", "ปี", "ข้อความบรรณานุกรมเต็ม"]
    },
    "en": {
        "org": ORG_NAME_EN,
        "title": "Thesis Reference Validation System (ASCE Standard)",
        "subtitle": "Automatically cross-match in-text citations against the reference list and verify formatting compliance with the American Society of Civil Engineers (ASCE) style guide.",
        "sleep_notice": "💡 **Usage Note:** If you open this link and see a notice screen with a blue **\"Yes, get this app back up!\"** button, simply click that button and wait about 15 seconds for the system to load normally.",
        "guide_title": "📌 File Preparation Guide & Validation Scope (Click to expand)",
        "guide_body": """
**1. File Preparation Before Uploading:**
* Supports **`.docx` (recommended for highest accuracy)** and **`.pdf`** files up to 200 MB.
* Your document must contain a clear standalone heading named **`References`** (or **`เอกสารอ้างอิง`**) before the reference list begins.

**2. What This System Checks:**
* **Cross-Matching:** Detects in-text citations missing from the reference list, uncited entries in the reference list, and publication year mismatches between text and references.
* **ASCE Formatting:** Checks for proper use of `and` instead of `&`, absence of commas between author and year in in-text citations, periods in `et al.`, alphabetical ordering (A–Z), and quotation marks around journal article titles.
        """,
        "upload_label": "📂 Upload Your Thesis Draft (.docx or .pdf)",
        "spinner": "Analyzing document structure and validating citations...",
        "success": "Validation completed for **{filename}**!",
        "m_intext": "In-Text Citations",
        "m_refs": "Reference Entries",
        "m_cross": "Cross-Match Discrepancies",
        "m_format": "ASCE Format Issues",
        "dl_btn": "📥 Download Full Validation Report (Excel)",
        "tab1": "❌ Cross-Match Discrepancies",
        "tab2": "⚠️ ASCE Formatting Issues",
        "tab3": "📋 Parsed Reference List",
        "no_cross": "Great job! No discrepancies found. All in-text citations and reference entries match.",
        "no_format": "Great job! No basic ASCE formatting issues were detected.",
        "no_refs": "No reference entries were detected in the document.",
        "err_prefix": "Unable to process document:",
        "err_ext": "Only .docx and .pdf files are supported.",
        "err_no_heading": "Could not find a standalone 'References' heading. Please ensure your reference section starts with a clear 'References' heading on its own line.",
        "sb_header": "📖 Quick ASCE Style Guide",
        "sb_intext_title": "1. In-Text Citations",
        "sb_intext_ex": """
* **One author:** `(Smith 2020)` or `Smith (2020)`
* **Two authors:** `(Smith and Jones 2020)` *(Use 'and', never '&')*
* **Three or more:** `(Smith et al. 2020)` *(Include period after 'al.')*
* **Multiple works:** `(Lee 2019; Smith 2020)`
* **Note:** Do NOT put a comma between the author's name and the year.
        """,
        "sb_ref_title": "2. Reference List Entries",
        "sb_ref_ex": """
* **Journal Article:**
  `Smith, J. A., and Jones, B. C. (2020). "Title of the paper." J. Struct. Eng., 146(4), 04020035.`
* **Book:**
  `Brown, R. D. (2018). Structural analysis, 3rd Ed., Wiley, New York.`
* **Ordering:** Arrange entries alphabetically by the first author's surname (A–Z).
        """,
        "sb_privacy_title": "🔒 Data Privacy Notice",
        "sb_privacy_text": "Uploaded documents are processed strictly in temporary session memory (RAM) and are never stored or retained on any server.",
        "footer": CONTACT_EN,
        "col_refs": ["No.", "Authors", "Year", "Full Reference Text"]
    }
}


# =========================================================
# 4. คลาสประมวลผลและตรวจสอบมาตรฐาน ASCE (รองรับ 2 ภาษา)
# =========================================================
class ASCEReferenceChecker:
    def __init__(self, file_name: str, file_bytes: bytes, lang: str = "th"):
        self.file_name = file_name
        self.file_bytes = file_bytes
        self.lang = lang
        self.ext = Path(file_name).suffix.lower()

        self.full_text = ""
        self.body_text = ""
        self.ref_text = ""

        self.in_text_citations = []
        self.references = []
        self.format_issues = []
        self.cross_match_issues = []

    def extract_text(self):
        if self.ext == ".docx":
            doc = docx.Document(io.BytesIO(self.file_bytes))
            paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
            self.full_text = "\n".join(paragraphs)
        elif self.ext == ".pdf":
            doc = fitz.open(stream=self.file_bytes, filetype="pdf")
            pages_text = []
            for page in doc:
                text = page.get_text("text")
                lines = [
                    line.strip() for line in text.split("\n")
                    if line.strip() and not re.fullmatch(r"[-–]?\s*\d+\s*[-–]?", line.strip())
                ]
                pages_text.append("\n".join(lines))
            self.full_text = "\n".join(pages_text)
        else:
            raise ValueError(TEXTS[self.lang]["err_ext"])

    def split_sections(self):
        pattern = re.compile(
            r"\n\s*(?:REFERENCES|เอกสารอ้างอิง|บรรณานุกรม)\s*:?\s*\n",
            re.IGNORECASE
        )
        matches = list(pattern.finditer(self.full_text))
        if not matches:
            raise ValueError(TEXTS[self.lang]["err_no_heading"])

        last_match = matches[-1]
        self.body_text = self.full_text[:last_match.start()]
        raw_ref_text = self.full_text[last_match.end():]

        appendix_match = re.search(
            r"\n\s*(?:APPENDIX|ภาคผนวก|ประวัติผู้เขียน|BIOGRAPHY)\b",
            raw_ref_text,
            re.IGNORECASE
        )
        if appendix_match:
            raw_ref_text = raw_ref_text[:appendix_match.start()]

        self.ref_text = raw_ref_text.strip()

    def _add_format_issue(self, loc_th, loc_en, snippet, type_th, type_en, fix_th, fix_en):
        if self.lang == "th":
            self.format_issues.append({
                "ตำแหน่ง": loc_th,
                "ข้อความที่พบ": snippet,
                "ประเภทปัญหา": type_th,
                "คำแนะนำการแก้ไข": fix_th
            })
        else:
            self.format_issues.append({
                "Location": loc_en,
                "Detected Text": snippet,
                "Issue Type": type_en,
                "Suggested Fix": fix_en
            })

    def _add_cross_issue(self, type_th, type_en, item_text, detail_th, detail_en):
        if self.lang == "th":
            self.cross_match_issues.append({
                "ประเภทความคลาดเคลื่อน": type_th,
                "รายการที่พบ": item_text,
                "รายละเอียด": detail_th
            })
        else:
            self.cross_match_issues.append({
                "Discrepancy Type": type_en,
                "Item": item_text,
                "Details": detail_en
            })

    def parse_in_text_citations(self):
        paren_blocks = re.findall(r"\(([^()]*?\b(?:19|20|25)\d{2}[a-zก-ฮ]?\b[^()]*?)\)", self.body_text)

        for block in paren_blocks:
            if any(word in block.lower() for word in ["eq.", "fig.", "table", "สมการ", "ตาราง", "รูปที่"]):
                continue

            parts = [p.strip() for p in block.split(";")]
            for part in parts:
                match = re.match(
                    r"^([A-Zก-๙][A-Za-zก-๙\-\.'\s&]+?)[,\s]+((?:19|20|25)\d{2}[a-zก-ฮ]?)$",
                    part
                )
                if match:
                    author_part = match.group(1).strip()
                    year_part = match.group(2).strip()
                    raw_cite = f"({part})"

                    self._check_in_text_asce_style(raw_cite, part, author_part)
                    key = self._make_key(author_part, year_part)
                    self.in_text_citations.append({
                        "raw": raw_cite,
                        "author": author_part,
                        "year": year_part,
                        "key": key
                    })

        narrative_pattern = re.compile(
            r"\b([A-Z][a-zA-Z\-']+(?:\s+(?:and|&)\s+[A-Z][a-zA-Z\-']+|\s+et\s+al\.?)?)\s*\(((?:19|20|25)\d{2}[a-z]?)\)"
        )
        for m in narrative_pattern.finditer(self.body_text):
            author_part = m.group(1).strip()
            year_part = m.group(2).strip()
            raw_cite = f"{author_part} ({year_part})"

            self._check_in_text_asce_style(raw_cite, raw_cite, author_part)
            key = self._make_key(author_part, year_part)
            self.in_text_citations.append({
                "raw": raw_cite,
                "author": author_part,
                "year": year_part,
                "key": key
            })

    def _check_in_text_asce_style(self, raw_cite: str, inner_text: str, author_part: str):
        if "&" in author_part:
            self._add_format_issue(
                "ในเนื้อหา (In-text)", "In-text Citation",
                raw_cite,
                "ใช้เครื่องหมาย '&' ผิดรูปแบบ ASCE", "Invalid '&' symbol used",
                "เปลี่ยน '&' เป็น 'and' (เช่น Smith and Jones 2020)",
                "Replace '&' with 'and' (e.g., Smith and Jones 2020)"
            )

        if re.search(r",\s*(?:19|20|25)\d{2}", inner_text):
            self._add_format_issue(
                "ในเนื้อหา (In-text)", "In-text Citation",
                raw_cite,
                "มีเครื่องหมายจุลภาค (,) คั่นก่อนปี พ.ศ./ค.ศ.", "Comma (,) before publication year",
                "ลบเครื่องหมายจุลภาคออก (ASCE ใช้รูปแบบ Author Year เช่น Smith 2020)",
                "Remove the comma (ASCE uses Author Year format, e.g., Smith 2020)"
            )

        if re.search(r"\bet\s+al\b(?!\.)", author_part):
            self._add_format_issue(
                "ในเนื้อหา (In-text)", "In-text Citation",
                raw_cite,
                "คำว่า 'et al.' ขาดเครื่องหมายจุด", "Missing period in 'et al.'",
                "เติมจุดหลัง al เป็น 'et al.'",
                "Add a period after 'al' to form 'et al.'"
            )

    def parse_references(self):
        entries = re.split(
            r"\n(?=[A-Zก-๙][^\n]{1,120}?\((?:19|20|25)\d{2}[a-zก-ฮ]?\)\.)",
            self.ref_text
        )
        entries = [re.sub(r"\s+", " ", e).strip() for e in entries if e.strip()]

        prev_sort_key = ""
        for idx, entry in enumerate(entries, start=1):
            match = re.match(
                r"^(.+?)\s*\(((?:19|20|25)\d{2}[a-zก-ฮ]?)\)\.\s*(.*)$",
                entry
            )
            if not match:
                self._add_format_issue(
                    f"บรรณานุกรมลำดับที่ {idx}", f"Reference #{idx}",
                    entry[:90] + "...",
                    "โครงสร้างส่วนต้นไม่ตรงตามรูปแบบ ASCE", "Invalid leading structure",
                    "ตรวจสอบรูปแบบ: ผู้แต่ง (ปี). ชื่อเรื่อง. แหล่งพิมพ์.",
                    "Verify leading format: Author(s) (Year). Title. Source."
                )
                continue

            authors = match.group(1).strip()
            year = match.group(2).strip()
            rest = match.group(3).strip()
            key = self._make_key(authors, year)

            self.references.append({
                "index": idx,
                "raw": entry,
                "authors": authors,
                "year": year,
                "key": key
            })

            current_sort_key = authors.lower()
            if prev_sort_key and current_sort_key < prev_sort_key:
                self._add_format_issue(
                    f"บรรณานุกรมลำดับที่ {idx}", f"Reference #{idx}",
                    f"{authors} ({year})",
                    "เรียงลำดับตัวอักษรไม่ถูกต้อง", "Incorrect alphabetical order",
                    "ASCE กำหนดให้เรียงบรรณานุกรมตามตัวอักษรตัวแรกของนามสกุลผู้แต่งคนแรก (A-Z)",
                    "Arrange references alphabetically by the first author's surname (A–Z)"
                )
            prev_sort_key = current_sort_key

            if "&" in authors:
                self._add_format_issue(
                    f"บรรณานุกรมลำดับที่ {idx}", f"Reference #{idx}",
                    f"{authors} ({year})",
                    "ใช้เครื่องหมาย '&' ในรายชื่อผู้แต่ง", "Invalid '&' in author list",
                    "เปลี่ยน '&' เป็น 'and' ก่อนชื่อผู้แต่งคนสุดท้าย",
                    "Replace '&' with 'and' before the last author's name"
                )

            if re.search(r"\d+\(\d+\)", rest) and not ('"' in rest or "“" in rest):
                self._add_format_issue(
                    f"บรรณานุกรมลำดับที่ {idx}", f"Reference #{idx}",
                    entry[:85] + "...",
                    "บทความวารสารอาจขาดเครื่องหมายคำพูด (\"...\") ครอบชื่อบทความ",
                    "Missing quotation marks around journal article title",
                    "ASCE กำหนดให้ใส่เครื่องหมายคำพูดครอบชื่อบทความวารสาร เช่น \"Article title.\"",
                    "Enclose journal article titles in quotation marks, e.g., \"Article title.\""
                )

    def cross_match(self):
        in_text_keys = {item["key"] for item in self.in_text_citations}
        ref_keys = {item["key"] for item in self.references}

        ref_by_author = {}
        for r in self.references:
            author_norm = r["key"][0]
            ref_by_author.setdefault(author_norm, []).append(r["year"])

        in_text_by_author = {}
        for c in self.in_text_citations:
            author_norm = c["key"][0]
            in_text_by_author.setdefault(author_norm, []).append(c["year"])

        seen_missing = set()
        for item in self.in_text_citations:
            k = item["key"]
            if k not in ref_keys and k not in seen_missing:
                seen_missing.add(k)
                author_norm, year = k
                if author_norm in ref_by_author:
                    years_found = ", ".join(set(ref_by_author[author_norm]))
                    self._add_cross_issue(
                        "ปี พ.ศ./ค.ศ. ไม่ตรงกัน (Year Mismatch)", "Year Mismatch",
                        item["raw"],
                        f"ในเนื้อหาระบุปี {year} แต่ในบรรณานุกรมพบผู้แต่งนี้ในปี {years_found}",
                        f"Cited as {year} in text, but listed under year {years_found} in References"
                    )
                else:
                    self._add_cross_issue(
                        "ตกหล่นในบรรณานุกรม (Missing in References)", "Missing in References",
                        item["raw"],
                        "มีการอ้างอิงในเนื้อหา แต่ไม่พบรายชื่อในเอกสารอ้างอิงท้ายเล่ม",
                        "Cited in text, but no matching entry found in the Reference list"
                    )

        for item in self.references:
            k = item["key"]
            if k not in in_text_keys:
                author_norm, year = k
                if author_norm not in in_text_by_author:
                    self._add_cross_issue(
                        "รายการเกินท้ายเล่ม (Uncited Reference)", "Uncited Reference",
                        f"{item['authors']} ({item['year']})",
                        "ปรากฏในเอกสารอ้างอิงท้ายเล่ม แต่ไม่พบการอ้างอิงในเนื้อหา",
                        "Listed in References, but never cited in the main text"
                    )

    def _make_key(self, author_str: str, year_str: str):
        first_token = re.split(r"[,\s\.]+", author_str.strip())[0].lower()
        return (first_token, year_str.lower())

    def run(self):
        self.extract_text()
        self.split_sections()
        self.parse_in_text_citations()
        self.parse_references()
        self.cross_match()

    def build_excel_bytes(self) -> bytes:
        t = TEXTS[self.lang]
        if self.lang == "th":
            summary_data = [
                {"หัวข้อ": "ชื่อไฟล์ที่ตรวจสอบ", "รายละเอียด": self.file_name},
                {"หัวข้อ": t["m_intext"], "รายละเอียด": len(self.in_text_citations)},
                {"หัวข้อ": t["m_refs"], "รายละเอียด": len(self.references)},
                {"หัวข้อ": t["m_cross"], "รายละเอียด": len(self.cross_match_issues)},
                {"หัวข้อ": t["m_format"], "รายละเอียด": len(self.format_issues)},
            ]
            empty_df = pd.DataFrame([{"ผลลัพธ์": "ไม่พบข้อผิดพลาด"}])
        else:
            summary_data = [
                {"Metric": "Checked File", "Value": self.file_name},
                {"Metric": t["m_intext"], "Value": len(self.in_text_citations)},
                {"Metric": t["m_refs"], "Value": len(self.references)},
                {"Metric": t["m_cross"], "Value": len(self.cross_match_issues)},
                {"Metric": t["m_format"], "Value": len(self.format_issues)},
            ]
            empty_df = pd.DataFrame([{"Result": "No issues found"}])

        df_summary = pd.DataFrame(summary_data)
        df_cross = pd.DataFrame(self.cross_match_issues) if self.cross_match_issues else empty_df
        df_format = pd.DataFrame(self.format_issues) if self.format_issues else empty_df
        df_refs = pd.DataFrame(self.references)[["index", "authors", "year", "raw"]] if self.references else pd.DataFrame()
        if not df_refs.empty:
            df_refs.columns = t["col_refs"]

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_summary.to_excel(writer, sheet_name="Summary", index=False)
            df_cross.to_excel(writer, sheet_name="Cross_Match_Errors", index=False)
            df_format.to_excel(writer, sheet_name="ASCE_Format_Issues", index=False)
            df_refs.to_excel(writer, sheet_name="Parsed_References", index=False)
        return buffer.getvalue()


# =========================================================
# 5. ส่วนติดต่อผู้ใช้งาน (Web Interface)
# =========================================================

# ตัวสลับภาษาด้านบนขวา
top_left, top_right = st.columns([4, 1.5])
with top_right:
    lang_choice = st.radio(
        "Language / ภาษา",
        options=["🇹🇭 ภาษาไทย", "🇺🇸 English"],
        horizontal=True,
        label_visibility="collapsed"
    )

lang = "th" if "ไทย" in lang_choice else "en"
t = TEXTS[lang]

# แถบด้านข้าง (Sidebar) แสดงคู่มือและตัวอย่าง ASCE
with st.sidebar:
    st.header(t["sb_header"])
    st.subheader(t["sb_intext_title"])
    st.markdown(t["sb_intext_ex"])
    st.divider()
    st.subheader(t["sb_ref_title"])
    st.markdown(t["sb_ref_ex"])
    st.divider()
    st.info(f"**{t['sb_privacy_title']}**\n\n{t['sb_privacy_text']}")

# แบนเนอร์หัวข้อหลัก
st.markdown(f"""
<div class="hero-banner">
    <div class="hero-badge">🎓 {t["org"]}</div>
    <div class="hero-title">{t["title"]}</div>
    <p class="hero-subtitle">{t["subtitle"]}</p>
</div>
""", unsafe_allow_html=True)

# กล่องข้อความแจ้งเตือนกรณีระบบเข้าสู่โหมดพัก
st.info(t["sleep_notice"])

# กล่องคำแนะนำการเตรียมไฟล์
with st.expander(t["guide_title"], expanded=False):
    st.markdown(t["guide_body"])

# กล่องอัปโหลดไฟล์
uploaded_file = st.file_uploader(
    t["upload_label"],
    type=["docx", "pdf"]
)

if uploaded_file is not None:
    try:
        with st.spinner(t["spinner"]):
            checker = ASCEReferenceChecker(uploaded_file.name, uploaded_file.getvalue(), lang=lang)
            checker.run()

        st.success(t["success"].format(filename=uploaded_file.name))

        # แสดงการ์ดตัวเลขสรุป 4 ช่อง
        col1, col2, col3, col4 = st.columns(4)
        col1.metric(t["m_intext"], len(checker.in_text_citations))
        col2.metric(t["m_refs"], len(checker.references))
        col3.metric(t["m_cross"], len(checker.cross_match_issues))
        col4.metric(t["m_format"], len(checker.format_issues))

        st.write("")

        # ปุ่มดาวน์โหลดรายงาน Excel
        excel_data = checker.build_excel_bytes()
        report_filename = f"ASCE_Report_{Path(uploaded_file.name).stem}.xlsx"
        st.download_button(
            label=t["dl_btn"],
            data=excel_data,
            file_name=report_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

        st.divider()

        # แสดงตารางรายละเอียดแยกตามแท็บ
        tab1, tab2, tab3 = st.tabs([
            f"{t['tab1']} ({len(checker.cross_match_issues)})",
            f"{t['tab2']} ({len(checker.format_issues)})",
            f"{t['tab3']} ({len(checker.references)})"
        ])

        with tab1:
            if checker.cross_match_issues:
                st.dataframe(pd.DataFrame(checker.cross_match_issues), use_container_width=True)
            else:
                st.info(t["no_cross"])

        with tab2:
            if checker.format_issues:
                st.dataframe(pd.DataFrame(checker.format_issues), use_container_width=True)
            else:
                st.info(t["no_format"])

        with tab3:
            if checker.references:
                df_refs_display = pd.DataFrame(checker.references)[["index", "authors", "year", "raw"]]
                df_refs_display.columns = t["col_refs"]
                st.dataframe(df_refs_display, use_container_width=True)
            else:
                st.warning(t["no_refs"])

    except Exception as e:
        st.error(f"{t['err_prefix']} {e}")

# ส่วนท้ายเว็บ
st.markdown(f"""
<div class="footer-box">
    {t["footer"]}
</div>
""", unsafe_allow_html=True)
