import io
import re
from pathlib import Path
import docx
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st

# ตั้งค่าหน้าเว็บไซต์
st.set_page_config(
    page_title="ระบบตรวจสอบเอกสารอ้างอิง (ASCE)",
    page_icon="📚",
    layout="wide"
)


class ASCEReferenceChecker:
    def __init__(self, file_name: str, file_bytes: bytes):
        self.file_name = file_name
        self.file_bytes = file_bytes
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
            raise ValueError("รองรับเฉพาะไฟล์ .docx และ .pdf เท่านั้น")

    def split_sections(self):
        pattern = re.compile(
            r"\n\s*(?:REFERENCES|เอกสารอ้างอิง|บรรณานุกรม)\s*:?\s*\n",
            re.IGNORECASE
        )
        matches = list(pattern.finditer(self.full_text))
        if not matches:
            raise ValueError(
                "ไม่พบหัวข้อ 'References' หรือ 'เอกสารอ้างอิง' ในไฟล์ กรุณาตรวจสอบชื่อหัวข้อในเอกสาร"
            )

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
            self.format_issues.append({
                "ตำแหน่ง": "ในเนื้อหา (In-text)",
                "ข้อความที่พบ": raw_cite,
                "ประเภทปัญหา": "ใช้เครื่องหมาย '&' ผิดรูปแบบ ASCE",
                "คำแนะนำการแก้ไข": "เปลี่ยน '&' เป็น 'and' (เช่น Smith and Jones 2020)"
            })

        if re.search(r",\s*(?:19|20|25)\d{2}", inner_text):
            self.format_issues.append({
                "ตำแหน่ง": "ในเนื้อหา (In-text)",
                "ข้อความที่พบ": raw_cite,
                "ประเภทปัญหา": "มีเครื่องหมายจุลภาค (,) คั่นก่อนปี พ.ศ./ค.ศ.",
                "คำแนะนำการแก้ไข": "ลบเครื่องหมายจุลภาคออก (ASCE ใช้รูปแบบ Author Year เช่น Smith 2020)"
            })

        if re.search(r"\bet\s+al\b(?!\.)", author_part):
            self.format_issues.append({
                "ตำแหน่ง": "ในเนื้อหา (In-text)",
                "ข้อความที่พบ": raw_cite,
                "ประเภทปัญหา": "คำว่า 'et al.' ขาดเครื่องหมายจุด",
                "คำแนะนำการแก้ไข": "เติมจุดหลัง al เป็น 'et al.'"
            })

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
                self.format_issues.append({
                    "ตำแหน่ง": f"บรรณานุกรมลำดับที่ {idx}",
                    "ข้อความที่พบ": entry[:100] + "...",
                    "ประเภทปัญหา": "โครงสร้างส่วนต้นไม่ตรงตามรูปแบบ ASCE",
                    "คำแนะนำการแก้ไข": "ตรวจสอบรูปแบบ: ผู้แต่ง (ปี). ชื่อเรื่อง. แหล่งพิมพ์."
                })
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
                self.format_issues.append({
                    "ตำแหน่ง": f"บรรณานุกรมลำดับที่ {idx}",
                    "ข้อความที่พบ": f"{authors} ({year})",
                    "ประเภทปัญหา": "เรียงลำดับตัวอักษรไม่ถูกต้อง",
                    "คำแนะนำการแก้ไข": "ASCE กำหนดให้เรียงบรรณานุกรมตามตัวอักษรตัวแรกของนามสกุลผู้แต่งคนแรก (A-Z)"
                })
            prev_sort_key = current_sort_key

            if "&" in authors:
                self.format_issues.append({
                    "ตำแหน่ง": f"บรรณานุกรมลำดับที่ {idx}",
                    "ข้อความที่พบ": f"{authors} ({year})",
                    "ประเภทปัญหา": "ใช้เครื่องหมาย '&' ในรายชื่อผู้แต่ง",
                    "คำแนะนำการแก้ไข": "เปลี่ยน '&' เป็น 'and' ก่อนชื่อผู้แต่งคนสุดท้าย"
                })

            if re.search(r"\d+\(\d+\)", rest) and not ('"' in rest or "“" in rest):
                self.format_issues.append({
                    "ตำแหน่ง": f"บรรณานุกรมลำดับที่ {idx}",
                    "ข้อความที่พบ": entry[:90] + "...",
                    "ประเภทปัญหา": "บทความวารสารอาจขาดเครื่องหมายคำพูด (\"...\") ครอบชื่อบทความ",
                    "คำแนะนำการแก้ไข": "ASCE กำหนดให้ใส่เครื่องหมายคำพูดครอบชื่อบทความวารสาร เช่น \"Article title.\""
                })

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
                    self.cross_match_issues.append({
                        "ประเภทความคลาดเคลื่อน": "ปี พ.ศ./ค.ศ. ไม่ตรงกัน (Year Mismatch)",
                        "รายการที่พบ": item["raw"],
                        "รายละเอียด": f"ในเนื้อหาระบุปี {year} แต่ในบรรณานุกรมพบผู้แต่งนี้ในปี {', '.join(set(ref_by_author[author_norm]))}"
                    })
                else:
                    self.cross_match_issues.append({
                        "ประเภทความคลาดเคลื่อน": "ตกหล่นในบรรณานุกรม (Missing in References)",
                        "รายการที่พบ": item["raw"],
                        "รายละเอียด": "มีการอ้างอิงในเนื้อหา แต่ไม่พบรายชื่อในเอกสารอ้างอิงท้ายเล่ม"
                    })

        for item in self.references:
            k = item["key"]
            if k not in in_text_keys:
                author_norm, year = k
                if author_norm not in in_text_by_author:
                    self.cross_match_issues.append({
                        "ประเภทความคลาดเคลื่อน": "รายการเกินท้ายเล่ม (Uncited Reference)",
                        "รายการที่พบ": f"{item['authors']} ({item['year']})",
                        "รายละเอียด": "ปรากฏในเอกสารอ้างอิงท้ายเล่ม แต่ไม่พบการอ้างอิงในเนื้อหา"
                    })

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
        summary_data = [
            {"หัวข้อ": "ชื่อไฟล์ที่ตรวจสอบ", "จำนวน / รายละเอียด": self.file_name},
            {"หัวข้อ": "จำนวนการอ้างอิงในเนื้อหาทั้งหมด (จุด)", "จำนวน / รายละเอียด": len(self.in_text_citations)},
            {"หัวข้อ": "จำนวนรายการบรรณานุกรมท้ายเล่ม (รายการ)", "จำนวน / รายละเอียด": len(self.references)},
            {"หัวข้อ": "จุดที่ไม่สอดคล้องกัน (Cross-Match Errors)", "จำนวน / รายละเอียด": len(self.cross_match_issues)},
            {"หัวข้อ": "จุดที่ผิดรูปแบบมาตรฐาน ASCE (Format Issues)", "จำนวน / รายละเอียด": len(self.format_issues)},
        ]

        df_summary = pd.DataFrame(summary_data)
        df_cross = pd.DataFrame(self.cross_match_issues) if self.cross_match_issues else pd.DataFrame([{"ผลลัพธ์": "ไม่พบข้อผิดพลาด"}])
        df_format = pd.DataFrame(self.format_issues) if self.format_issues else pd.DataFrame([{"ผลลัพธ์": "ไม่พบข้อผิดพลาด"}])
        df_refs = pd.DataFrame(self.references)[["index", "authors", "year", "raw"]] if self.references else pd.DataFrame()

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_summary.to_excel(writer, sheet_name="Summary", index=False)
            df_cross.to_excel(writer, sheet_name="Cross_Match_Errors", index=False)
            df_format.to_excel(writer, sheet_name="ASCE_Format_Issues", index=False)
            df_refs.to_excel(writer, sheet_name="Parsed_References", index=False)
        return buffer.getvalue()


# =========================================================
# ส่วนติดต่อผู้ใช้งาน (Web Interface)
# =========================================================
st.title("📚 ระบบตรวจสอบเอกสารอ้างอิงวิทยานิพนธ์ (มาตรฐาน ASCE)")
st.markdown(
    "อัปโหลดไฟล์ร่างวิทยานิพนธ์ (`.docx` หรือ `.pdf`) เพื่อตรวจสอบความสอดคล้องระหว่างการอ้างอิงในเนื้อหากับบรรณานุกรมท้ายเล่ม และตรวจสอบรูปแบบการพิมพ์ตามเกณฑ์ **ASCE (American Society of Civil Engineers)**"
)

uploaded_file = st.file_uploader(
    "ลากไฟล์มาวางที่นี่ หรือกดปุ่ม Browse files เพื่อเลือกไฟล์ (.docx / .pdf)",
    type=["docx", "pdf"]
)

if uploaded_file is not None:
    try:
        with st.spinner("กำลังวิเคราะห์เอกสารและตรวจสอบรายการอ้างอิง..."):
            checker = ASCEReferenceChecker(uploaded_file.name, uploaded_file.getvalue())
            checker.run()

        st.success(f"ตรวจสอบไฟล์ **{uploaded_file.name}** เสร็จสมบูรณ์")

        # แสดงตัวเลขสรุป 4 ช่อง
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("การอ้างอิงในเนื้อหา (จุด)", len(checker.in_text_citations))
        col2.metric("บรรณานุกรมท้ายเล่ม (รายการ)", len(checker.references))
        col3.metric("จุดไม่สอดคล้องกัน (Cross-Match)", len(checker.cross_match_issues))
        col4.metric("จุดผิดรูปแบบ ASCE (Format)", len(checker.format_issues))

        # ปุ่มดาวน์โหลดรายงาน Excel
        excel_data = checker.build_excel_bytes()
        report_filename = f"ASCE_Report_{Path(uploaded_file.name).stem}.xlsx"
        st.download_button(
            label="📥 ดาวน์โหลดรายงานผลการตรวจสอบ (Excel)",
            data=excel_data,
            file_name=report_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

        st.divider()

        # แสดงตารางรายละเอียดแยกตามแท็บ
        tab1, tab2, tab3 = st.tabs([
            f"❌ ความคลาดเคลื่อนเนื้อหา-ท้ายเล่ม ({len(checker.cross_match_issues)})",
            f"⚠️ จุดผิดรูปแบบ ASCE ({len(checker.format_issues)})",
            f"📋 รายชื่อบรรณานุกรมทั้งหมด ({len(checker.references)})"
        ])

        with tab1:
            if checker.cross_match_issues:
                st.dataframe(pd.DataFrame(checker.cross_match_issues), use_container_width=True)
            else:
                st.info("ไม่พบความคลาดเคลื่อน การอ้างอิงในเนื้อหาและบรรณานุกรมท้ายเล่มสอดคล้องกันครบถ้วน")

        with tab2:
            if checker.format_issues:
                st.dataframe(pd.DataFrame(checker.format_issues), use_container_width=True)
            else:
                st.info("ไม่พบข้อผิดพลาดด้านรูปแบบตามเกณฑ์ ASCE เบื้องต้น")

        with tab3:
            if checker.references:
                df_refs_display = pd.DataFrame(checker.references)[["index", "authors", "year", "raw"]]
                df_refs_display.columns = ["ลำดับ", "ผู้แต่ง", "ปี", "ข้อความบรรณานุกรมเต็ม"]
                st.dataframe(df_refs_display, use_container_width=True)
            else:
                st.warning("ไม่พบรายการบรรณานุกรม")

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการอ่านเอกสาร: {e}")
