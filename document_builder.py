#!/usr/bin/env python3
"""Interactive IT security report builder for Word (.docx) and Excel (.xlsx)."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date as dt_date
from pathlib import Path
from typing import List, Tuple

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

COLORS = {
    "primary_dark_blue": "1F3864",
    "secondary_blue": "2E5FA3",
    "accent_light_blue": "D6E4F7",
    "light_gray_fill": "F2F2F2",
    "placeholder_gray": "888888",
    "body_text": "1A1A1A",
    "white": "FFFFFF",
    "warning_yellow": "FFF8E1",
    "error_red_pink": "FEECEC",
    "light_border": "D9D9D9",
}


@dataclass
class ReportData:
    report_title: str = "General Title"
    site_name: str = "Site Name"
    environment_name: str = "Site / Environment Name"
    date: str = field(default_factory=lambda: dt_date.today().isoformat())
    classification: str = "Confidential — Internal Use Only"
    site_labels: Tuple[str, str] = ("NDH", "HPH")
    recommendation_rows: List[Tuple[str, str]] = field(default_factory=list)
    status_value: str = "In Progress"
    next_steps: List[str] = field(default_factory=list)
    prepared_by: str = "IT Security team"


def ask(prompt: str, default: str = "") -> str:
    v = input(f"{prompt}{f' [{default}]' if default else ''}: ").strip()
    return v or default


def ask_int(prompt: str, default: int) -> int:
    while True:
        v = ask(prompt, str(default))
        try:
            return int(v)
        except ValueError:
            print("Enter a valid number.")


def collect_data() -> tuple[str, ReportData, int]:
    print("=== IT Security Report Builder ===")
    doc_type = ask("Output type (word/excel)", "word").lower()
    while doc_type not in {"word", "excel"}:
        doc_type = ask("Please choose word or excel", "word").lower()

    data = ReportData(
        report_title=ask("Report title", "General Title"),
        site_name=ask("Site name", "Site Name"),
        environment_name=ask("Environment name", "Site / Environment Name"),
        date=ask("Assessment date", dt_date.today().isoformat()),
        classification=ask("Classification", "Confidential — Internal Use Only"),
        status_value=ask("Status (In Progress/Completed/Pending Review)", "In Progress"),
        prepared_by=ask("Prepared by", "IT Security team"),
    )
    rows = ask_int("How many recommendation rows", 2)
    for i in range(rows):
        left = ask(f"Recommendation {i+1}", "[ Recommendation ]")
        right = ask(f"Response {i+1}", "[ Current Limitation / Response ]")
        data.recommendation_rows.append((left, right))

    for i in range(1, 5):
        data.next_steps.append(ask(f"Next step {i}", f"[ Next step {i} ]" if i < 4 else "[ Schedule follow-up review ]"))

    return doc_type, data, rows


def _set_cell_fill(cell, color_hex: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    tc_pr.append(shd)


def _set_table_borders(table, color_hex: str = "D9D9D9", sz: int = 8) -> None:
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        elem = OxmlElement(f"w:{edge}")
        elem.set(qn("w:val"), "single")
        elem.set(qn("w:sz"), str(sz))
        elem.set(qn("w:space"), "0")
        elem.set(qn("w:color"), color_hex)
        borders.append(elem)
    tbl_pr.append(borders)


def _set_paragraph_bottom_border(paragraph, color_hex: str, size_eighth_pt: int) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size_eighth_pt))
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color_hex)
    pbdr.append(bottom)
    p_pr.append(pbdr)


def _add_styles(doc: Document) -> None:
    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Arial"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    normal.font.size = Pt(10.5)

    def add_paragraph_style(name: str):
        return styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH) if name not in styles else styles[name]

    style_map = {
        "TITLE_MAIN": (26, True, False, "primary_dark_blue"),
        "SUBTITLE": (16, False, True, "secondary_blue"),
        "METADATA": (9, False, False, "placeholder_gray"),
        "HEADING_SECTION": (13, True, False, "primary_dark_blue"),
        "BODY_TEXT": (10.5, False, False, "body_text"),
        "PLACEHOLDER_TEXT": (10.5, False, True, "placeholder_gray"),
        "SITE_LABEL": (11, True, False, "primary_dark_blue"),
        "FOOTER_TEXT": (8, False, False, "placeholder_gray"),
    }
    for n, (size, bold, italic, color) in style_map.items():
        s = add_paragraph_style(n)
        s.font.name = "Arial"; s._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
        s.font.size = Pt(size); s.font.bold = bold; s.font.italic = italic
        s.font.color.rgb = RGBColor.from_string(COLORS[color])


def build_word_report(data: ReportData, output_path: Path) -> Path:
    doc = Document(); _add_styles(doc)
    sec = doc.sections[0]
    sec.page_width = Inches(8.5); sec.page_height = Inches(11)
    sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Inches(1)

    top = doc.add_paragraph(style="BODY_TEXT")
    run = top.add_run(data.report_title); run.bold = True
    top.add_run(f" | {data.site_name} | {data.date}")
    doc.add_paragraph(data.report_title, style="TITLE_MAIN")
    doc.add_paragraph(data.environment_name, style="SUBTITLE")
    doc.add_paragraph(f"Assessment Date: {data.date}     |     Classification: {data.classification}", style="METADATA")
    line_para = doc.add_paragraph(""); _set_paragraph_bottom_border(line_para, COLORS["secondary_blue"], 12)

    h = doc.add_paragraph("Executive Summary", style="HEADING_SECTION"); _set_paragraph_bottom_border(h, COLORS["secondary_blue"], 8)
    doc.add_paragraph("The following table summarizes all findings identified during the CrowdStrike review.", style="BODY_TEXT")
    t = doc.add_table(rows=7, cols=3)
    headers = ["Finding Category", data.site_labels[0], data.site_labels[1]]
    for i, txt in enumerate(headers):
        c = t.cell(0, i); c.text = txt; _set_cell_fill(c, COLORS["secondary_blue"])
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.runs[0]; r.bold = True; r.font.color.rgb = RGBColor.from_string(COLORS["white"])
    for i in range(1, 7):
        t.cell(i, 0).text = f"Section {i}"; _set_cell_fill(t.cell(i, 0), COLORS["light_gray_fill"])
        t.cell(i, 1).text = "[ Describe ]"; t.cell(i, 2).text = "[ Describe ]"
    _set_table_borders(t, COLORS["light_border"], 6)

    doc.add_paragraph("5. Section 5", style="HEADING_SECTION")
    rec = doc.add_table(rows=max(2, len(data.recommendation_rows) + 1), cols=2)
    rec.cell(0, 0).text = "Recommendation"; rec.cell(0, 1).text = "Current Limitation / Response"
    for i, (left, right) in enumerate(data.recommendation_rows, start=1):
        rec.cell(i, 0).text = left; rec.cell(i, 1).text = right
        _set_cell_fill(rec.cell(i, 0), COLORS["warning_yellow"]); _set_cell_fill(rec.cell(i, 1), COLORS["error_red_pink"])
    _set_table_borders(rec, COLORS["light_border"], 6)

    h = doc.add_paragraph("Recommended Next Steps", style="HEADING_SECTION")
    for step in data.next_steps:
        doc.add_paragraph(step, style="List Number")
    doc.add_paragraph(f"Document prepared by the {data.prepared_by} for internal distribution only. All findings should be treated as confidential.", style="BODY_TEXT")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)
    return output_path


def build_excel_report(data: ReportData, output_path: Path) -> Path:
    wb = Workbook(); ws = wb.active; ws.title = "IT Security Report"
    ws.merge_cells("A1:C1"); ws["A1"] = data.report_title
    ws["A1"].font = Font(name="Arial", size=18, bold=True, color=COLORS["white"])
    ws["A1"].fill = PatternFill("solid", fgColor=COLORS["primary_dark_blue"])
    ws["A1"].alignment = Alignment(horizontal="center")
    ws["A2"] = f"Site: {data.site_name}"
    ws["B2"] = f"Environment: {data.environment_name}"
    ws["C2"] = f"Date: {data.date}"
    ws.append(["Recommendation", "Current Limitation / Response", "Status"])
    for c in ws[3]:
        c.font = Font(name="Arial", bold=True, color=COLORS["white"])
        c.fill = PatternFill("solid", fgColor=COLORS["secondary_blue"])
    for left, right in data.recommendation_rows:
        ws.append([left, right, data.status_value])
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 50
    ws.column_dimensions["C"].width = 24
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path


if __name__ == "__main__":
    kind, report, _ = collect_data()
    out_dir = Path("output")
    if kind == "excel":
        out = build_excel_report(report, out_dir / "it_security_report.xlsx")
    else:
        out = build_word_report(report, out_dir / "it_security_report_template.docx")
    print(f"Generated: {out}")
