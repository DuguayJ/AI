#!/usr/bin/env python3
"""Generate a strict IT Security report Word template (.docx)."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date as dt_date
from pathlib import Path
from typing import List, Tuple

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

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
    site_name: str = "[ Site Name ]"
    environment_name: str = "[ Site / Environment Name ]"
    date: str = field(default_factory=lambda: dt_date.today().isoformat())
    classification: str = "Confidential — Internal Use Only"
    site_labels: Tuple[str, str] = ("NDH", "HPH")
    section_content: dict = field(default_factory=lambda: {
        "1_main": "[ Describe ]",
        "1_ndh": "[ Describe ]",
        "1_hph": "[ Describe ]",
        "2_ndh": "[ Describe ]",
        "2_hph": "[ Describe ]",
        "3_ndh": "[ Describe ]",
        "3_hph": "[ Describe ]",
        "4_main": "[ Describe ]",
        "4_ndh": "[ Describe ]",
        "4_hph": "[ Describe ]",
        "6_main": "[ Describe ]",
    })
    recommendation_rows: List[Tuple[str, str]] = field(default_factory=lambda: [
        ("[ Recommendation ]", "[ Current Limitation / Response ]"),
        ("[ Recommendation ]", "[ Current Limitation / Response ]"),
    ])
    status_value: str = "[ In Progress / Completed / Pending Review ]"
    next_steps: List[str] = field(default_factory=lambda: [
        "[ Next step 1 ]",
        "[ Next step 2 ]",
        "[ Next step 3 ]",
        "[ Schedule follow-up review ]",
    ])
    prepared_by: str = "IT Security team"


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

    s = add_paragraph_style("TITLE_MAIN")
    s.font.name = "Arial"; s._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    s.font.size = Pt(26); s.font.bold = True; s.font.color.rgb = RGBColor.from_string(COLORS["primary_dark_blue"])
    s.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    s.paragraph_format.line_spacing = 1.0

    s = add_paragraph_style("SUBTITLE")
    s.font.name = "Arial"; s._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    s.font.size = Pt(16); s.font.italic = True; s.font.color.rgb = RGBColor.from_string(COLORS["secondary_blue"])
    s.paragraph_format.line_spacing = 1.0

    s = add_paragraph_style("METADATA")
    s.font.name = "Arial"; s._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    s.font.size = Pt(9); s.font.color.rgb = RGBColor.from_string(COLORS["placeholder_gray"])

    s = add_paragraph_style("HEADING_SECTION")
    s.font.name = "Arial"; s._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    s.font.size = Pt(13); s.font.bold = True; s.font.color.rgb = RGBColor.from_string(COLORS["primary_dark_blue"])
    s.paragraph_format.space_before = Pt(16); s.paragraph_format.space_after = Pt(6); s.paragraph_format.line_spacing = 1.0

    s = add_paragraph_style("BODY_TEXT")
    s.font.name = "Arial"; s._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    s.font.size = Pt(10.5); s.font.color.rgb = RGBColor.from_string(COLORS["body_text"])
    s.paragraph_format.line_spacing = 1.15
    s.paragraph_format.space_after = Pt(12)

    s = add_paragraph_style("PLACEHOLDER_TEXT")
    s.font.name = "Arial"; s._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    s.font.size = Pt(10.5); s.font.italic = True; s.font.color.rgb = RGBColor.from_string(COLORS["placeholder_gray"])
    s.paragraph_format.line_spacing = 1.15

    s = add_paragraph_style("SITE_LABEL")
    s.font.name = "Arial"; s._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    s.font.size = Pt(11); s.font.bold = True; s.font.color.rgb = RGBColor.from_string(COLORS["primary_dark_blue"])

    s = add_paragraph_style("FOOTER_TEXT")
    s.font.name = "Arial"; s._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    s.font.size = Pt(8); s.font.color.rgb = RGBColor.from_string(COLORS["placeholder_gray"])


def build_security_report_template(data: ReportData, output_path: Path) -> Path:
    doc = Document()
    _add_styles(doc)

    section = doc.sections[0]
    section.page_width = Inches(8.5); section.page_height = Inches(11)
    section.top_margin = Inches(1); section.bottom_margin = Inches(1)
    section.left_margin = Inches(1); section.right_margin = Inches(1)

    footer_table = section.footer.add_table(rows=1, cols=2, width=Inches(6.5))
    footer_table.alignment = WD_TABLE_ALIGNMENT.LEFT
    footer_table.cell(0, 0).text = "Confidential — Internal Use Only"
    p = footer_table.cell(0, 0).paragraphs[0]; p.style = "FOOTER_TEXT"; p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p2 = footer_table.cell(0, 1).paragraphs[0]; p2.style = "FOOTER_TEXT"; p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p2.add_run("Page ")
    fld = OxmlElement('w:fldSimple'); fld.set(qn('w:instr'), 'PAGE'); p2._p.append(fld)

    top = doc.add_paragraph(style="BODY_TEXT")
    r1 = top.add_run(data.report_title); r1.bold = True
    top.add_run(f" | {data.site_name} | {data.date}")

    doc.add_paragraph(data.report_title, style="TITLE_MAIN")
    doc.add_paragraph(data.environment_name, style="SUBTITLE")
    doc.add_paragraph(f"Assessment Date: {data.date}     |     Classification: {data.classification}", style="METADATA")
    line_para = doc.add_paragraph("")
    _set_paragraph_bottom_border(line_para, COLORS["secondary_blue"], 12)

    h = doc.add_paragraph("Executive Summary", style="HEADING_SECTION")
    _set_paragraph_bottom_border(h, COLORS["secondary_blue"], 8)
    doc.add_paragraph("The following table summarizes all findings identified during the CrowdStrike review.", style="BODY_TEXT")

    t = doc.add_table(rows=7, cols=3)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = True
    headers = ["Finding Category", data.site_labels[0], data.site_labels[1]]
    for i, text in enumerate(headers):
        c = t.cell(0, i); c.text = text; _set_cell_fill(c, COLORS["secondary_blue"])
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.runs[0]; run.bold = True; run.font.color.rgb = RGBColor.from_string(COLORS["white"]); run.font.name = "Arial"
    for idx in range(1, 7):
        t.cell(idx, 0).text = f"Section {idx}"; _set_cell_fill(t.cell(idx, 0), COLORS["light_gray_fill"])
        for col in (1, 2):
            cell = t.cell(idx, col); cell.text = "[ Describe ]"; _set_cell_fill(cell, COLORS["white"])
            cell.paragraphs[0].style = "PLACEHOLDER_TEXT"
    _set_table_borders(t, COLORS["light_border"], 6)

    def add_label_and_placeholder(label: str, value: str):
        doc.add_paragraph(label, style="SITE_LABEL")
        doc.add_paragraph(value, style="PLACEHOLDER_TEXT")

    h = doc.add_paragraph("1. Section 1", style="HEADING_SECTION"); _set_paragraph_bottom_border(h, COLORS["secondary_blue"], 8)
    doc.add_paragraph(data.section_content["1_main"], style="PLACEHOLDER_TEXT")
    add_label_and_placeholder(data.site_labels[0], data.section_content["1_ndh"])
    add_label_and_placeholder(data.site_labels[1], data.section_content["1_hph"])

    for n in (2, 3):
        h = doc.add_paragraph(f"{n}.", style="HEADING_SECTION"); _set_paragraph_bottom_border(h, COLORS["secondary_blue"], 8)
        add_label_and_placeholder(data.site_labels[0], data.section_content[f"{n}_ndh"])
        add_label_and_placeholder(data.site_labels[1], data.section_content[f"{n}_hph"])

    h = doc.add_paragraph("4.", style="HEADING_SECTION"); _set_paragraph_bottom_border(h, COLORS["secondary_blue"], 8)
    doc.add_paragraph(data.section_content["4_main"], style="PLACEHOLDER_TEXT")
    add_label_and_placeholder(data.site_labels[0], data.section_content["4_ndh"])
    add_label_and_placeholder(data.site_labels[1], data.section_content["4_hph"])

    h = doc.add_paragraph("5. Section 5", style="HEADING_SECTION"); _set_paragraph_bottom_border(h, COLORS["secondary_blue"], 8)
    doc.add_paragraph("[List accounts reviewed, confirm whether each is legitimate, and note any CrowdStrike recommendations and their feasibility.]", style="PLACEHOLDER_TEXT")
    rec = doc.add_table(rows=max(2, len(data.recommendation_rows) + 1), cols=2)
    rec.cell(0, 0).text = "Recommendation"; rec.cell(0, 1).text = "Current Limitation / Response"
    for c in rec.rows[0].cells:
        _set_cell_fill(c, COLORS["secondary_blue"])
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.runs[0]; run.bold = True; run.font.color.rgb = RGBColor.from_string(COLORS["white"])
    for i, (left, right) in enumerate(data.recommendation_rows, start=1):
        rec.cell(i, 0).text = left; rec.cell(i, 1).text = right
        _set_cell_fill(rec.cell(i, 0), COLORS["warning_yellow"])
        _set_cell_fill(rec.cell(i, 1), COLORS["error_red_pink"])
        rec.cell(i, 0).paragraphs[0].style = "PLACEHOLDER_TEXT"
        rec.cell(i, 1).paragraphs[0].style = "PLACEHOLDER_TEXT"
    _set_table_borders(rec, COLORS["light_border"], 6)

    h = doc.add_paragraph("6.", style="HEADING_SECTION"); _set_paragraph_bottom_border(h, COLORS["secondary_blue"], 8)
    doc.add_paragraph(data.section_content["6_main"], style="PLACEHOLDER_TEXT")
    status = doc.add_table(rows=1, cols=1)
    _set_cell_fill(status.cell(0, 0), COLORS["accent_light_blue"])
    p = status.cell(0, 0).paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("Status: "); run.bold = True
    p.add_run(data.status_value)

    h = doc.add_paragraph("Recommended Next Steps", style="HEADING_SECTION"); _set_paragraph_bottom_border(h, COLORS["secondary_blue"], 8)
    for step in data.next_steps:
        doc.add_paragraph(step, style="List Number")

    doc.add_paragraph(
        f"Document prepared by the {data.prepared_by} for internal distribution only. All findings should be treated as confidential.",
        style="BODY_TEXT",
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)
    return output_path


if __name__ == "__main__":
    out = build_security_report_template(ReportData(), Path("output/it_security_report_template.docx"))
    print(f"Generated: {out}")
