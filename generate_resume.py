#!/usr/bin/env python3
"""Generate Love Taneja's resume in editable DOCX and print-ready PDF formats.

The content below is transcribed from the supplied CV and intentionally retains
its wording. Only typography, spacing, hierarchy, and page layout are changed.
"""

from __future__ import annotations

import html
import re
from pathlib import Path

import pymupdf
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor
from pypdf import PdfReader
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    HRFlowable,
    KeepTogether,
    PageBreak,
    PageTemplate,
    Paragraph,
)

ROOT = Path(__file__).resolve().parent
DOCX_PATH = ROOT / "Love_Taneja_Resume.docx"
PDF_PATH = ROOT / "Love_Taneja_Resume.pdf"
PREVIEW_DIR = ROOT / "preview"

BLUE = "2F4EA2"
ACCENT_BLUE = "1F5FBF"
DARK = "26364A"
MUTED = "5A6573"
RULE = "D8DEE8"
FONT_FAMILY = "Times New Roman"

PROFILE_SUMMARY = (
    "Qualified Chartered Accountant with 6+ years of progressive experience in Financial Planning and "
    "analysis, RTR, statutory audits, financial reporting, and compliance management. Specialized in "
    "Ind-AS/IFRS standards, audit execution, and financial analysis across manufacturing, finance, and "
    "hospitality sectors. Proven ability to collaborate with Big 4 firms (Deloitte & EY), manage complex "
    "audit engagements ensuring regulatory compliance under tight deadlines, and deliver accurate "
    "financial insights while ensuring adherence to regulatory frameworks."
)

CONTACT_ITEMS = [
    ("Email: ", "Lovetaneja100@gmail.com", "mailto:Lovetaneja100@gmail.com"),
    ("Mobile: ", "+91-9999027020", "tel:+919999027020"),
    (
        "LinkedIn: ",
        "http://www.linkedin.com/in/ca-love-taneja-6bba2a19b",
        "http://www.linkedin.com/in/ca-love-taneja-6bba2a19b",
    ),
]
ADDRESS = "Current Address: Shalimar Bagh, New Delhi, 110088"

EXPERIENCES = [
    {
        "heading": "OFB TECH LIMITED. (OF Business Group ) /ASSISTANT MANAGER / JAN 2026 – PRESENT",
        "subtitle": None,
        "bullets": [
            "Reconciliation of related party and non-related vendor, supplier ledger statement, and <b>Salary, ESI, PF, Advance of employee</b> with Plant & HR team.",
            "Reconciliation of Bank (BRS), petty-cash statement, and 26-AS portal vs books.",
            "Preparing daily vendors payment and Import/export payments.",
            "Calculation of monthly TDS, TDS return preparation.",
            "GST, TDS, PF, ESI, PT, LWF challan payment and reconciliation.",
            "GST R1 and 3B return prepared and reconciliation of GST 2B with books.",
            "Preparation of Interim/ Stat. audit data coordination with plant & Auditors (Big -4), Schedule and manage audit timelines to ensure timely completion and provide necessary documentation and support to auditors as required.",
            "Preparation of tax audit related data and necessary documents to auditor.",
            "Preparing monthly books closing and coordinating with plant for MIS and preparing MIS of monthly basis.",
            "Check with Compliance team BR, ROC, XBRL related work.",
            "Handling bank documentations related work- coordinate with bank person resolve queries.",
            "Preparation aging of debtors, creditors, MSME parties, advance payment and preparing provision for exp. and prepaid exp schedule.",
            "Support in financial reporting and analysis tasks as required.",
            "Investigate and resolve any discrepancies or variances in AR/AP accounts.",
            "Preparing annual budgets, rolling forecasts, and Variance analysis comparing actuals vs. budget/forecast, investigating deviations, and recommending corrective actions with management.",
            "Strategic Analysis supporting decision-making with profitability, cost, and revenue analysis.",
            "Implementation of new accounting software SME and SAP.",
        ],
    },
    {
        "heading": "IFCI Ltd. (Industrial Financial Corporation Limited)| Executive | September 2024 – Present | Delhi NCR, India",
        "subtitle": "(Finance, Executive working in Government Advisory)",
        "bullets": [
            "Managed claim settlements exceeding <b>1000 Cr.</b> ensuring compliance with regulatory guidelines of the PLI Auto scheme.",
            "Prepared and presented monthly <b>MIS reports Budget & forecasting</b> summarizing financial performance.",
            "Involved in <b>drafting guidelines</b> of Scheme to promote scheme of <b>Automobiles and Auto Component</b> Industry (PLI Auto) with <b>BSR & Co. LLP</b> while taking inputs from various stakeholders of automobile industry.",
            "Review <b>SOPs</b> & prepared for key financial areas, including <b>PP&E, Cash, Debt, AP/AR, Accruals, Payroll,</b> and <b>Operating Expenses.</b>",
            "Conducting <b>Financial Due Diligence</b>& Compliance, <b>Budget Calculation</b> for Applicants, Bank Reconciliation",
            "Conducted physical verification of <b>fixed assets and revenue records</b> across multiple manufacturing plants of major automobile players including Tata Motors (Chakan), Mahindra & Mahindra (Chakan), Delphi-TVS (Oragadam).",
            "Evaluated eligibility and verified claims under the Production Linked Incentive scheme of <b>Automobiles and Auto Component</b> Industry (PLI Auto) for major OEMs and Tier-1 suppliers.",
            "Conducted monthly <b>sales reconciliation</b> between BOA and external reports to ensure compliance.",
        ],
    },
    {
        "heading": "KDH Textile Private Limited | Finance Accounts Manager | Feb 2023 – Aug 2024 | New Delhi",
        "subtitle": None,
        "bullets": [
            "Manager Over <b>$ 50 Million</b> in ledger Reconciliation without discrepancies and <b>executed comprehensive audit,</b> assessing compliance with <b>Indian GAAP, IFRS-aligned standards, and delivering actionable insights</b> senior leadership.",
            "Performed <b>overall analytical reviews and substantive audit testing</b> for high-risk areas including <b>revenue, expenses,</b> and <b>receivables/payables,</b> ensuring audit accuracy and identifying material variances",
            "Conducted detailed <b>financial analysis,</b> including <b>ratio analysis, variance analysis,</b> and <b>input-output consumption reviews,</b> identifying operational inefficiencies.",
            "Preparation complete set of <b>Financial Statements</b> and related reports.",
            "Reviewed restructuring feasibility to support informed decision-making by stakeholders.",
        ],
    },
    {
        "heading": "Accenture Solution Pvt. Ltd. Finance process Enablement Analyst (R to R)| Finance Process Enablement Analyst | June 2022 – Jan 2023",
        "subtitle": None,
        "bullets": [
            "Preparing Bank Reconciliations , Preparing Monthly Accruals, Prepaid Amortization and their schedule , Fixed Asset Accounting and their schedule.",
            "Participation in Month End Activities , Pass GL to GL Entry , MIS Reporting. , Analysis of Ledger and financial Statement ,Managing and reviewing work of team and quality as per SLA for the process.",
            "Preparation and maintain all kinds of tracker required for the operation work, Passing of Journal Entries on month end , Handling Client calls and meetings Reconciliation of Books between IFRS and local GAAP.",
        ],
    },
    {
        "heading": "GKV & Co. Chartered Accountants | Audit Associates | March 2018 – May 2022 | Delhi NCR, India",
        "subtitle": None,
        "bullets": [
            "Finalized consolidated and standalone financial statements by accurately mapping and reconciling trial balances",
            "Led end-to-end statutory audit planning and execution across diverse industries, ensuring compliance with GAAP/Ind-AS and audit standards.",
            "Managed client communications and supervised audit teams for both listed and private entities.",
            "Performed <b>financial due diligence on tender files</b> and <b>analyzed debtors aging reports,</b> ensuring transparency, credit policy compliance, and timely corrective actions.",
            "Performing physical <b>asset verification</b> and <b>valuation reviews,</b> ensuring existence and accuracy of assets worth millions INR across multiple business units.",
        ],
    },
]

ARTICLE_HEADING = "Article Assistant and Industrial Trainee (Completed 3 Years)"
ARTICLE_BULLETS = [
    "Karan Jasuja & Associates (Chartered Accountants), Delhi &",
    "Industrial trainee from Indian oil Corporation of India, Noida (State office)",
]

KEY_RESPONSIBILITIES = [
    "Conducted statutory, internal, tax, and project audits across sectors including manufacturing, finance, hospitality, and non-profit organizations, ensuring compliance with applicable accounting and regulatory frameworks.",
    "Performed Test of Controls (TOC) and Test of Details (TOD) for key financial areas such as revenue, purchases, fixed assets, accounts payable, and accounts receivable, with comprehensive audit documentation.",
    "Led and participated in physical verification of fixed assets and inventory, ensuring accurate asset tagging, existence confirmation, and valuation in line with accounting standards.",
    "Executed cash verification during bank audits, ensuring internal controls over cash handling and reporting integrity.",
    "Assisted in government audit engagements, focusing on inventory management, compliance verification, and data reconciliation for public sector units.",
    "Developed working knowledge of audit sampling, client communication, and audit documentation aligned with ICAI Standards on Auditing (SAs).",
]

EDUCATION = [
    (
        "Chartered Accountant",
        "May 2024 (60.60%) | Institute of Chartered Accountants of India",
    ),
    ("Graduation", "B. Com (H) (51%) | 2014 | Delhi University"),
]

CERTIFICATIONS = [
    "Completed IT Training and Advanced ITT Training (Adv ITT).",
    "Completed both general management and communication skills (Adv-GMCS)",
]

TECHNICAL_SKILLS = [
    "Microsoft Excel | Microsoft Word | MS Power-Point",
    "Tally ERP | Oracle | SAP | Computax",
]

LANGUAGE = "English | Hindi"
CORE_COMPETENCIES = (
    "Statutory Audits | Internal Audits | Accounting | Banking | IFRS (Ind-AS) Financial Reporting | "
    "Financial Analysis | Risk Assessment | Accounts Payable | Policy & Regulatory Analysis | "
    "Public Finance Advisory"
)
HOBBIES = (
    "Global Travel & Cultural Exploration | Adventure Sports Enthusiast | Swimming free style | "
    "Social Impact Volunteering"
)
NATIONALITY = "INDIAN"


def strip_markup(text: str) -> str:
    return re.sub(r"</?[bi]>", "", text)


def reportlab_markup(text: str) -> str:
    """Escape text while retaining the small supported rich-text tag set."""
    parts = re.split(r"(</?[bi]>)", text)
    return "".join(part if re.fullmatch(r"</?[bi]>", part) else html.escape(part) for part in parts)


def set_repeat_font(run, name: str = FONT_FAMILY) -> None:
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:ascii"), name)
    run._element.rPr.rFonts.set(qn("w:hAnsi"), name)
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run._element.rPr.rFonts.set(qn("w:cs"), name)


def set_cell_margins(*_args, **_kwargs) -> None:
    # Reserved for future table-based layouts; current layout intentionally mirrors
    # the sample resume's clean single-column structure.
    return None


def set_keep_with_next(paragraph, value: bool = True) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    keep_next = p_pr.find(qn("w:keepNext"))
    if value and keep_next is None:
        p_pr.append(OxmlElement("w:keepNext"))
    elif not value and keep_next is not None:
        p_pr.remove(keep_next)


def set_widow_control(paragraph) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    if p_pr.find(qn("w:widowControl")) is None:
        p_pr.append(OxmlElement("w:widowControl"))


def set_bottom_border(paragraph, color: str = RULE, size: str = "8") -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = p_pr.find(qn("w:pBdr"))
    if p_bdr is None:
        p_bdr = OxmlElement("w:pBdr")
        p_pr.append(p_bdr)
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), size)
    bottom.set(qn("w:space"), "3")
    bottom.set(qn("w:color"), color)
    p_bdr.append(bottom)


def add_markup_runs(paragraph, text: str, *, color: str | None = None, size: float | None = None) -> None:
    bold = False
    italic = False
    for part in re.split(r"(</?[bi]>)", text):
        if part == "<b>":
            bold = True
        elif part == "</b>":
            bold = False
        elif part == "<i>":
            italic = True
        elif part == "</i>":
            italic = False
        elif part:
            run = paragraph.add_run(part)
            set_repeat_font(run)
            run.bold = bold
            run.italic = italic
            if color:
                run.font.color.rgb = RGBColor.from_string(color)
            if size:
                run.font.size = Pt(size)


def add_hyperlink(paragraph, text: str, url: str, *, color: str = DARK, underline: bool = False) -> None:
    relationship_id = paragraph.part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), relationship_id)
    new_run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    r_fonts = OxmlElement("w:rFonts")
    for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
        r_fonts.set(qn(f"w:{attr}"), FONT_FAMILY)
    r_pr.append(r_fonts)
    c = OxmlElement("w:color")
    c.set(qn("w:val"), color)
    r_pr.append(c)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "14")
    r_pr.append(sz)
    if underline:
        u = OxmlElement("w:u")
        u.set(qn("w:val"), "single")
        r_pr.append(u)
    new_run.append(r_pr)
    text_element = OxmlElement("w:t")
    text_element.text = text
    new_run.append(text_element)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def add_docx_section_heading(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    set_repeat_font(run)
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor.from_string(DARK)
    set_bottom_border(p)


def add_docx_bullet(doc: Document, markup: str) -> None:
    p = doc.add_paragraph()
    fmt = p.paragraph_format
    fmt.left_indent = Inches(0.13)
    fmt.first_line_indent = Inches(-0.09)
    fmt.right_indent = Inches(0)
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(1.2)
    fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    fmt.line_spacing = Pt(9.6)
    bullet = p.add_run("• ")
    set_repeat_font(bullet)
    bullet.font.size = Pt(8.35)
    bullet.font.color.rgb = RGBColor.from_string(DARK)
    add_markup_runs(p, markup, color=DARK, size=8.35)
    set_widow_control(p)


def add_docx_job(doc: Document, job: dict[str, object]) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4.5)
    p.paragraph_format.space_after = Pt(1.5)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(str(job["heading"]))
    set_repeat_font(run)
    run.bold = True
    run.font.size = Pt(9.05)
    run.font.color.rgb = RGBColor.from_string(ACCENT_BLUE)
    if job["subtitle"]:
        sub = doc.add_paragraph()
        sub.paragraph_format.space_after = Pt(1.5)
        sub.paragraph_format.keep_with_next = True
        r = sub.add_run(str(job["subtitle"]))
        set_repeat_font(r)
        r.bold = True
        r.italic = True
        r.font.size = Pt(7.9)
        r.font.color.rgb = RGBColor.from_string(ACCENT_BLUE)
    for bullet_text in job["bullets"]:  # type: ignore[index]
        add_docx_bullet(doc, str(bullet_text))


def configure_docx_styles(doc: Document) -> None:
    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = FONT_FAMILY
    normal._element.rPr.rFonts.set(qn("w:ascii"), FONT_FAMILY)
    normal._element.rPr.rFonts.set(qn("w:hAnsi"), FONT_FAMILY)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_FAMILY)
    normal.font.size = Pt(8.35)
    normal.font.color.rgb = RGBColor.from_string(DARK)
    normal.paragraph_format.space_after = Pt(1)
    normal.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    normal.paragraph_format.line_spacing = Pt(9.6)

    if "Resume Small" not in [style.name for style in styles]:
        small = styles.add_style("Resume Small", WD_STYLE_TYPE.PARAGRAPH)
    else:
        small = styles["Resume Small"]
    small.font.name = FONT_FAMILY
    small.font.size = Pt(7)
    small.font.color.rgb = RGBColor.from_string(DARK)
    small.paragraph_format.space_after = Pt(0)


def build_docx() -> None:
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Inches(0.39)
    section.bottom_margin = Inches(0.39)
    section.left_margin = Inches(0.45)
    section.right_margin = Inches(0.45)
    section.header_distance = Inches(0.15)
    section.footer_distance = Inches(0.15)

    configure_docx_styles(doc)
    props = doc.core_properties
    props.title = "Love Taneja — Chartered Accountant Resume"
    props.subject = "Professional resume"
    props.author = "Love Taneja"
    props.keywords = "Chartered Accountant, Finance, Audit, Resume"
    props.comments = "Reformatted from the supplied CV without rewriting its content."

    name = doc.add_paragraph()
    name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    name.paragraph_format.space_after = Pt(0.5)
    r = name.add_run("Love Taneja")
    set_repeat_font(r)
    r.bold = True
    r.font.size = Pt(17)
    r.font.color.rgb = RGBColor.from_string(BLUE)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_after = Pt(2.5)
    r = title.add_run("CHARTERED ACCOUNTANT")
    set_repeat_font(r)
    r.bold = True
    r.font.size = Pt(9.2)
    r.font.color.rgb = RGBColor.from_string(MUTED)

    contact = doc.add_paragraph()
    contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    contact.paragraph_format.space_after = Pt(0)
    for index, (label, value, url) in enumerate(CONTACT_ITEMS):
        if index:
            sep = contact.add_run("  |  ")
            set_repeat_font(sep)
            sep.font.size = Pt(7)
            sep.font.color.rgb = RGBColor.from_string(MUTED)
        label_run = contact.add_run(label)
        set_repeat_font(label_run)
        label_run.bold = True
        label_run.font.size = Pt(7)
        label_run.font.color.rgb = RGBColor.from_string(DARK)
        add_hyperlink(contact, value, url)

    address = doc.add_paragraph()
    address.alignment = WD_ALIGN_PARAGRAPH.CENTER
    address.paragraph_format.space_after = Pt(2)
    r = address.add_run(ADDRESS)
    set_repeat_font(r)
    r.font.size = Pt(7)
    r.font.color.rgb = RGBColor.from_string(DARK)

    add_docx_section_heading(doc, "PROFILE SUMMARY")
    summary = doc.add_paragraph()
    summary.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    summary.paragraph_format.space_after = Pt(1.5)
    summary.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    summary.paragraph_format.line_spacing = Pt(10)
    add_markup_runs(summary, PROFILE_SUMMARY, color=DARK, size=8.35)

    add_docx_section_heading(doc, "EXPERIENCES")
    for index, job in enumerate(EXPERIENCES):
        add_docx_job(doc, job)
        if index == 4:
            doc.add_page_break()

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1.5)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(ARTICLE_HEADING)
    set_repeat_font(r)
    r.bold = True
    r.font.size = Pt(9.05)
    r.font.color.rgb = RGBColor.from_string(ACCENT_BLUE)
    for item in ARTICLE_BULLETS:
        add_docx_bullet(doc, item)

    add_docx_section_heading(doc, "KEY RESPONSIBILITIES")
    sub = doc.add_paragraph()
    sub.paragraph_format.space_after = Pt(1.5)
    sub.paragraph_format.keep_with_next = True
    r = sub.add_run("FINANCE, ACCOUNTS & TAXATION")
    set_repeat_font(r)
    r.bold = True
    r.font.size = Pt(9.05)
    r.font.color.rgb = RGBColor.from_string(ACCENT_BLUE)
    for item in KEY_RESPONSIBILITIES:
        add_docx_bullet(doc, item)

    add_docx_section_heading(doc, "ACADEMIC QUALIFICATIONS")
    for qualification, detail in EDUCATION:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(qualification)
        set_repeat_font(r)
        r.bold = True
        r.font.size = Pt(8.8)
        r.font.color.rgb = RGBColor.from_string(ACCENT_BLUE)
        d = doc.add_paragraph()
        d.paragraph_format.space_after = Pt(2)
        add_markup_runs(d, detail, color=DARK, size=8.25)

    add_docx_section_heading(doc, "CERTIFICATION COURSES")
    for item in CERTIFICATIONS:
        add_docx_bullet(doc, item)

    add_docx_section_heading(doc, "TECHNICAL SKILLS")
    for item in TECHNICAL_SKILLS:
        add_docx_bullet(doc, item)

    add_docx_section_heading(doc, "LANGUAGE")
    p = doc.add_paragraph()
    add_markup_runs(p, LANGUAGE, color=DARK, size=8.35)

    add_docx_section_heading(doc, "CORE COMPETENCIES")
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing = Pt(10)
    add_markup_runs(p, CORE_COMPETENCIES, color=DARK, size=8.35)

    add_docx_section_heading(doc, "HOBBIES")
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing = Pt(10)
    add_markup_runs(p, HOBBIES, color=DARK, size=8.35)

    add_docx_section_heading(doc, "NATIONALITY")
    p = doc.add_paragraph()
    add_markup_runs(p, NATIONALITY, color=DARK, size=8.35)

    # Document compatibility and display settings.
    settings = doc.settings.element
    compat = settings.find(qn("w:compat"))
    if compat is None:
        compat = OxmlElement("w:compat")
        settings.append(compat)
    use_word_2013 = OxmlElement("w:compatSetting")
    use_word_2013.set(qn("w:name"), "compatibilityMode")
    use_word_2013.set(qn("w:uri"), "http://schemas.microsoft.com/office/word")
    use_word_2013.set(qn("w:val"), "15")
    compat.append(use_word_2013)

    doc.save(DOCX_PATH)


def register_pdf_fonts() -> None:
    font_dir = Path("/usr/share/fonts/liberation-serif")
    pdfmetrics.registerFont(TTFont("ResumeSerif", str(font_dir / "LiberationSerif-Regular.ttf")))
    pdfmetrics.registerFont(TTFont("ResumeSerif-Bold", str(font_dir / "LiberationSerif-Bold.ttf")))
    pdfmetrics.registerFont(TTFont("ResumeSerif-Italic", str(font_dir / "LiberationSerif-Italic.ttf")))
    pdfmetrics.registerFont(TTFont("ResumeSerif-BoldItalic", str(font_dir / "LiberationSerif-BoldItalic.ttf")))
    pdfmetrics.registerFontFamily(
        "ResumeSerif",
        normal="ResumeSerif",
        bold="ResumeSerif-Bold",
        italic="ResumeSerif-Italic",
        boldItalic="ResumeSerif-BoldItalic",
    )


def pdf_styles() -> dict[str, ParagraphStyle]:
    return {
        "name": ParagraphStyle(
            "Name",
            fontName="ResumeSerif-Bold",
            fontSize=17,
            leading=18,
            textColor=colors.HexColor(f"#{BLUE}"),
            alignment=TA_CENTER,
            spaceAfter=0.5,
        ),
        "title": ParagraphStyle(
            "Title",
            fontName="ResumeSerif-Bold",
            fontSize=9.2,
            leading=10,
            textColor=colors.HexColor(f"#{MUTED}"),
            alignment=TA_CENTER,
            spaceAfter=2.5,
        ),
        "contact": ParagraphStyle(
            "Contact",
            fontName="ResumeSerif",
            fontSize=7,
            leading=8.4,
            textColor=colors.HexColor(f"#{DARK}"),
            alignment=TA_CENTER,
            spaceAfter=0.5,
        ),
        "section": ParagraphStyle(
            "Section",
            fontName="ResumeSerif-Bold",
            fontSize=10,
            leading=11,
            textColor=colors.HexColor(f"#{DARK}"),
            alignment=TA_LEFT,
            spaceBefore=6.5,
            spaceAfter=1.5,
            keepWithNext=True,
        ),
        "summary": ParagraphStyle(
            "Summary",
            fontName="ResumeSerif",
            fontSize=8.35,
            leading=10,
            textColor=colors.HexColor(f"#{DARK}"),
            alignment=TA_JUSTIFY,
            spaceAfter=1.5,
        ),
        "job": ParagraphStyle(
            "Job",
            fontName="ResumeSerif-Bold",
            fontSize=9.05,
            leading=10.1,
            textColor=colors.HexColor(f"#{ACCENT_BLUE}"),
            alignment=TA_LEFT,
            spaceBefore=4.2,
            spaceAfter=1.4,
            keepWithNext=True,
        ),
        "subtitle": ParagraphStyle(
            "Subtitle",
            fontName="ResumeSerif-BoldItalic",
            fontSize=7.9,
            leading=8.8,
            textColor=colors.HexColor(f"#{ACCENT_BLUE}"),
            alignment=TA_LEFT,
            spaceAfter=1.2,
            keepWithNext=True,
        ),
        "bullet": ParagraphStyle(
            "Bullet",
            fontName="ResumeSerif",
            fontSize=8.35,
            leading=9.6,
            textColor=colors.HexColor(f"#{DARK}"),
            alignment=TA_JUSTIFY,
            leftIndent=8,
            rightIndent=0,
            firstLineIndent=0,
            bulletIndent=0,
            bulletFontName="ResumeSerif",
            bulletFontSize=7,
            spaceAfter=1.15,
            allowWidows=0,
            allowOrphans=0,
        ),
        "body": ParagraphStyle(
            "Body",
            fontName="ResumeSerif",
            fontSize=8.35,
            leading=9.8,
            textColor=colors.HexColor(f"#{DARK}"),
            alignment=TA_LEFT,
            spaceAfter=1.2,
        ),
        "education": ParagraphStyle(
            "Education",
            fontName="ResumeSerif-Bold",
            fontSize=8.8,
            leading=9.5,
            textColor=colors.HexColor(f"#{ACCENT_BLUE}"),
            alignment=TA_LEFT,
            spaceBefore=1,
            spaceAfter=0,
            keepWithNext=True,
        ),
    }


def section_flowables(text: str, styles: dict[str, ParagraphStyle]):
    return KeepTogether(
        [
            Paragraph(html.escape(text), styles["section"]),
            HRFlowable(
                width="100%",
                thickness=0.55,
                color=colors.HexColor(f"#{RULE}"),
                spaceBefore=0,
                spaceAfter=2.5,
                hAlign="LEFT",
            ),
        ]
    )


def add_pdf_bullet(story: list, text: str, styles: dict[str, ParagraphStyle]) -> None:
    story.append(Paragraph(reportlab_markup(text), styles["bullet"], bulletText="•"))


def on_pdf_page(canvas, doc) -> None:
    canvas.saveState()
    canvas.setTitle("Love Taneja — Chartered Accountant Resume")
    canvas.setAuthor("Love Taneja")
    canvas.setSubject("Professional resume")
    canvas.setKeywords("Chartered Accountant, Finance, Audit, Resume")
    canvas.restoreState()


def build_pdf() -> None:
    register_pdf_fonts()
    styles = pdf_styles()
    doc = BaseDocTemplate(
        str(PDF_PATH),
        pagesize=A4,
        leftMargin=11.5 * mm,
        rightMargin=11.5 * mm,
        topMargin=10 * mm,
        bottomMargin=10 * mm,
        title="Love Taneja — Chartered Accountant Resume",
        author="Love Taneja",
        subject="Professional resume",
    )
    frame = Frame(
        doc.leftMargin,
        doc.bottomMargin,
        doc.width,
        doc.height,
        leftPadding=0,
        rightPadding=0,
        topPadding=0,
        bottomPadding=0,
        id="resume-frame",
    )
    doc.addPageTemplates([PageTemplate(id="resume", frames=[frame], onPage=on_pdf_page)])

    story: list = []
    story.append(Paragraph("Love Taneja", styles["name"]))
    story.append(Paragraph("CHARTERED ACCOUNTANT", styles["title"]))

    contact_bits = []
    for label, value, url in CONTACT_ITEMS:
        contact_bits.append(f"<b>{html.escape(label)}</b><link href=\"{html.escape(url)}\">{html.escape(value)}</link>")
    story.append(Paragraph(" &nbsp;|&nbsp; ".join(contact_bits), styles["contact"]))
    story.append(Paragraph(html.escape(ADDRESS), styles["contact"]))

    story.append(section_flowables("PROFILE SUMMARY", styles))
    story.append(Paragraph(html.escape(PROFILE_SUMMARY), styles["summary"]))

    story.append(section_flowables("EXPERIENCES", styles))
    for index, job in enumerate(EXPERIENCES):
        story.append(Paragraph(html.escape(str(job["heading"])), styles["job"]))
        if job["subtitle"]:
            story.append(Paragraph(html.escape(str(job["subtitle"])), styles["subtitle"]))
        for bullet_text in job["bullets"]:  # type: ignore[index]
            add_pdf_bullet(story, str(bullet_text), styles)
        if index == 4:
            story.append(PageBreak())

    story.append(Paragraph(html.escape(ARTICLE_HEADING), styles["job"]))
    for item in ARTICLE_BULLETS:
        add_pdf_bullet(story, item, styles)

    story.append(section_flowables("KEY RESPONSIBILITIES", styles))
    story.append(Paragraph("FINANCE, ACCOUNTS &amp; TAXATION", styles["job"]))
    for item in KEY_RESPONSIBILITIES:
        add_pdf_bullet(story, item, styles)

    story.append(section_flowables("ACADEMIC QUALIFICATIONS", styles))
    for qualification, detail in EDUCATION:
        story.append(Paragraph(html.escape(qualification), styles["education"]))
        story.append(Paragraph(html.escape(detail), styles["body"]))

    story.append(section_flowables("CERTIFICATION COURSES", styles))
    for item in CERTIFICATIONS:
        add_pdf_bullet(story, item, styles)

    story.append(section_flowables("TECHNICAL SKILLS", styles))
    for item in TECHNICAL_SKILLS:
        add_pdf_bullet(story, item, styles)

    story.append(section_flowables("LANGUAGE", styles))
    story.append(Paragraph(html.escape(LANGUAGE), styles["body"]))

    story.append(section_flowables("CORE COMPETENCIES", styles))
    story.append(Paragraph(html.escape(CORE_COMPETENCIES), styles["body"]))

    story.append(section_flowables("HOBBIES", styles))
    story.append(Paragraph(html.escape(HOBBIES), styles["body"]))

    story.append(section_flowables("NATIONALITY", styles))
    story.append(Paragraph(html.escape(NATIONALITY), styles["body"]))

    doc.build(story)


def expected_text_blocks() -> list[str]:
    blocks = [
        "Love Taneja",
        "CHARTERED ACCOUNTANT",
        *[f"{label}{value}" for label, value, _ in CONTACT_ITEMS],
        ADDRESS,
        "PROFILE SUMMARY",
        PROFILE_SUMMARY,
        "EXPERIENCES",
    ]
    for job in EXPERIENCES:
        blocks.append(str(job["heading"]))
        if job["subtitle"]:
            blocks.append(str(job["subtitle"]))
        blocks.extend(strip_markup(str(item)) for item in job["bullets"])  # type: ignore[index]
    blocks.extend([ARTICLE_HEADING, *ARTICLE_BULLETS, "KEY RESPONSIBILITIES", "FINANCE, ACCOUNTS & TAXATION"])
    blocks.extend(KEY_RESPONSIBILITIES)
    blocks.append("ACADEMIC QUALIFICATIONS")
    for qualification, detail in EDUCATION:
        blocks.extend([qualification, detail])
    blocks.extend(["CERTIFICATION COURSES", *CERTIFICATIONS])
    blocks.extend(["TECHNICAL SKILLS", *TECHNICAL_SKILLS])
    blocks.extend(["LANGUAGE", LANGUAGE])
    blocks.extend(["CORE COMPETENCIES", CORE_COMPETENCIES])
    blocks.extend(["HOBBIES", HOBBIES])
    blocks.extend(["NATIONALITY", NATIONALITY])
    return blocks


def normalize_text(text: str) -> str:
    text = text.replace("\u00a0", " ").replace("–", "-").replace("—", "-")
    return re.sub(r"\s+", " ", text).strip()


def validate_outputs() -> tuple[int, int]:
    # Validate DOCX text completeness by reopening the generated file.
    doc = Document(DOCX_PATH)
    docx_text = normalize_text("\n".join(paragraph.text for paragraph in doc.paragraphs))

    # Validate PDF text completeness and count pages.
    reader = PdfReader(str(PDF_PATH))
    pdf_text = normalize_text("\n".join(page.extract_text() or "" for page in reader.pages))

    missing_docx: list[str] = []
    missing_pdf: list[str] = []
    for block in expected_text_blocks():
        normalized = normalize_text(strip_markup(block))
        if normalized not in docx_text:
            missing_docx.append(normalized)
        if normalized not in pdf_text:
            missing_pdf.append(normalized)

    if missing_docx or missing_pdf:
        details = []
        if missing_docx:
            details.append("DOCX missing: " + " | ".join(missing_docx[:5]))
        if missing_pdf:
            details.append("PDF missing: " + " | ".join(missing_pdf[:5]))
        raise RuntimeError("Content validation failed. " + "; ".join(details))

    return len(doc.paragraphs), len(reader.pages)


def render_pdf_previews() -> list[Path]:
    PREVIEW_DIR.mkdir(exist_ok=True)
    for old_preview in PREVIEW_DIR.glob("page-*.png"):
        old_preview.unlink()
    pdf = pymupdf.open(PDF_PATH)
    paths: list[Path] = []
    matrix = pymupdf.Matrix(1.5, 1.5)
    for index, page in enumerate(pdf):
        path = PREVIEW_DIR / f"page-{index + 1}.png"
        page.get_pixmap(matrix=matrix, alpha=False).save(path)
        paths.append(path)
    return paths


def main() -> None:
    build_docx()
    build_pdf()
    paragraph_count, page_count = validate_outputs()
    previews = render_pdf_previews()
    print(f"Generated {DOCX_PATH.name} ({paragraph_count} paragraphs)")
    print(f"Generated {PDF_PATH.name} ({page_count} pages)")
    print(f"Rendered {len(previews)} preview image(s)")


if __name__ == "__main__":
    main()
