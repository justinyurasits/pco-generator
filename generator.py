#!/usr/bin/env python3
"""
Change Order Generator — v2
Grounded in CHG-51 (Change Event SOP) and the firm pricing database.

Key changes from v1:
- Document is now correctly typed as a PCO (Potential Change Order)
- Scope narrative follows CHG-51 structure: baseline → change → inclusions/exclusions → assumptions
- Markup applied per pricing database: 1% app fee + 10% overhead + 10% fee = 21%
- Validity period added (PCOs require expiration)
- Cause categories aligned to CHG-51 taxonomy
- Authorization language aligned to SOP non-negotiables
"""

import os
from datetime import datetime, timedelta
from pathlib import Path

import anthropic
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable, Table, TableStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib import colors

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ---------------------------------------------------------------------------
# Pricing database constants (from SYS-09 Estimating Schedule)
# ---------------------------------------------------------------------------

MARKUP = {
    "application_fee_pct": 0.01,   # 1%
    "overhead_pct":        0.10,   # 10%
    "fee_pct":             0.10,   # 10%
    "total_pct":           0.21,   # 21% combined
}

# Key labor rates for sanity-checking field inputs
LABOR_BENCHMARKS = {
    "demo_crew_day":         1920,
    "carpentry_crew_day":    2486,
    "general_labor_hr_low":  75,
    "general_labor_hr_high": 85,
    "skilled_trade_hr":      89,
}

# Common material benchmarks for change order work
MATERIAL_BENCHMARKS = {
    "subfloor_3_4_material_sf_low":  3.26,
    "subfloor_3_4_material_sf_high": 3.76,
    "subfloor_install_sf_low":       5.25,
    "subfloor_install_sf_high":      6.25,
    "joist_hangers_allowance":       765,
    "trash_removal_per_load":        1301,
    "site_protection_ls":            1000,
}

# CHG-51 cause categories
CAUSE_CATEGORIES = [
    "Owner-Requested Change",
    "Design Revision / ASI",
    "Unforeseen Field Condition",
    "Allowance Over/Under",
    "Coordination Conflict",
    "Unit Price Work",
    "Other",
]

PCO_VALIDITY_DAYS = 21  # Standard validity period for PCOs


# ---------------------------------------------------------------------------
# Markup calculator
# ---------------------------------------------------------------------------

def apply_markup(labor: float, materials: float) -> dict:
    """Apply the firm's standard markup structure to construction costs."""
    subtotal = labor + materials
    app_fee  = round(subtotal * MARKUP["application_fee_pct"], 2)
    overhead = round(subtotal * MARKUP["overhead_pct"], 2)
    fee      = round(subtotal * MARKUP["fee_pct"], 2)
    total    = round(subtotal + app_fee + overhead + fee, 2)
    return {
        "labor":    labor,
        "materials": materials,
        "subtotal":  subtotal,
        "app_fee":   app_fee,
        "overhead":  overhead,
        "fee":       fee,
        "total":     total,
    }


def parse_cost(value: str) -> float:
    """Parse a cost string like '2,400' or '$2400' to a float."""
    if not value:
        return 0.0
    cleaned = value.replace("$", "").replace(",", "").strip()
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


# ---------------------------------------------------------------------------
# System prompt — grounded in CHG-51 and pricing database
# ---------------------------------------------------------------------------

SYSTEM_PROMPT = f"""You are a construction contract specialist working within a firm that follows CHG-51 (Change Event Standard Operating Procedure). Your job is to convert field notes about a scope change into a professional Potential Change Order (PCO) that a project manager can review and send to the owner for approval.

DOCUMENT TYPE:
This is a PCO — a Potential Change Order. It is a proposal for owner approval. Work described herein SHALL NOT proceed until the owner has signed and returned this document. This is not an executed Change Order (PCCO/AIA G701).

CAUSE CATEGORIES (use exactly one):
{chr(10).join(f'- {c}' for c in CAUSE_CATEGORIES)}

SCOPE NARRATIVE STRUCTURE (CHG-51 requirement — produce all four elements):
1. BASELINE: Describe what the original contract included relevant to this change (what was assumed or not included).
2. CHANGE: Describe precisely what is different — what work is being added, modified, or deleted.
3. INCLUSIONS/EXCLUSIONS: State explicitly what is included in this PCO and what is excluded (owner-supplied items, adjacent trades, permits if applicable).
4. ASSUMPTIONS: State any access requirements, sequencing constraints, protection requirements, or conditions that must be true for the price to hold.

COST STRUCTURE:
The cost section will be provided to you as pre-calculated values. Present them exactly as given. Do not recalculate or modify the numbers.

OUTPUT FORMAT — produce these exact section headers, plain text only, in this exact order:

SCOPE OF ADDITIONAL WORK
[Baseline, then Change, in professional contract language]

INCLUSIONS AND EXCLUSIONS
[What is included / what is excluded from this PCO]

ASSUMPTIONS AND CONDITIONS
[Access, sequencing, protection, or pricing conditions]

COST SUMMARY
[Will be inserted from the pricing calculation — output the placeholder text: COST_BLOCK_PLACEHOLDER]

SCHEDULE IMPACT
[Calendar days added and note that revised completion date is subject to schedule review]

CAUSE OF CHANGE
[One cause category from the list above, then one sentence explanation]

CONTRACT MODIFICATION
This Potential Change Order, when executed by both parties, will modify the original contract dated [original contract date] between [company name] and the Owner. All other terms and conditions of the original contract remain in full force and effect.

VALIDITY
This proposal is valid for {PCO_VALIDITY_DAYS} days from the date of issue. If not executed within this period, pricing is subject to revision.

AUTHORIZATION
No work described in this Potential Change Order shall commence until this document has been executed (signed by both parties). Verbal or written direction to proceed without a signed PCO constitutes authorization for T&M billing per the terms of the original contract.

RULES:
- Never use informal language
- Never fabricate cost numbers — the cost block will be provided separately
- Never add scope not described in the input
- Write the scope as GC scope — not as a vendor quote (you control scope language and boundaries)
- Output plain text only — no markdown, no bullet points, no # headers
- Write COST_BLOCK_PLACEHOLDER exactly as shown — the system will replace it
- This is a DRAFT for PM review"""


# ---------------------------------------------------------------------------
# Claude call
# ---------------------------------------------------------------------------

def generate_pco_text(data: dict) -> str:
    """Call Claude to generate PCO text from structured field data."""

    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

    user_message = f"""Please generate a PCO draft from the following field data:

Company Name: {data.get('company_name', '[TBD]')}
Project Name: {data.get('project_name', '[TBD]')}
Project Address: {data.get('project_address', '[TBD]')}
PCO Number: {data.get('change_order_number', '001')}
Date: {data.get('date', datetime.now().strftime('%B %d, %Y'))}
Original Contract Date: {data.get('original_contract_date', '[TBD]')}

Field Description of Scope Change:
{data.get('scope_description', '[No description provided]')}

Schedule Impact: {data.get('schedule_days', '[TBD]')} calendar days
Reason/Cause: {data.get('reason_for_change', '[TBD]')}

Note: The cost section will be inserted by the system — output COST_BLOCK_PLACEHOLDER exactly where the cost summary should appear."""

    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1800,
        temperature=0,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_message}]
    )

    return message.content[0].text


# ---------------------------------------------------------------------------
# Section parser
# ---------------------------------------------------------------------------

SECTION_HEADERS = [
    "SCOPE OF ADDITIONAL WORK",
    "INCLUSIONS AND EXCLUSIONS",
    "ASSUMPTIONS AND CONDITIONS",
    "COST SUMMARY",
    "SCHEDULE IMPACT",
    "CAUSE OF CHANGE",
    "CONTRACT MODIFICATION",
    "VALIDITY",
    "AUTHORIZATION",
]

def parse_sections(text: str, cost_block: str) -> dict:
    """Parse Claude's output into sections, replacing the cost placeholder."""

    # Replace the cost placeholder before parsing
    text = text.replace("COST_BLOCK_PLACEHOLDER", cost_block)

    sections = {}
    current_header = None
    current_lines = []

    for line in text.strip().split('\n'):
        stripped = line.strip().upper()
        matched = next((h for h in SECTION_HEADERS if stripped == h), None)

        if matched:
            if current_header and current_lines:
                sections[current_header] = '\n'.join(current_lines).strip()
            current_header = matched
            current_lines = []
        elif current_header:
            current_lines.append(line)

    if current_header and current_lines:
        sections[current_header] = '\n'.join(current_lines).strip()

    return sections


def format_cost_block(costs: dict) -> str:
    """Format the markup-applied cost breakdown as plain text."""
    if costs["subtotal"] == 0:
        return "Labor: [TBD]\nMaterials: [TBD]\nConstruction Subtotal: [TBD]\nApplication Fee (1%): [TBD]\nOverhead (10%): [TBD]\nFee (10%): [TBD]\nTotal PCO Amount: [TBD]"

    return (
        f"Labor:                    ${costs['labor']:>10,.2f}\n"
        f"Materials:                ${costs['materials']:>10,.2f}\n"
        f"Construction Subtotal:    ${costs['subtotal']:>10,.2f}\n"
        f"Application Fee (1%):     ${costs['app_fee']:>10,.2f}\n"
        f"Overhead (10%):           ${costs['overhead']:>10,.2f}\n"
        f"Fee (10%):                ${costs['fee']:>10,.2f}\n"
        f"Total PCO Amount:         ${costs['total']:>10,.2f}"
    )


# ---------------------------------------------------------------------------
# PDF builder
# ---------------------------------------------------------------------------

def build_pdf(data: dict, sections: dict, costs: dict, output_path: str):
    """Produce a professional PDF PCO document."""

    doc = SimpleDocTemplate(
        output_path,
        pagesize=letter,
        rightMargin=1 * inch,
        leftMargin=1 * inch,
        topMargin=1 * inch,
        bottomMargin=1 * inch
    )

    company_style = ParagraphStyle(
        'Company', fontSize=14, fontName='Helvetica-Bold',
        alignment=TA_LEFT, spaceAfter=4
    )
    title_style = ParagraphStyle(
        'Title', fontSize=18, fontName='Helvetica-Bold',
        alignment=TA_CENTER, spaceBefore=8, spaceAfter=2
    )
    subtitle_style = ParagraphStyle(
        'Subtitle', fontSize=11, fontName='Helvetica',
        alignment=TA_CENTER, spaceAfter=4, textColor=colors.HexColor('#555555')
    )
    draft_style = ParagraphStyle(
        'Draft', fontSize=9, fontName='Helvetica-Bold',
        alignment=TA_CENTER, textColor=colors.red, spaceAfter=12
    )
    section_header_style = ParagraphStyle(
        'SectionHeader', fontSize=10, fontName='Helvetica-Bold',
        spaceBefore=14, spaceAfter=4
    )
    body_style = ParagraphStyle(
        'Body', fontSize=10, fontName='Helvetica',
        spaceAfter=6, leading=15
    )
    cost_style = ParagraphStyle(
        'Cost', fontSize=10, fontName='Courier',
        spaceAfter=6, leading=15
    )

    story = []

    # Letterhead
    story.append(Paragraph(data.get('company_name', ''), company_style))
    story.append(HRFlowable(width="100%", thickness=2, color=colors.black, spaceAfter=4))
    story.append(Paragraph("POTENTIAL CHANGE ORDER", title_style))
    story.append(Paragraph("For Owner Review and Approval", subtitle_style))
    story.append(Paragraph("DRAFT — FOR PM REVIEW ONLY", draft_style))

    # Header block
    co_date = data.get('date', datetime.now().strftime('%B %d, %Y'))
    validity_date = (datetime.now() + timedelta(days=PCO_VALIDITY_DAYS)).strftime('%B %d, %Y')

    header_rows = [
        [f"PCO No.:  {data.get('change_order_number', '001')}",
         f"Date:  {co_date}"],
        [f"Project:  {data.get('project_name', '')}",
         f"Original Contract Date:  {data.get('original_contract_date', '[TBD]')}"],
        [f"Address:  {data.get('project_address', '')}",
         f"Valid Through:  {validity_date}"],
    ]
    header_table = Table(header_rows, colWidths=[3.75 * inch, 2.75 * inch])
    header_table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 8))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.grey, spaceAfter=4))

    # Generated sections — use monospace for cost block
    for header in SECTION_HEADERS:
        if header in sections:
            story.append(Paragraph(header, section_header_style))
            style = cost_style if header == "COST SUMMARY" else body_style
            content = sections[header].replace('\n', '<br/>')
            story.append(Paragraph(content, style))

    story.append(Spacer(1, 24))

    # Signature blocks
    sig_rows = [
        ["Contractor:", "_" * 38, "Date:", "_" * 14],
        ["", "", "", ""],
        ["Print Name:", "_" * 35, "", ""],
        ["", "", "", ""],
        ["Owner:", "_" * 41, "Date:", "_" * 14],
        ["", "", "", ""],
        ["Print Name:", "_" * 35, "", ""],
    ]
    sig_table = Table(sig_rows, colWidths=[1.1 * inch, 2.3 * inch, 0.6 * inch, 2.5 * inch])
    sig_table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]))
    story.append(sig_table)

    doc.build(story)


# ---------------------------------------------------------------------------
# Word doc builder
# ---------------------------------------------------------------------------

def _add_rule(doc):
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '000000')
    pBdr.append(bottom)
    pPr.append(pBdr)


def build_word(data: dict, sections: dict, costs: dict, output_path: str):
    """Produce an editable Word doc PCO for PM review."""

    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Company
    p = doc.add_paragraph()
    run = p.add_run(data.get('company_name', ''))
    run.bold = True
    run.font.size = Pt(14)

    _add_rule(doc)

    # Title
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("POTENTIAL CHANGE ORDER")
    run.bold = True
    run.font.size = Pt(18)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("For Owner Review and Approval")
    run.font.size = Pt(11)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("DRAFT — FOR PM REVIEW ONLY")
    run.bold = True
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)

    doc.add_paragraph()

    co_date = data.get('date', datetime.now().strftime('%B %d, %Y'))
    validity_date = (datetime.now() + timedelta(days=PCO_VALIDITY_DAYS)).strftime('%B %d, %Y')

    header_lines = [
        f"PCO No.:  {data.get('change_order_number', '001')}",
        f"Date:  {co_date}",
        f"Project:  {data.get('project_name', '')}",
        f"Address:  {data.get('project_address', '')}",
        f"Original Contract Date:  {data.get('original_contract_date', '[TBD]')}",
        f"Valid Through:  {validity_date}",
    ]
    for line in header_lines:
        p = doc.add_paragraph()
        run = p.add_run(line)
        run.font.size = Pt(10)

    doc.add_paragraph()
    _add_rule(doc)

    for header in SECTION_HEADERS:
        if header in sections:
            p = doc.add_paragraph()
            run = p.add_run(header)
            run.bold = True
            run.font.size = Pt(10)

            p = doc.add_paragraph()
            font_name = 'Courier New' if header == "COST SUMMARY" else 'Calibri'
            run = p.add_run(sections[header])
            run.font.size = Pt(10)
            run.font.name = font_name

            doc.add_paragraph()

    # Signature blocks
    doc.add_paragraph()
    sig_lines = [
        "Contractor: ____________________________________________   Date: ________________",
        "",
        "Print Name: ___________________________________________",
        "",
        "",
        "Owner: _________________________________________________   Date: ________________",
        "",
        "Print Name: ___________________________________________",
    ]
    for line in sig_lines:
        p = doc.add_paragraph()
        run = p.add_run(line)
        run.font.size = Pt(10)

    doc.save(output_path)


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def generate(data: dict, output_dir: str = "./output") -> dict:
    """
    Generate a PCO from structured field data.

    Args:
        data: dict with keys:
            company_name, project_name, project_address,
            change_order_number, date, original_contract_date,
            scope_description, labor_cost, material_cost,
            schedule_days, reason_for_change
        output_dir: where to write PDF and Word files

    Returns:
        dict with keys: pdf, word, generated_text, costs, sections
    """

    Path(output_dir).mkdir(parents=True, exist_ok=True)

    # Calculate markup
    labor     = parse_cost(data.get('labor_cost', ''))
    materials = parse_cost(data.get('material_cost', ''))
    costs     = apply_markup(labor, materials)
    cost_block = format_cost_block(costs)

    co_num = data.get('change_order_number', '001')
    slug   = data.get('project_name', 'project').replace(' ', '_').lower()[:30]
    date_str = datetime.now().strftime('%Y%m%d')
    base   = f"pco_{co_num}_{slug}_{date_str}"

    pdf_path  = os.path.join(output_dir, f"{base}.pdf")
    word_path = os.path.join(output_dir, f"{base}.docx")

    print("Calling Claude (CHG-51 grounded prompt)...")
    generated_text = generate_pco_text(data)

    print("Parsing sections and applying markup...")
    sections = parse_sections(generated_text, cost_block)

    print(f"Building PDF  → {pdf_path}")
    build_pdf(data, sections, costs, pdf_path)

    print(f"Building Word → {word_path}")
    build_word(data, sections, costs, word_path)

    print("Done.")
    return {
        "pdf":            pdf_path,
        "word":           word_path,
        "generated_text": generated_text,
        "sections":       sections,
        "costs":          costs,
    }
