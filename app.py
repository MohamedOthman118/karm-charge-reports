"""
KARM CHARGE — EV Charger Operations & Diagnostic Report Generator
─────────────────────────────────────────────────────────────────
Workflow:
  1. Copy the AI prompt shown in the app.
  2. Paste it + your raw field notes into ChatGPT / Claude / Gemini / etc.
  3. The AI returns a structured JSON payload.
  4. Paste that JSON into the box on this site.
  5. Click "Generate Report" → download the .docx file.

No API keys. No external calls. 100% offline document generation.
"""

import io
import json
from datetime import date, datetime

import streamlit as st
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

# ═══════════════════════════════════════════════════════════════════════════════
# THE AI PROMPT TEMPLATE
# ═══════════════════════════════════════════════════════════════════════════════

AI_PROMPT = """You are an EV Charging Infrastructure Operations Specialist working for KARM CHARGE.

I will paste raw field notes from a site visit. Your job is to convert them into a structured JSON object that follows the exact schema below. Do not invent data — if a field is not mentioned in the notes, leave it as an empty string, an empty list, or 0 as appropriate.

For the "executive_summary" field, rephrase the raw notes into a professional, concise paragraph (3-6 sentences) focusing on technical diagnostics, root causes, and actions taken. Use clear business language suitable for a management report — no bullets, no emojis.

Return ONLY valid JSON, no markdown fences, no commentary before or after.

JSON SCHEMA:
{
  "site": "e.g. Arkan — Crown Plaza",
  "visit_date": "YYYY-MM-DD",
  "report_id": "OPS-XXX-YYYY-MM-DD  (or leave empty to auto-generate)",
  "prepared_by": "Mohamed Othman",
  "role": "Customer Ops Specialist",
  "personnel": [
    "Eng. Mohamed Medhat — Solargy, Engineer",
    "Maysara — Solargy, Technician"
  ],
  "executive_summary": "A professional 3-6 sentence summary written in your own words.",
  "kpis": {
    "total_inspected": 4,
    "operational": 0,
    "offline": 4,
    "open_actions": 2
  },
  "asset_status": [
    {
      "asset_id": "AR-01-EVAC-01-F10 (Y13)",
      "location": "Arkan",
      "physical_status": "Out of service",
      "dashboard_status": "Offline",
      "severity": "Critical"
    }
  ],
  "diagnostics": [
    {
      "issue": "Charger Fault — Continuous Indicator Light Cycling",
      "affected_units": "AR-01-EVAC-01-F10 (Y13)",
      "symptom": "Charger cycles through red/blue/green; unit non-functional.",
      "root_cause": "Communication board issue per Solargy report.",
      "action_taken": "Swapped circuit breaker wire from I10 to Y13 — fault persisted.",
      "status": "Open"
    }
  ],
  "open_action_items": [
    {
      "issue": "Y13 offline awaiting comms-board quote",
      "next_step": "Receive and approve Solargy quotation",
      "owner": "Solargy",
      "priority": "Critical",
      "target_date": "30 May 2026"
    }
  ]
}

RAW FIELD NOTES:
<<<PASTE YOUR NOTES HERE>>>
"""

# ═══════════════════════════════════════════════════════════════════════════════
# SAMPLE JSON (loaded by the "Load sample" button)
# ═══════════════════════════════════════════════════════════════════════════════

SAMPLE_JSON = {
    "site": "Arkan — Crown Plaza",
    "visit_date": "2026-05-17",
    "report_id": "OPS-ARK-2026-05-17",
    "prepared_by": "Mohamed Othman",
    "role": "Customer Ops Specialist",
    "personnel": [
        "Eng. Mohamed Medhat — Solargy, Engineer",
        "Maysara — Solargy, Technician",
        "Abdullah — KarmSolar, Technician",
    ],
    "executive_summary": (
        "On 17 May 2026, a joint site visit was conducted at Arkan — Crown Plaza to "
        "diagnose the persistent fault on charger AR-01-EVAC-01-F10 (Y13). The unit "
        "continued to exhibit a continuous red/blue/green indicator cycle and was "
        "confirmed non-functional. A breaker-swap test ruled out the upstream supply, "
        "pointing to a communication-board failure as the most likely root cause. "
        "Ethernet access to the board could not be established on-site, leaving the "
        "diagnosis inconclusive. A subsequent update on 19 May 2026 reported that "
        "AR-02-EVAC-02-H30 has also faulted and requires a software update from Solargy."
    ),
    "kpis": {
        "total_inspected": 4,
        "operational": 0,
        "offline": 4,
        "open_actions": 2,
    },
    "asset_status": [
        {"asset_id": "AR-01-EVAC-01-F10 (Y13)", "location": "Arkan",
         "physical_status": "Out of service", "dashboard_status": "Offline",
         "severity": "Critical"},
        {"asset_id": "AR-02-EVAC-02-H30", "location": "Arkan",
         "physical_status": "Out of service", "dashboard_status": "Offline",
         "severity": "Critical"},
        {"asset_id": "AR-02-EVAC-02-H29", "location": "Arkan",
         "physical_status": "Out of service", "dashboard_status": "Offline",
         "severity": "Critical"},
        {"asset_id": "AR-01-EVAC-02-I10", "location": "Arkan",
         "physical_status": "Out of service", "dashboard_status": "Offline",
         "severity": "Critical"},
    ],
    "diagnostics": [
        {
            "issue": "Charger Fault — Continuous Indicator Light Cycling (Y13)",
            "affected_units": "AR-01-EVAC-01-F10 (Y13)",
            "symptom": "Charger continuously cycling through red, blue, and green indicator lights; unit non-functional.",
            "root_cause": "Communication board issue per Solargy's report.",
            "action_taken": "Circuit-breaker wire removed from AR-01-EVAC-02-I10 and installed in Y13 to test breaker hypothesis — fault persisted.",
            "status": "Open",
        },
        {
            "issue": "Communication Board — Inconclusive Diagnosis",
            "affected_units": "AR-01-EVAC-01-F10 (Y13)",
            "symptom": "Unable to establish Ethernet connection to communication board; IP address inaccessible.",
            "root_cause": "Unclear — either failed comms board or credentials/configuration issue.",
            "action_taken": "Connection attempted via laptop over Ethernet; IP access failed. No further testing completed on-site.",
            "status": "Open — requires re-test by qualified engineer with confirmed credentials",
        },
        {
            "issue": "AR-02-EVAC-02-H30 Fault (Update — 19 May 2026)",
            "affected_units": "AR-02-EVAC-02-H30",
            "symptom": "Previously operational unit now faulted; was the sole remaining working charger at Arkan.",
            "root_cause": "Software update required per Abdelrahman Medhat (Solargy).",
            "action_taken": "Hussainy informed; Solargy asked to deploy update.",
            "status": "Open — awaiting Solargy",
        },
    ],
    "open_action_items": [
        {
            "issue": "AR-02-EVAC-02-H30 faulted — software update required",
            "next_step": "Solargy to deploy software update and re-verify operation",
            "owner": "Solargy",
            "priority": "Critical",
            "target_date": "30 May 2026",
        },
        {
            "issue": "AR-01-EVAC-01-F10 (Y13) offline",
            "next_step": "Awaiting Solargy quotation for the communication board replacement",
            "owner": "Solargy",
            "priority": "Critical",
            "target_date": "—",
        },
    ],
}

# ═══════════════════════════════════════════════════════════════════════════════
# DOCX STYLING HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

BRAND_GREEN = RGBColor(0x2E, 0x86, 0x48)
HEADER_BLUE = RGBColor(0x1F, 0x49, 0x7D)
WHITE_TEXT  = RGBColor(0xFF, 0xFF, 0xFF)
GREY_TEXT   = RGBColor(0x60, 0x60, 0x60)

LIGHT_FILL  = "D9EAD3"   # pale green
DARK_FILL   = "1F497D"   # dark blue
ZEBRA_FILL  = "EAF4EA"   # very pale green


def _set_cell_fill(cell, hex_color: str):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)


def _bold_run(p, text, size_pt=11, color=None, italic=False):
    r = p.add_run(text)
    r.bold = True
    r.italic = italic
    r.font.size = Pt(size_pt)
    if color:
        r.font.color.rgb = color
    return r


def _normal_run(p, text, size_pt=10):
    r = p.add_run(text)
    r.font.size = Pt(size_pt)
    return r


def _add_section_heading(doc, text):
    """Blue underlined section heading."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after  = Pt(4)
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(13)
    run.font.color.rgb = HEADER_BLUE
    # bottom border
    pPr  = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bot  = OxmlElement("w:bottom")
    bot.set(qn("w:val"),   "single")
    bot.set(qn("w:sz"),    "6")
    bot.set(qn("w:space"), "1")
    bot.set(qn("w:color"), "1F497D")
    pBdr.append(bot)
    pPr.append(pBdr)


def _add_kv_table(doc, rows):
    """Two-column label/value table."""
    table = doc.add_table(rows=len(rows), cols=2)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    for ri, (label, value) in enumerate(rows):
        lc, vc = table.rows[ri].cells
        _set_cell_fill(lc, LIGHT_FILL)
        lp = lc.paragraphs[0]
        lp.paragraph_format.space_before = Pt(3)
        lp.paragraph_format.space_after  = Pt(3)
        _bold_run(lp, label, size_pt=10, color=HEADER_BLUE)
        vp = vc.paragraphs[0]
        vp.paragraph_format.space_before = Pt(3)
        vp.paragraph_format.space_after  = Pt(3)
        _normal_run(vp, str(value), size_pt=10)


def _add_data_table(doc, columns, data_rows):
    """Header row + striped data rows."""
    table = doc.add_table(rows=1 + len(data_rows), cols=len(columns))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    # header
    for ci, col in enumerate(columns):
        cell = table.rows[0].cells[ci]
        _set_cell_fill(cell, DARK_FILL)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after  = Pt(3)
        _bold_run(p, col, size_pt=10, color=WHITE_TEXT)
    # data
    for ri, data_row in enumerate(data_rows):
        fill = ZEBRA_FILL if ri % 2 == 0 else "FFFFFF"
        for ci, value in enumerate(data_row):
            cell = table.rows[ri + 1].cells[ci]
            _set_cell_fill(cell, fill)
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after  = Pt(2)
            _normal_run(p, str(value), size_pt=10)


def _add_kpi_table(doc, kpis):
    """kpis = list of (label, value, hex_fill)."""
    table = doc.add_table(rows=1, cols=len(kpis))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    for ci, (label, value, fill_hex) in enumerate(kpis):
        cell = table.rows[0].cells[ci]
        _set_cell_fill(cell, fill_hex)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after  = Pt(8)
        vr = p.add_run(f"{value}\n")
        vr.bold = True
        vr.font.size = Pt(22)
        vr.font.color.rgb = WHITE_TEXT
        lr = p.add_run(label.upper())
        lr.bold = True
        lr.font.size = Pt(9)
        lr.font.color.rgb = WHITE_TEXT


def _add_at_a_glance_table(doc, summary_text):
    table = doc.add_table(rows=2, cols=1)
    table.style = "Table Grid"
    # title
    tc = table.rows[0].cells[0]
    _set_cell_fill(tc, DARK_FILL)
    tp = tc.paragraphs[0]
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tp.paragraph_format.space_before = Pt(4)
    tp.paragraph_format.space_after  = Pt(4)
    _bold_run(tp, "AT A GLANCE", size_pt=11, color=WHITE_TEXT)
    # body
    bc = table.rows[1].cells[0]
    _set_cell_fill(bc, "F9FDF9")
    bp = bc.paragraphs[0]
    bp.paragraph_format.space_before = Pt(6)
    bp.paragraph_format.space_after  = Pt(6)
    _normal_run(bp, summary_text or "(No summary provided)", size_pt=10)


# ═══════════════════════════════════════════════════════════════════════════════
# DOCX BUILDER
# ═══════════════════════════════════════════════════════════════════════════════

def build_docx(data: dict) -> bytes:
    """Assemble the full report from a validated JSON dict."""
    doc = Document()
    for section in doc.sections:
        section.top_margin = section.bottom_margin = Inches(1)
        section.left_margin = section.right_margin = Inches(1)

    # ── Title block ──────────────────────────────────────────────────────────
    tp = doc.add_paragraph()
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = tp.add_run("KARM CHARGE")
    tr.bold = True
    tr.font.size = Pt(22)
    tr.font.color.rgb = BRAND_GREEN

    sp = doc.add_paragraph()
    sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr = sp.add_run("EV Charger Operations & Diagnostic Report")
    sr.bold = True
    sr.font.size = Pt(14)
    sr.font.color.rgb = HEADER_BLUE

    gp = doc.add_paragraph()
    gp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    gr = gp.add_run("Site Visit Operations Report")
    gr.italic = True
    gr.font.size = Pt(10)
    gr.font.color.rgb = GREY_TEXT

    doc.add_paragraph()

    # ── Document Control ─────────────────────────────────────────────────────
    visit_str = data["visit_date_str"]
    _add_section_heading(doc, "Document Control")
    _add_kv_table(doc, [
        ("Report ID",       data["report_id"]),
        ("Site / District", data.get("site", "")),
        ("Date of Visit",   visit_str),
        ("Prepared By",     data.get("prepared_by", "Mohamed Othman")),
        ("Role",            data.get("role", "Customer Ops Specialist")),
        ("Classification",  "Internal – Operations"),
    ])
    doc.add_paragraph()

    # ── Accompanying Personnel ───────────────────────────────────────────────
    _add_section_heading(doc, "Accompanying Personnel")
    personnel = [p.strip() for p in data.get("personnel", []) if str(p).strip()]
    if personnel:
        for person in personnel:
            bp = doc.add_paragraph(style="List Bullet")
            bp.paragraph_format.space_before = Pt(2)
            bp.paragraph_format.space_after  = Pt(2)
            _normal_run(bp, person, size_pt=10)
    else:
        doc.add_paragraph("No accompanying personnel listed.").runs[0].font.size = Pt(10)
    doc.add_paragraph()

    # ── 1. Executive Summary ─────────────────────────────────────────────────
    _add_section_heading(doc, "1. Executive Summary")
    _add_at_a_glance_table(doc, data.get("executive_summary", ""))
    doc.add_paragraph()

    # ── 2. KPIs ──────────────────────────────────────────────────────────────
    _add_section_heading(doc, "2. Visit KPIs at a Glance")
    kpis = data.get("kpis", {})
    _add_kpi_table(doc, [
        ("Total Inspected",   kpis.get("total_inspected", 0), "2E8648"),
        ("Operational",       kpis.get("operational", 0),     "1F497D"),
        ("Offline / Faulted", kpis.get("offline", 0),         "C0392B"),
        ("Open Actions",      kpis.get("open_actions", 0),    "D35400"),
    ])
    doc.add_paragraph()

    # ── 3. Asset Status ──────────────────────────────────────────────────────
    _add_section_heading(doc, "3. Asset Status & Observations")
    asset_rows = []
    for a in data.get("asset_status", []):
        asset_rows.append([
            a.get("asset_id", ""),
            a.get("location", ""),
            a.get("physical_status", ""),
            a.get("dashboard_status", ""),
            a.get("severity", ""),
        ])
    if asset_rows:
        _add_data_table(
            doc,
            ["Asset ID", "Location", "Physical Status", "Dashboard Status", "Severity"],
            asset_rows,
        )
    else:
        doc.add_paragraph("No asset records provided.").runs[0].font.size = Pt(10)
    doc.add_paragraph()

    # ── 4. Diagnostics ───────────────────────────────────────────────────────
    _add_section_heading(doc, "4. Diagnostics & Actions Taken")
    diagnostics = data.get("diagnostics", [])
    if not diagnostics:
        doc.add_paragraph("No diagnostic entries provided.").runs[0].font.size = Pt(10)
    else:
        for idx, d in enumerate(diagnostics, start=1):
            # Issue heading
            ih = doc.add_paragraph()
            ih.paragraph_format.space_before = Pt(8)
            ih.paragraph_format.space_after  = Pt(2)
            _bold_run(ih, f"Issue {idx}: {d.get('issue', '—')}",
                      size_pt=11, color=HEADER_BLUE)
            # Detail bullets
            fields = [
                ("Affected Units", d.get("affected_units", "")),
                ("Symptom",        d.get("symptom",        "")),
                ("Root Cause",     d.get("root_cause",     "")),
                ("Action Taken",   d.get("action_taken",   "")),
                ("Status",         d.get("status",         "")),
            ]
            for label, value in fields:
                if str(value).strip():
                    bp = doc.add_paragraph(style="List Bullet")
                    bp.paragraph_format.space_before = Pt(1)
                    bp.paragraph_format.space_after  = Pt(1)
                    _bold_run(bp, f"{label}: ", size_pt=10)
                    _normal_run(bp, str(value), size_pt=10)
    doc.add_paragraph()

    # ── 5. Open Action Items ─────────────────────────────────────────────────
    _add_section_heading(doc, "5. Open Action Items")
    action_rows = []
    for i, a in enumerate(data.get("open_action_items", []), start=1):
        action_rows.append([
            str(i),
            a.get("issue", ""),
            a.get("next_step", ""),
            a.get("owner", ""),
            a.get("priority", ""),
            a.get("target_date", ""),
        ])
    if action_rows:
        _add_data_table(
            doc,
            ["#", "Issue", "Next Step", "Owner", "Priority", "Target Date"],
            action_rows,
        )
    else:
        doc.add_paragraph("No open action items logged.").runs[0].font.size = Pt(10)

    # ── Footer line ──────────────────────────────────────────────────────────
    fp = doc.add_paragraph()
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.paragraph_format.space_before = Pt(24)
    fr = fp.add_run(
        f"KARM CHARGE  ·  {data['report_id']}  ·  Confidential – Internal Use Only"
    )
    fr.italic = True
    fr.font.size = Pt(8)
    fr.font.color.rgb = GREY_TEXT

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════════
# JSON VALIDATION & NORMALISATION
# ═══════════════════════════════════════════════════════════════════════════════

def _strip_markdown_fences(raw: str) -> str:
    """AIs sometimes wrap JSON in ```json ... ```. Strip that out."""
    s = raw.strip()
    if s.startswith("```"):
        # Drop the first fence line
        s = s.split("\n", 1)[1] if "\n" in s else s
        # Drop trailing fence
        if s.rstrip().endswith("```"):
            s = s.rstrip()[:-3]
    return s.strip()


def parse_and_validate(raw_json: str) -> tuple[dict | None, list[str]]:
    """
    Parse the pasted JSON. Returns (normalised_data_dict, warnings_list).
    On hard parse failure, returns (None, [error_string]).
    """
    warnings: list[str] = []

    cleaned = _strip_markdown_fences(raw_json)
    if not cleaned:
        return None, ["JSON is empty. Paste the AI's response into the box above."]

    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError as e:
        return None, [f"❌ Invalid JSON: {e.msg} (line {e.lineno}, column {e.colno})"]

    if not isinstance(data, dict):
        return None, ["❌ Top-level JSON must be an object {...}, not a list or value."]

    # ── Normalise & default missing fields ───────────────────────────────────
    def _get(key, default, expected_type=None):
        val = data.get(key, default)
        if expected_type and not isinstance(val, expected_type):
            warnings.append(f"⚠️ Field '{key}' should be {expected_type.__name__} — using default.")
            return default
        return val

    site = str(_get("site", "")).strip()

    # Visit date — accept YYYY-MM-DD or fall back to today
    raw_date = str(_get("visit_date", "")).strip()
    visit_dt: date
    try:
        visit_dt = datetime.strptime(raw_date, "%Y-%m-%d").date()
    except (ValueError, TypeError):
        visit_dt = date.today()
        if raw_date:
            warnings.append(
                f"⚠️ Could not parse visit_date '{raw_date}' — expected YYYY-MM-DD. "
                f"Defaulted to today ({visit_dt.isoformat()})."
            )

    # Report ID — auto-generate if blank
    report_id = str(_get("report_id", "")).strip()
    if not report_id:
        prefix = "".join(c for c in site.upper() if c.isalpha())[:3] or "GEN"
        report_id = f"OPS-{prefix}-{visit_dt.isoformat()}"
        warnings.append(f"ℹ️ report_id was blank — auto-generated as '{report_id}'.")

    # KPIs — coerce to int
    kpis_raw = _get("kpis", {}, dict)
    kpis = {}
    for k in ("total_inspected", "operational", "offline", "open_actions"):
        try:
            kpis[k] = int(kpis_raw.get(k, 0))
        except (TypeError, ValueError):
            kpis[k] = 0
            warnings.append(f"⚠️ KPI '{k}' was not a number — defaulted to 0.")

    # Lists — ensure they're lists
    personnel  = _get("personnel",         [], list)
    assets     = _get("asset_status",      [], list)
    diagnostics= _get("diagnostics",       [], list)
    actions    = _get("open_action_items", [], list)

    # Build the normalised structure used by build_docx
    normalised = {
        "site":              site,
        "visit_date":        visit_dt.isoformat(),
        "visit_date_str":    visit_dt.strftime("%A – %d %B %Y"),
        "report_id":         report_id,
        "prepared_by":       str(_get("prepared_by", "Mohamed Othman")).strip() or "Mohamed Othman",
        "role":              str(_get("role", "Customer Ops Specialist")).strip() or "Customer Ops Specialist",
        "personnel":         personnel,
        "executive_summary": str(_get("executive_summary", "")).strip(),
        "kpis":              kpis,
        "asset_status":      assets,
        "diagnostics":       diagnostics,
        "open_action_items": actions,
    }

    if not normalised["executive_summary"]:
        warnings.append("⚠️ 'executive_summary' is empty — the AT A GLANCE box will show a placeholder.")

    return normalised, warnings


# ═══════════════════════════════════════════════════════════════════════════════
# STREAMLIT UI
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    st.set_page_config(
        page_title="KARM CHARGE — Report Generator",
        page_icon="⚡",
        layout="wide",
        initial_sidebar_state="collapsed",
    )

    # Initialise session state for the textarea
    if "json_text" not in st.session_state:
        st.session_state["json_text"] = ""

    # ── Banner ───────────────────────────────────────────────────────────────
    st.markdown(
        """
        <div style='background:linear-gradient(90deg,#1F497D,#2E8648);
                    padding:18px 24px;border-radius:8px;margin-bottom:12px'>
          <h2 style='color:white;margin:0'>⚡ KARM CHARGE</h2>
          <p style='color:#d0eaff;margin:4px 0 0'>
            EV Charger Operations &amp; Diagnostic Report Generator
          </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ── Quick how-to ─────────────────────────────────────────────────────────
    st.markdown(
        """
        **How to use this tool**
        1. Copy the **AI prompt** below.
        2. Paste it into any AI (ChatGPT, Claude, Gemini, etc.) and append your raw field notes where indicated.
        3. The AI will return a **JSON object**.
        4. Paste that JSON into the box on the right and click **Generate Report**.
        """
    )

    st.divider()

    # ═══════════════════════════════════════════════════════════════════════════
    # Two-column layout: prompt (left) | JSON input + actions (right)
    # ═══════════════════════════════════════════════════════════════════════════
    left, right = st.columns([1, 1], gap="large")

    # ── LEFT: AI prompt (read-only, copyable) ────────────────────────────────
    with left:
        st.subheader("📋 Step 1 — Copy this prompt into your AI")
        st.caption("Click the copy icon in the top-right corner of the box below.")
        st.code(AI_PROMPT, language="markdown")

        with st.expander("💡 Don't have an AI handy? See alternatives"):
            st.markdown(
                """
                - **ChatGPT** — https://chat.openai.com
                - **Claude** — https://claude.ai
                - **Gemini** — https://gemini.google.com
                - **Mistral** — https://chat.mistral.ai

                Any chatbot that can follow instructions and return JSON will work.
                """
            )

    # ── RIGHT: JSON input + generate ─────────────────────────────────────────
    with right:
        st.subheader("📥 Step 2 — Paste the AI's JSON response")

        col_a, col_b = st.columns([1, 1])
        with col_a:
            if st.button("📝 Load sample JSON", use_container_width=True,
                         help="Loads a worked example so you can see the expected format."):
                st.session_state["json_text"] = json.dumps(SAMPLE_JSON, indent=2, ensure_ascii=False)
                st.rerun()
        with col_b:
            if st.button("🗑️ Clear", use_container_width=True):
                st.session_state["json_text"] = ""
                # Also clear any cached doc
                for k in ("docx_bytes", "docx_filename", "parsed_data", "warnings"):
                    st.session_state.pop(k, None)
                st.rerun()

        json_text = st.text_area(
            "JSON",
            value=st.session_state["json_text"],
            height=400,
            placeholder='{\n  "site": "...",\n  "visit_date": "2026-05-17",\n  ...\n}',
            label_visibility="collapsed",
            key="json_text_widget",
        )
        # Keep state in sync
        st.session_state["json_text"] = json_text

        generate = st.button(
            "🚀 Generate Report",
            type="primary",
            use_container_width=True,
        )

    # ─────────────────────────────────────────────────────────────────────────
    # Generation flow
    # ─────────────────────────────────────────────────────────────────────────
    if generate:
        data, warnings = parse_and_validate(st.session_state["json_text"])
        if data is None:
            for msg in warnings:
                st.error(msg)
        else:
            try:
                docx_bytes = build_docx(data)
                st.session_state["docx_bytes"]    = docx_bytes
                st.session_state["docx_filename"] = f"{data['report_id']}.docx"
                st.session_state["parsed_data"]   = data
                st.session_state["warnings"]      = warnings
                st.success("✅ Report generated successfully!")
            except Exception as exc:  # noqa: BLE001
                st.exception(exc)

    # ─────────────────────────────────────────────────────────────────────────
    # Output / preview area
    # ─────────────────────────────────────────────────────────────────────────
    if st.session_state.get("docx_bytes"):
        st.divider()
        data = st.session_state["parsed_data"]

        # Warnings (if any)
        for w in st.session_state.get("warnings", []):
            st.warning(w)

        # Download button — front and centre
        st.download_button(
            label    = f"⬇️ Download  {st.session_state['docx_filename']}",
            data     = st.session_state["docx_bytes"],
            file_name= st.session_state["docx_filename"],
            mime     = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            type="primary",
        )

        st.divider()
        st.subheader("📄 Report Preview")

        # Document Control mini-view
        c1, c2, c3 = st.columns(3)
        c1.metric("Report ID",       data["report_id"])
        c2.metric("Site / District", data["site"] or "—")
        c3.metric("Date of Visit",   data["visit_date_str"])

        # KPIs
        st.markdown("##### Visit KPIs")
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Total Inspected",  data["kpis"]["total_inspected"])
        k2.metric("Operational",      data["kpis"]["operational"])
        k3.metric("Offline / Faulted",data["kpis"]["offline"])
        k4.metric("Open Actions",     data["kpis"]["open_actions"])

        # Executive summary
        with st.expander("📝 Executive Summary", expanded=True):
            st.markdown(
                f"""<div style='background:#f0f7f0;border-left:4px solid #2E8648;
                                padding:14px 16px;border-radius:6px;
                                font-size:14px;line-height:1.6;white-space:pre-wrap'>
                {data['executive_summary'] or '(empty)'}</div>""",
                unsafe_allow_html=True,
            )

        # Personnel
        with st.expander(f"👥 Personnel ({len(data['personnel'])})"):
            if data["personnel"]:
                for p in data["personnel"]:
                    st.markdown(f"- {p}")
            else:
                st.caption("— none —")

        # Asset table
        with st.expander(f"🔋 Asset Status ({len(data['asset_status'])} units)"):
            if data["asset_status"]:
                st.dataframe(data["asset_status"], use_container_width=True, hide_index=True)
            else:
                st.caption("— none —")

        # Diagnostics
        with st.expander(f"🔧 Diagnostics ({len(data['diagnostics'])} issues)"):
            for i, d in enumerate(data["diagnostics"], start=1):
                st.markdown(f"**Issue {i}: {d.get('issue', '—')}**")
                for label, key in [
                    ("Affected Units", "affected_units"),
                    ("Symptom",        "symptom"),
                    ("Root Cause",     "root_cause"),
                    ("Action Taken",   "action_taken"),
                    ("Status",         "status"),
                ]:
                    if str(d.get(key, "")).strip():
                        st.markdown(f"- **{label}:** {d[key]}")
                st.markdown("")

        # Open actions
        with st.expander(f"📋 Open Action Items ({len(data['open_action_items'])})"):
            if data["open_action_items"]:
                st.dataframe(data["open_action_items"], use_container_width=True, hide_index=True)
            else:
                st.caption("— none —")


if __name__ == "__main__":
    main()
