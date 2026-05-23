"""
KARM CHARGE — EV Charger Operations & Diagnostic Report Generator
─────────────────────────────────────────────────────────────────
Workflow:
  1. Copy the AI prompt → paste it + raw field notes into ChatGPT/Claude/Gemini.
  2. AI returns JSON.
  3. Paste JSON here → click "Load JSON" → every form field gets pre-filled.
  4. Review and edit ANY field manually before generating.
  5. Click "Generate Report" → download polished .docx.

No API keys. No external calls. Full editorial control before export.
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
# CONSTANTS
# ═══════════════════════════════════════════════════════════════════════════════

SEVERITY_OPTIONS  = ["Critical", "High", "Medium", "Low", "Resolved", "Unknown"]
PHYS_STATUS_OPTS  = ["Operational", "Good", "Out of service", "Degraded", "Under repair", "Other"]
DASH_STATUS_OPTS  = ["Online", "Offline", "Charging", "Available", "Faulted", "Unknown"]
PRIORITY_OPTIONS  = ["Critical", "High", "Medium", "Low"]
DIAG_STATUS_OPTS  = ["Open", "In Progress", "Resolved", "Escalated", "Closed"]

EMPTY_ASSET  = {"asset_id": "", "location": "", "physical_status": "Out of service",
                "dashboard_status": "Offline", "severity": "Critical"}
EMPTY_DIAG   = {"issue": "", "affected_units": "", "symptom": "",
                "root_cause": "", "action_taken": "", "status": "Open"}
EMPTY_ACTION = {"issue": "", "next_step": "", "owner": "",
                "priority": "Medium", "target_date": ""}

# ═══════════════════════════════════════════════════════════════════════════════
# AI PROMPT TEMPLATE (user copies this into their AI of choice)
# ═══════════════════════════════════════════════════════════════════════════════

AI_PROMPT = """You are an EV Charging Infrastructure Operations Specialist working for KARM CHARGE.

I will paste raw field notes from a site visit. Your job is to convert them into a structured JSON object that follows the exact schema below. Do not invent data — if a field is not mentioned in the notes, leave it as an empty string, an empty list, or 0 as appropriate.

For "executive_summary", rephrase the raw notes into a professional, concise paragraph (3-6 sentences) focusing on technical diagnostics, root causes, and actions taken. Use clear business language — no bullets, no emojis, no markdown.

Return ONLY valid JSON. No markdown code fences. No commentary before or after.

JSON SCHEMA:
{
  "site": "e.g. Arkan — Crown Plaza",
  "visit_date": "YYYY-MM-DD",
  "report_id": "OPS-XXX-YYYY-MM-DD  (leave empty to auto-generate)",
  "prepared_by": "Mohamed Othman",
  "role": "Customer Ops Specialist",
  "personnel": [
    "Eng. Mohamed Medhat — Solargy, Engineer",
    "Maysara — Solargy, Technician"
  ],
  "executive_summary": "A professional 3-6 sentence summary in your own words.",
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
      "physical_status": "Out of service",   // Operational | Good | Out of service | Degraded | Under repair
      "dashboard_status": "Offline",         // Online | Offline | Charging | Available | Faulted
      "severity": "Critical"                 // Critical | High | Medium | Low | Resolved
    }
  ],
  "diagnostics": [
    {
      "issue": "Charger Fault — Continuous Indicator Light Cycling",
      "affected_units": "AR-01-EVAC-01-F10 (Y13)",
      "symptom": "Continuous red/blue/green cycle; non-functional.",
      "root_cause": "Communication board issue per Solargy report.",
      "action_taken": "Swapped circuit-breaker wire from I10 — fault persisted.",
      "status": "Open"                       // Open | In Progress | Resolved | Escalated | Closed
    }
  ],
  "open_action_items": [
    {
      "issue": "Y13 offline awaiting comms-board quote",
      "next_step": "Receive and approve Solargy quotation",
      "owner": "Solargy",
      "priority": "Critical",                // Critical | High | Medium | Low
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
    "kpis": {"total_inspected": 4, "operational": 0, "offline": 4, "open_actions": 2},
    "asset_status": [
        {"asset_id": "AR-01-EVAC-01-F10 (Y13)", "location": "Arkan",
         "physical_status": "Out of service", "dashboard_status": "Offline", "severity": "Critical"},
        {"asset_id": "AR-02-EVAC-02-H30", "location": "Arkan",
         "physical_status": "Out of service", "dashboard_status": "Offline", "severity": "Critical"},
        {"asset_id": "AR-02-EVAC-02-H29", "location": "Arkan",
         "physical_status": "Out of service", "dashboard_status": "Offline", "severity": "Critical"},
        {"asset_id": "AR-01-EVAC-02-I10", "location": "Arkan",
         "physical_status": "Out of service", "dashboard_status": "Offline", "severity": "Critical"},
    ],
    "diagnostics": [
        {"issue": "Charger Fault — Continuous Indicator Light Cycling (Y13)",
         "affected_units": "AR-01-EVAC-01-F10 (Y13)",
         "symptom": "Charger continuously cycling through red, blue, and green indicator lights; unit non-functional.",
         "root_cause": "Communication board issue per Solargy's report.",
         "action_taken": "Circuit-breaker wire removed from AR-01-EVAC-02-I10 and installed in Y13 — fault persisted.",
         "status": "Open"},
        {"issue": "Communication Board — Inconclusive Diagnosis",
         "affected_units": "AR-01-EVAC-01-F10 (Y13)",
         "symptom": "Unable to establish Ethernet connection; IP address inaccessible.",
         "root_cause": "Unclear — either failed board or credentials/config issue.",
         "action_taken": "Connection attempted via laptop over Ethernet; IP access failed.",
         "status": "Open"},
    ],
    "open_action_items": [
        {"issue": "AR-02-EVAC-02-H30 faulted — software update required",
         "next_step": "Solargy to deploy software update and re-verify operation",
         "owner": "Solargy", "priority": "Critical", "target_date": "30 May 2026"},
        {"issue": "AR-01-EVAC-01-F10 (Y13) offline",
         "next_step": "Awaiting Solargy quotation for communication board replacement",
         "owner": "Solargy", "priority": "Critical", "target_date": "—"},
    ],
}

# ═══════════════════════════════════════════════════════════════════════════════
# JSON PARSING / NORMALISATION
# ═══════════════════════════════════════════════════════════════════════════════

def _strip_markdown_fences(raw: str) -> str:
    """AIs sometimes wrap JSON in ```json ... ```. Strip that out."""
    s = raw.strip()
    if s.startswith("```"):
        s = s.split("\n", 1)[1] if "\n" in s else s
        if s.rstrip().endswith("```"):
            s = s.rstrip()[:-3]
    return s.strip()


def _safe_choice(value: str, options: list[str], default: str) -> str:
    """Return value if it's in options (case-insensitive), else default."""
    if not value:
        return default
    for opt in options:
        if str(value).strip().lower() == opt.lower():
            return opt
    return default


def parse_json(raw_json: str) -> tuple[dict | None, list[str]]:
    """Parse the pasted JSON. Returns (normalised_dict, warnings)."""
    warnings: list[str] = []

    cleaned = _strip_markdown_fences(raw_json)
    if not cleaned:
        return None, ["JSON is empty — paste the AI's response into the box above."]

    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError as e:
        return None, [f"❌ Invalid JSON: {e.msg} (line {e.lineno}, col {e.colno})"]

    if not isinstance(data, dict):
        return None, ["❌ Top-level JSON must be an object {...}."]

    # ── Scalars ──────────────────────────────────────────────────────────────
    site = str(data.get("site", "")).strip()

    raw_date = str(data.get("visit_date", "")).strip()
    try:
        visit_dt = datetime.strptime(raw_date, "%Y-%m-%d").date()
    except (ValueError, TypeError):
        visit_dt = date.today()
        if raw_date:
            warnings.append(f"⚠️ Could not parse visit_date '{raw_date}' — defaulted to today.")

    report_id = str(data.get("report_id", "")).strip()
    if not report_id:
        prefix = "".join(c for c in site.upper() if c.isalpha())[:3] or "GEN"
        report_id = f"OPS-{prefix}-{visit_dt.isoformat()}"

    # ── KPIs ─────────────────────────────────────────────────────────────────
    kpis_raw = data.get("kpis", {}) or {}
    def _kpi(k):
        try:    return int(kpis_raw.get(k, 0))
        except: return 0
    kpis = {
        "total_inspected": _kpi("total_inspected"),
        "operational":     _kpi("operational"),
        "offline":         _kpi("offline"),
        "open_actions":    _kpi("open_actions"),
    }

    # ── Personnel ────────────────────────────────────────────────────────────
    personnel = [str(p).strip() for p in (data.get("personnel") or []) if str(p).strip()]
    if not personnel:
        personnel = [""]

    # ── Asset rows ───────────────────────────────────────────────────────────
    assets = []
    for a in (data.get("asset_status") or []):
        if not isinstance(a, dict):
            continue
        assets.append({
            "asset_id":         str(a.get("asset_id", "")).strip(),
            "location":         str(a.get("location", "")).strip(),
            "physical_status":  _safe_choice(a.get("physical_status"), PHYS_STATUS_OPTS, "Out of service"),
            "dashboard_status": _safe_choice(a.get("dashboard_status"), DASH_STATUS_OPTS, "Offline"),
            "severity":         _safe_choice(a.get("severity"), SEVERITY_OPTIONS, "Critical"),
        })
    if not assets:
        assets = [dict(EMPTY_ASSET)]

    # ── Diagnostic entries ───────────────────────────────────────────────────
    diags = []
    for d in (data.get("diagnostics") or []):
        if not isinstance(d, dict):
            continue
        diags.append({
            "issue":          str(d.get("issue", "")).strip(),
            "affected_units": str(d.get("affected_units", "")).strip(),
            "symptom":        str(d.get("symptom", "")).strip(),
            "root_cause":     str(d.get("root_cause", "")).strip(),
            "action_taken":   str(d.get("action_taken", "")).strip(),
            "status":         _safe_choice(d.get("status"), DIAG_STATUS_OPTS, "Open"),
        })
    if not diags:
        diags = [dict(EMPTY_DIAG)]

    # ── Open action items ────────────────────────────────────────────────────
    actions = []
    for a in (data.get("open_action_items") or []):
        if not isinstance(a, dict):
            continue
        actions.append({
            "issue":       str(a.get("issue", "")).strip(),
            "next_step":   str(a.get("next_step", "")).strip(),
            "owner":       str(a.get("owner", "")).strip(),
            "priority":    _safe_choice(a.get("priority"), PRIORITY_OPTIONS, "Medium"),
            "target_date": str(a.get("target_date", "")).strip(),
        })
    if not actions:
        actions = [dict(EMPTY_ACTION)]

    return {
        "site":              site,
        "visit_date":        visit_dt,
        "report_id":         report_id,
        "prepared_by":       str(data.get("prepared_by", "Mohamed Othman")).strip() or "Mohamed Othman",
        "role":              str(data.get("role", "Customer Ops Specialist")).strip() or "Customer Ops Specialist",
        "personnel":         personnel,
        "executive_summary": str(data.get("executive_summary", "")).strip(),
        "kpis":              kpis,
        "asset_status":      assets,
        "diagnostics":       diags,
        "open_action_items": actions,
    }, warnings


# ═══════════════════════════════════════════════════════════════════════════════
# DOCX STYLING HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

BRAND_GREEN = RGBColor(0x2E, 0x86, 0x48)
HEADER_BLUE = RGBColor(0x1F, 0x49, 0x7D)
WHITE_TEXT  = RGBColor(0xFF, 0xFF, 0xFF)
GREY_TEXT   = RGBColor(0x60, 0x60, 0x60)
LIGHT_FILL  = "D9EAD3"
DARK_FILL   = "1F497D"
ZEBRA_FILL  = "EAF4EA"


def _set_cell_fill(cell, hex_color):
    tc = cell._tc; tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def _bold_run(p, text, size_pt=11, color=None, italic=False):
    r = p.add_run(text); r.bold = True; r.italic = italic; r.font.size = Pt(size_pt)
    if color: r.font.color.rgb = color
    return r


def _normal_run(p, text, size_pt=10):
    r = p.add_run(text); r.font.size = Pt(size_pt); return r


def _add_section_heading(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after  = Pt(4)
    run = p.add_run(text); run.bold = True; run.font.size = Pt(13); run.font.color.rgb = HEADER_BLUE
    pPr = p._p.get_or_add_pPr(); pBdr = OxmlElement("w:pBdr"); bot = OxmlElement("w:bottom")
    bot.set(qn("w:val"), "single"); bot.set(qn("w:sz"), "6"); bot.set(qn("w:space"), "1"); bot.set(qn("w:color"), "1F497D")
    pBdr.append(bot); pPr.append(pBdr)


def _add_kv_table(doc, rows):
    table = doc.add_table(rows=len(rows), cols=2)
    table.style = "Table Grid"; table.alignment = WD_TABLE_ALIGNMENT.LEFT
    for ri, (label, value) in enumerate(rows):
        lc, vc = table.rows[ri].cells
        _set_cell_fill(lc, LIGHT_FILL)
        lp = lc.paragraphs[0]; lp.paragraph_format.space_before = Pt(3); lp.paragraph_format.space_after = Pt(3)
        _bold_run(lp, label, size_pt=10, color=HEADER_BLUE)
        vp = vc.paragraphs[0]; vp.paragraph_format.space_before = Pt(3); vp.paragraph_format.space_after = Pt(3)
        _normal_run(vp, str(value), size_pt=10)


def _add_data_table(doc, columns, data_rows):
    table = doc.add_table(rows=1 + len(data_rows), cols=len(columns))
    table.style = "Table Grid"; table.alignment = WD_TABLE_ALIGNMENT.LEFT
    for ci, col in enumerate(columns):
        cell = table.rows[0].cells[ci]; _set_cell_fill(cell, DARK_FILL)
        p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(3); p.paragraph_format.space_after = Pt(3)
        _bold_run(p, col, size_pt=10, color=WHITE_TEXT)
    for ri, data_row in enumerate(data_rows):
        fill = ZEBRA_FILL if ri % 2 == 0 else "FFFFFF"
        for ci, value in enumerate(data_row):
            cell = table.rows[ri + 1].cells[ci]; _set_cell_fill(cell, fill)
            p = cell.paragraphs[0]; p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(2)
            _normal_run(p, str(value), size_pt=10)


def _add_kpi_table(doc, kpis):
    table = doc.add_table(rows=1, cols=len(kpis))
    table.style = "Table Grid"; table.alignment = WD_TABLE_ALIGNMENT.LEFT
    for ci, (label, value, fill_hex) in enumerate(kpis):
        cell = table.rows[0].cells[ci]; _set_cell_fill(cell, fill_hex)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(8); p.paragraph_format.space_after = Pt(8)
        vr = p.add_run(f"{value}\n"); vr.bold = True; vr.font.size = Pt(22); vr.font.color.rgb = WHITE_TEXT
        lr = p.add_run(label.upper()); lr.bold = True; lr.font.size = Pt(9); lr.font.color.rgb = WHITE_TEXT


def _add_at_a_glance_table(doc, summary_text):
    table = doc.add_table(rows=2, cols=1); table.style = "Table Grid"
    tc = table.rows[0].cells[0]; _set_cell_fill(tc, DARK_FILL)
    tp = tc.paragraphs[0]; tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tp.paragraph_format.space_before = Pt(4); tp.paragraph_format.space_after = Pt(4)
    _bold_run(tp, "AT A GLANCE", size_pt=11, color=WHITE_TEXT)
    bc = table.rows[1].cells[0]; _set_cell_fill(bc, "F9FDF9")
    bp = bc.paragraphs[0]; bp.paragraph_format.space_before = Pt(6); bp.paragraph_format.space_after = Pt(6)
    _normal_run(bp, summary_text or "(No summary provided)", size_pt=10)


# ═══════════════════════════════════════════════════════════════════════════════
# DOCX BUILDER
# ═══════════════════════════════════════════════════════════════════════════════

def build_docx(data: dict) -> bytes:
    doc = Document()
    for section in doc.sections:
        section.top_margin = section.bottom_margin = Inches(1)
        section.left_margin = section.right_margin = Inches(1)

    # ── Title ────────────────────────────────────────────────────────────────
    for text, size, color, italic in [
        ("KARM CHARGE",                                     22, BRAND_GREEN, False),
        ("EV Charger Operations & Diagnostic Report",       14, HEADER_BLUE, False),
        ("Site Visit Operations Report",                    10, GREY_TEXT,   True),
    ]:
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(text); r.bold = not italic; r.italic = italic
        r.font.size = Pt(size); r.font.color.rgb = color

    doc.add_paragraph()

    # ── Document Control ─────────────────────────────────────────────────────
    visit_str = data["visit_date"].strftime("%A – %d %B %Y") if hasattr(data["visit_date"], "strftime") else str(data["visit_date"])
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

    # ── Personnel ────────────────────────────────────────────────────────────
    _add_section_heading(doc, "Accompanying Personnel")
    people = [p.strip() for p in data.get("personnel", []) if str(p).strip()]
    if people:
        for person in people:
            bp = doc.add_paragraph(style="List Bullet")
            bp.paragraph_format.space_before = Pt(2); bp.paragraph_format.space_after = Pt(2)
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
    asset_rows = [
        [a.get("asset_id", ""), a.get("location", ""),
         a.get("physical_status", ""), a.get("dashboard_status", ""),
         a.get("severity", "")]
        for a in data.get("asset_status", [])
        if any(str(v).strip() for v in a.values())
    ]
    if asset_rows:
        _add_data_table(doc,
            ["Asset ID", "Location", "Physical Status", "Dashboard Status", "Severity"],
            asset_rows)
    else:
        doc.add_paragraph("No asset records provided.").runs[0].font.size = Pt(10)
    doc.add_paragraph()

    # ── 4. Diagnostics ───────────────────────────────────────────────────────
    _add_section_heading(doc, "4. Diagnostics & Actions Taken")
    diagnostics = [d for d in data.get("diagnostics", []) if any(str(v).strip() for v in d.values())]
    if not diagnostics:
        doc.add_paragraph("No diagnostic entries provided.").runs[0].font.size = Pt(10)
    else:
        for idx, d in enumerate(diagnostics, start=1):
            ih = doc.add_paragraph()
            ih.paragraph_format.space_before = Pt(8); ih.paragraph_format.space_after = Pt(2)
            _bold_run(ih, f"Issue {idx}: {d.get('issue', '—')}", size_pt=11, color=HEADER_BLUE)
            for label, key in [
                ("Affected Units", "affected_units"),
                ("Symptom",        "symptom"),
                ("Root Cause",     "root_cause"),
                ("Action Taken",   "action_taken"),
                ("Status",         "status"),
            ]:
                value = str(d.get(key, "")).strip()
                if value:
                    bp = doc.add_paragraph(style="List Bullet")
                    bp.paragraph_format.space_before = Pt(1); bp.paragraph_format.space_after = Pt(1)
                    _bold_run(bp, f"{label}: ", size_pt=10)
                    _normal_run(bp, value, size_pt=10)
    doc.add_paragraph()

    # ── 5. Open Action Items ─────────────────────────────────────────────────
    _add_section_heading(doc, "5. Open Action Items")
    action_rows = []
    for i, a in enumerate(data.get("open_action_items", []), start=1):
        if any(str(v).strip() for v in a.values()):
            action_rows.append([str(len(action_rows) + 1),
                a.get("issue", ""), a.get("next_step", ""),
                a.get("owner", ""), a.get("priority", ""), a.get("target_date", "")])
    if action_rows:
        _add_data_table(doc,
            ["#", "Issue", "Next Step", "Owner", "Priority", "Target Date"],
            action_rows)
    else:
        doc.add_paragraph("No open action items logged.").runs[0].font.size = Pt(10)

    # ── Footer ───────────────────────────────────────────────────────────────
    fp = doc.add_paragraph(); fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.paragraph_format.space_before = Pt(24)
    fr = fp.add_run(f"KARM CHARGE  ·  {data['report_id']}  ·  Confidential – Internal Use Only")
    fr.italic = True; fr.font.size = Pt(8); fr.font.color.rgb = GREY_TEXT

    buf = io.BytesIO(); doc.save(buf); return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════════
# SESSION-STATE INITIALISATION
# ═══════════════════════════════════════════════════════════════════════════════

def init_state():
    """Initialise everything once."""
    defaults = {
        "form_version":      0,      # bumped on every JSON load → forces widget reset
        "site":              "",
        "visit_date":        date.today(),
        "report_id":         "",
        "prepared_by":       "Mohamed Othman",
        "role":              "Customer Ops Specialist",
        "personnel":         [""],
        "executive_summary": "",
        "kpi_total":         0,
        "kpi_operational":   0,
        "kpi_offline":       0,
        "kpi_open_actions":  0,
        "asset_rows":        [dict(EMPTY_ASSET)],
        "diag_rows":         [dict(EMPTY_DIAG)],
        "action_rows":       [dict(EMPTY_ACTION)],
        "json_text":         "",
        "docx_bytes":        None,
        "docx_filename":     None,
        "load_warnings":     [],
        "show_json_panel":   True,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def load_from_json(raw_json: str) -> bool:
    """Parse JSON and pour the data into every form field. Returns True on success."""
    parsed, warnings = parse_json(raw_json)
    if parsed is None:
        st.session_state["load_warnings"] = warnings
        return False

    # Bump version so all widgets are recreated with fresh `value=` parameters
    st.session_state["form_version"]      += 1
    st.session_state["site"]               = parsed["site"]
    st.session_state["visit_date"]         = parsed["visit_date"]
    st.session_state["report_id"]          = parsed["report_id"]
    st.session_state["prepared_by"]        = parsed["prepared_by"]
    st.session_state["role"]               = parsed["role"]
    st.session_state["personnel"]          = parsed["personnel"]
    st.session_state["executive_summary"]  = parsed["executive_summary"]
    st.session_state["kpi_total"]          = parsed["kpis"]["total_inspected"]
    st.session_state["kpi_operational"]    = parsed["kpis"]["operational"]
    st.session_state["kpi_offline"]        = parsed["kpis"]["offline"]
    st.session_state["kpi_open_actions"]   = parsed["kpis"]["open_actions"]
    st.session_state["asset_rows"]         = parsed["asset_status"]
    st.session_state["diag_rows"]          = parsed["diagnostics"]
    st.session_state["action_rows"]        = parsed["open_action_items"]
    st.session_state["load_warnings"]      = warnings
    st.session_state["docx_bytes"]         = None  # invalidate any old build
    return True


def collect_form_data() -> dict:
    """Gather current values from session_state into the shape build_docx expects."""
    return {
        "site":              st.session_state["site"],
        "visit_date":        st.session_state["visit_date"],
        "report_id":         st.session_state["report_id"] or
                             f"OPS-GEN-{st.session_state['visit_date'].isoformat()}",
        "prepared_by":       st.session_state["prepared_by"],
        "role":              st.session_state["role"],
        "personnel":         st.session_state["personnel"],
        "executive_summary": st.session_state["executive_summary"],
        "kpis": {
            "total_inspected": st.session_state["kpi_total"],
            "operational":     st.session_state["kpi_operational"],
            "offline":         st.session_state["kpi_offline"],
            "open_actions":    st.session_state["kpi_open_actions"],
        },
        "asset_status":      st.session_state["asset_rows"],
        "diagnostics":       st.session_state["diag_rows"],
        "open_action_items": st.session_state["action_rows"],
    }


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
    init_state()
    v = st.session_state["form_version"]   # versioned widget keys force a clean reset on JSON load

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

    # ═════════════════════════════════════════════════════════════════════════
    # STEP 1 — AI Prompt (collapsible)
    # ═════════════════════════════════════════════════════════════════════════
    with st.expander("🤖 **Step 1** — Need the AI prompt? Click to copy", expanded=False):
        st.caption("Copy this prompt, paste it into ChatGPT/Claude/Gemini along with your raw field notes, and the AI will return JSON.")
        st.code(AI_PROMPT, language="markdown")

    # ═════════════════════════════════════════════════════════════════════════
    # STEP 2 — Paste JSON to auto-fill form
    # ═════════════════════════════════════════════════════════════════════════
    with st.expander("📥 **Step 2** — Paste JSON to auto-fill the form below",
                     expanded=st.session_state["show_json_panel"]):
        st.caption("Paste the JSON from your AI. Click **Load JSON** — every field below gets pre-filled. You can then edit anything before generating.")

        col_a, col_b = st.columns([1, 1])
        with col_a:
            if st.button("📝 Load sample JSON", use_container_width=True,
                         help="Fills the box with a worked example"):
                st.session_state["json_text"] = json.dumps(SAMPLE_JSON, indent=2, ensure_ascii=False)
                st.rerun()
        with col_b:
            if st.button("🗑️ Clear JSON box", use_container_width=True):
                st.session_state["json_text"] = ""
                st.session_state["load_warnings"] = []
                st.rerun()

        json_text = st.text_area(
            "JSON",
            value=st.session_state["json_text"],
            height=280,
            placeholder='{\n  "site": "...",\n  "visit_date": "2026-05-17",\n  ...\n}',
            label_visibility="collapsed",
            key=f"json_text_widget_v{v}",
        )
        st.session_state["json_text"] = json_text

        if st.button("⬇️ **Load JSON into form below**", type="primary",
                     use_container_width=True):
            if load_from_json(st.session_state["json_text"]):
                st.session_state["show_json_panel"] = False
                st.success("✅ Form populated from JSON. Scroll down to review and edit.")
                st.rerun()
            else:
                for w in st.session_state["load_warnings"]:
                    st.error(w)

        # Show any non-fatal warnings from the last load
        for w in st.session_state["load_warnings"]:
            if w.startswith("⚠️") or w.startswith("ℹ️"):
                st.warning(w)

    st.divider()

    # ═════════════════════════════════════════════════════════════════════════
    # STEP 3 — Editable form
    # ═════════════════════════════════════════════════════════════════════════
    st.markdown("### ✏️ **Step 3** — Review & edit before generating")

    # ── 📍 Visit Details ─────────────────────────────────────────────────────
    st.subheader("📍 Visit Details")
    c1, c2 = st.columns(2)
    with c1:
        st.session_state["site"] = st.text_input(
            "Site / District",
            value=st.session_state["site"],
            key=f"site_v{v}",
            placeholder="e.g. Arkan — Crown Plaza",
        )
    with c2:
        st.session_state["visit_date"] = st.date_input(
            "Date of Visit",
            value=st.session_state["visit_date"],
            key=f"visit_date_v{v}",
        )

    c3, c4, c5 = st.columns(3)
    with c3:
        st.session_state["report_id"] = st.text_input(
            "Report ID (leave blank to auto-generate)",
            value=st.session_state["report_id"],
            key=f"report_id_v{v}",
        )
    with c4:
        st.session_state["prepared_by"] = st.text_input(
            "Prepared By",
            value=st.session_state["prepared_by"],
            key=f"prep_v{v}",
        )
    with c5:
        st.session_state["role"] = st.text_input(
            "Role",
            value=st.session_state["role"],
            key=f"role_v{v}",
        )

    st.divider()

    # ── 👥 Personnel ─────────────────────────────────────────────────────────
    st.subheader("👥 Accompanying Personnel")
    st.caption("One name per row.")
    personnel = st.session_state["personnel"]
    for i in range(len(personnel)):
        col_p, col_x = st.columns([10, 1])
        with col_p:
            personnel[i] = st.text_input(
                f"Person {i+1}",
                value=personnel[i],
                key=f"pers_{i}_v{v}",
                label_visibility="collapsed",
                placeholder="e.g. Eng. Ahmed Nour — Solargy, Engineer",
            )
        with col_x:
            if st.button("✖", key=f"pers_del_{i}_v{v}", help="Remove this person"):
                if len(personnel) > 1:
                    personnel.pop(i)
                else:
                    personnel[0] = ""
                st.rerun()

    if st.button("➕ Add Person", key=f"add_pers_v{v}"):
        personnel.append("")
        st.rerun()

    st.divider()

    # ── 📝 Executive Summary ─────────────────────────────────────────────────
    st.subheader("📝 Executive Summary")
    st.caption("This goes into the AT A GLANCE box in the report. Edit freely.")
    st.session_state["executive_summary"] = st.text_area(
        "Executive Summary",
        value=st.session_state["executive_summary"],
        height=180,
        key=f"summary_v{v}",
        label_visibility="collapsed",
        placeholder="Write or paste a 3-6 sentence professional summary of the visit.",
    )

    st.divider()

    # ── 📊 KPIs ──────────────────────────────────────────────────────────────
    st.subheader("📊 Visit KPIs")
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.session_state["kpi_total"] = st.number_input(
            "Total Inspected", min_value=0, step=1,
            value=st.session_state["kpi_total"], key=f"kpi1_v{v}")
    with k2:
        st.session_state["kpi_operational"] = st.number_input(
            "Operational", min_value=0, step=1,
            value=st.session_state["kpi_operational"], key=f"kpi2_v{v}")
    with k3:
        st.session_state["kpi_offline"] = st.number_input(
            "Offline / Faulted", min_value=0, step=1,
            value=st.session_state["kpi_offline"], key=f"kpi3_v{v}")
    with k4:
        st.session_state["kpi_open_actions"] = st.number_input(
            "Open Actions", min_value=0, step=1,
            value=st.session_state["kpi_open_actions"], key=f"kpi4_v{v}")

    st.divider()

    # ── 🔌 Asset Status ──────────────────────────────────────────────────────
    st.subheader("🔌 Asset Status & Observations")
    asset_rows = st.session_state["asset_rows"]

    for i, row in enumerate(asset_rows):
        with st.container(border=True):
            top_l, top_r = st.columns([10, 1])
            with top_l:
                st.markdown(f"**Asset row {i+1}**")
            with top_r:
                if st.button("✖", key=f"asset_del_{i}_v{v}", help="Remove this row"):
                    if len(asset_rows) > 1:
                        asset_rows.pop(i)
                    else:
                        asset_rows[0] = dict(EMPTY_ASSET)
                    st.rerun()

            c1, c2 = st.columns(2)
            row["asset_id"] = c1.text_input(
                "Asset ID", value=row.get("asset_id", ""),
                key=f"a_id_{i}_v{v}", placeholder="AR-01-EVAC-01-F10")
            row["location"] = c2.text_input(
                "Location", value=row.get("location", ""),
                key=f"a_loc_{i}_v{v}", placeholder="Arkan – Crown Plaza")

            c3, c4, c5 = st.columns(3)
            row["physical_status"] = c3.selectbox(
                "Physical Status", options=PHYS_STATUS_OPTS,
                index=PHYS_STATUS_OPTS.index(row["physical_status"])
                      if row.get("physical_status") in PHYS_STATUS_OPTS else 2,
                key=f"a_ps_{i}_v{v}")
            row["dashboard_status"] = c4.selectbox(
                "Dashboard Status", options=DASH_STATUS_OPTS,
                index=DASH_STATUS_OPTS.index(row["dashboard_status"])
                      if row.get("dashboard_status") in DASH_STATUS_OPTS else 1,
                key=f"a_ds_{i}_v{v}")
            row["severity"] = c5.selectbox(
                "Severity", options=SEVERITY_OPTIONS,
                index=SEVERITY_OPTIONS.index(row["severity"])
                      if row.get("severity") in SEVERITY_OPTIONS else 0,
                key=f"a_sev_{i}_v{v}")

    if st.button("➕ Add Asset Row", key=f"add_asset_v{v}"):
        asset_rows.append(dict(EMPTY_ASSET))
        st.rerun()

    st.divider()

    # ── 🔧 Diagnostics ───────────────────────────────────────────────────────
    st.subheader("🔧 Diagnostics & Actions Taken")
    diag_rows = st.session_state["diag_rows"]

    for i, row in enumerate(diag_rows):
        title = row.get("issue") or f"(untitled — issue {i+1})"
        with st.expander(f"🔧 Issue {i+1}: {title}", expanded=(i == 0)):
            row["issue"] = st.text_input(
                "Issue Title", value=row.get("issue", ""),
                key=f"d_issue_{i}_v{v}",
                placeholder="e.g. Charger Fault — Continuous Indicator Light Cycling")
            row["affected_units"] = st.text_input(
                "Affected Units", value=row.get("affected_units", ""),
                key=f"d_units_{i}_v{v}",
                placeholder="e.g. AR-01-EVAC-01-F10 (Y13)")
            row["symptom"] = st.text_area(
                "Symptom", value=row.get("symptom", ""),
                key=f"d_sym_{i}_v{v}", height=70)
            row["root_cause"] = st.text_area(
                "Root Cause", value=row.get("root_cause", ""),
                key=f"d_rc_{i}_v{v}", height=70)
            row["action_taken"] = st.text_area(
                "Action Taken", value=row.get("action_taken", ""),
                key=f"d_at_{i}_v{v}", height=70)
            row["status"] = st.selectbox(
                "Status", options=DIAG_STATUS_OPTS,
                index=DIAG_STATUS_OPTS.index(row["status"])
                      if row.get("status") in DIAG_STATUS_OPTS else 0,
                key=f"d_stat_{i}_v{v}")

            if st.button("🗑️ Remove this issue", key=f"d_del_{i}_v{v}"):
                if len(diag_rows) > 1:
                    diag_rows.pop(i)
                else:
                    diag_rows[0] = dict(EMPTY_DIAG)
                st.rerun()

    if st.button("➕ Add Diagnostic", key=f"add_diag_v{v}"):
        diag_rows.append(dict(EMPTY_DIAG))
        st.rerun()

    st.divider()

    # ── 📌 Open Action Items ─────────────────────────────────────────────────
    st.subheader("📌 Open Action Items")
    action_rows = st.session_state["action_rows"]

    for i, row in enumerate(action_rows):
        with st.container(border=True):
            top_l, top_r = st.columns([10, 1])
            with top_l:
                st.markdown(f"**Action item {i+1}**")
            with top_r:
                if st.button("✖", key=f"act_del_{i}_v{v}", help="Remove this row"):
                    if len(action_rows) > 1:
                        action_rows.pop(i)
                    else:
                        action_rows[0] = dict(EMPTY_ACTION)
                    st.rerun()

            c1, c2 = st.columns(2)
            row["issue"] = c1.text_input(
                "Issue", value=row.get("issue", ""), key=f"ac_issue_{i}_v{v}")
            row["owner"] = c2.text_input(
                "Owner", value=row.get("owner", ""), key=f"ac_own_{i}_v{v}",
                placeholder="e.g. Solargy")
            row["next_step"] = st.text_area(
                "Next Step", value=row.get("next_step", ""),
                key=f"ac_ns_{i}_v{v}", height=60)
            c3, c4 = st.columns(2)
            row["priority"] = c3.selectbox(
                "Priority", options=PRIORITY_OPTIONS,
                index=PRIORITY_OPTIONS.index(row["priority"])
                      if row.get("priority") in PRIORITY_OPTIONS else 2,
                key=f"ac_pri_{i}_v{v}")
            row["target_date"] = c4.text_input(
                "Target Date", value=row.get("target_date", ""),
                key=f"ac_td_{i}_v{v}", placeholder="e.g. 30 May 2026")

    if st.button("➕ Add Action Item", key=f"add_act_v{v}"):
        action_rows.append(dict(EMPTY_ACTION))
        st.rerun()

    st.divider()

    # ═════════════════════════════════════════════════════════════════════════
    # STEP 4 — Generate & Download
    # ═════════════════════════════════════════════════════════════════════════
    st.markdown("### 🚀 **Step 4** — Generate the report")

    col_gen, col_reset = st.columns([3, 1])
    with col_gen:
        if st.button("🚀 Generate Report", type="primary", use_container_width=True):
            try:
                data = collect_form_data()
                docx_bytes = build_docx(data)
                st.session_state["docx_bytes"]    = docx_bytes
                st.session_state["docx_filename"] = f"{data['report_id']}.docx"
                st.success("✅ Report generated! Download button below.")
            except Exception as exc:  # noqa: BLE001
                st.exception(exc)
    with col_reset:
        if st.button("🔄 Reset all fields", use_container_width=True):
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.rerun()

    if st.session_state.get("docx_bytes"):
        st.download_button(
            label    = f"⬇️ Download  {st.session_state['docx_filename']}",
            data     = st.session_state["docx_bytes"],
            file_name= st.session_state["docx_filename"],
            mime     = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            type="primary",
        )


if __name__ == "__main__":
    main()
