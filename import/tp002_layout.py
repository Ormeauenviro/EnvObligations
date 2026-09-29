"""Layout knowledge for the DTI-HSEQ-TP002.1 Environmental Compliance Obligations Register.

Shared by the Phase 1 audit (``audit_workbook.py``) and the Phase 3 transform.
Everything that is specific to how the workbook is laid out lives here, so the
parsers stay generic.

Canonical field names used in the column maps (a sheet maps a subset of these):

    location        Project Location / Activity
    section         Section / group (Management Plan/Subplan, Compliance Matter)
    stage           Project Stage
    condition_no    Condition #
    requirement     Compliance Requirement (and sheet-specific variants)
    legislation     Relevant Legislation / Standard or Code (PFAS only)
    practical_impl  Practical Implementation (PFAS only)
    source_ref      Source / Reference
    action          Compliance Action / Compliance Controls
    timing          Timing / Frequency, Timing/ Trigger
    responsibility  Responsibility
    risk            Risk Rating (A-D)
    stakeholder     Stakeholder Review/ Approval/ Submission
    due             Due Date (free text or a real date)
    compliance_check  Compliance Check (PFAS only; how compliance is verified)
    status          Compliance Status
    notes           Compliance Notes
    last_reviewed   Compliance last reviewed
    evidence        Evidence / Supporting Documents (mostly a description of the
                    evidence expected, occasionally a hyperlink)
    link1, link2, link3   "Links" columns (hyperlinks to SharePoint folders/files)
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass, field
from typing import Any
from xml.etree import ElementTree as ET

# The workbook was saved from this SharePoint library. Relative hyperlinks in the
# DA sheets ("../../../../:f:/s/DTI-20151005_QTMP/...") resolve against it.
SHAREPOINT_HOST = "https://dtinfrastructurecomau.sharepoint.com"

HEADER_ROW = 5
GROUP_HEADER_ROW = 4
FIRST_DATA_ROW = 6

# Column map for the standard DTI template layout (_Template, External CHMP,
# RPP Exemption, Specification Requirements, Legislation).
TEMPLATE_COLS = {
    "A": "location", "B": "stage", "C": "requirement", "D": "source_ref",
    "E": "action", "F": "responsibility", "G": "risk", "H": "stakeholder",
    "I": "due", "J": "status", "K": "notes", "L": "evidence",
}

# The DA layout used by No Name DA Conditions and the two current decision
# notice blocks on Marine Plants DA Conditions.
DA_CONDITION_COLS = {
    "A": "location", "B": "stage", "C": "condition_no", "D": "requirement",
    "E": "source_ref", "F": "action", "G": "timing", "H": "responsibility",
    "I": "due", "J": "status", "K": "notes", "L": "evidence",
    "M": "link1", "N": "link2",
}


@dataclass
class Block:
    """A contiguous run of rows sharing one header row."""

    header_row: int
    first_row: int
    last_row: int
    columns: dict[str, str]
    # Extra context applied to every row in the block (e.g. permit number).
    context: dict[str, str] = field(default_factory=dict)
    hidden: bool = False
    note: str = ""


@dataclass
class SheetSpec:
    name: str
    kind: str  # "obligations" | "reference" | "empty"
    id_prefix: str = ""
    blocks: list[Block] = field(default_factory=list)
    purpose: str = ""


SHEETS: list[SheetSpec] = [
    SheetSpec("Instructions", "reference",
              purpose="How to use the template; cell colour key; Risk A-D list "
                      "(C30:C33) used as the Risk Rating drop-down source."),
    SheetSpec("_Template", "reference",
              purpose="Blank register layout (row 4 group headers, row 5 headers). "
                      "Reference only - defines the standard 12-column layout."),
    SheetSpec("Summary", "reference",
              purpose="Project details and COUNTIF status totals. A20:A25 is the "
                      "Compliance Status drop-down source for every sheet."),
    SheetSpec(
        "CEMP and General Requirements", "obligations", "CEMP",
        purpose="Monitoring, inspection and reporting obligations from the EMP, "
                "CEMP annexes, CMP and ISO 14001.",
        blocks=[Block(5, 6, 65, {
            "A": "section", "B": "stage", "C": "requirement", "D": "source_ref",
            "E": "action", "F": "timing", "G": "responsibility", "H": "due",
            "I": "status", "J": "notes", "K": "last_reviewed", "L": "evidence",
        })],
    ),
    SheetSpec(
        "PFAS CMP Requirements", "obligations", "PFAS",
        purpose="PFAS Compliance Management Plan (Rev C) Compliance Matters 1-3.",
        blocks=[Block(5, 6, 24, {
            "A": "location", "B": "stage", "C": "section", "D": "requirement",
            "E": "legislation", "F": "practical_impl", "G": "source_ref",
            "H": "action", "I": "responsibility", "J": "due",
            "K": "compliance_check", "L": "status", "M": "notes", "N": "evidence",
        })],
    ),
    SheetSpec(
        "ADR WWBW Conditions", "obligations", "ADR",
        purpose="Accepted development requirements - waterway barrier works.",
        blocks=[Block(5, 6, 12, {
            "A": "location", "B": "stage", "C": "requirement", "D": "source_ref",
            "E": "action", "F": "timing", "G": "responsibility", "H": "due",
            "I": "status", "J": "notes", "K": "evidence",
        })],
    ),
    SheetSpec(
        "Bridge Creek DA Conditions", "obligations", "BC-DA",
        purpose="SARA 2402-39300 SDA conditions (Bridge Creek diversion) plus "
                "Fish Salvage, Rehabilitation & Remediation, Fish Passage "
                "Monitoring and Aquatic Plant Monitoring plan commitments.",
        blocks=[Block(5, 6, 49, {
            "A": "location", "B": "stage", "C": "requirement", "D": "source_ref",
            "E": "action", "F": "timing", "G": "responsibility", "H": "due",
            "I": "status", "J": "notes", "K": "evidence",
            "L": "link1", "M": "link2", "N": "link3",
        }, context={"permit": "2402-39300 SDA"})],
    ),
    SheetSpec(
        "No Name DA Conditions", "obligations", "NN-DA",
        purpose="SARA 2506-46598 SDA conditions (No Name Creek waterway barrier works).",
        blocks=[Block(5, 6, 17, DA_CONDITION_COLS, context={"permit": "2506-46598 SDA"})],
    ),
    SheetSpec(
        "Marine Plants DA Conditions", "obligations", "MP",
        purpose="SARA marine plant approvals: 2308-36075 SDA (TNAR, temporary) and "
                "2309-36966 SDA (NAC, permanent), one block per decision notice.",
        blocks=[
            Block(5, 6, 9, {
                "A": "location", "B": "stage", "C": "requirement", "D": "source_ref",
                "E": "action", "F": "responsibility", "G": "due", "H": "status",
                "I": "notes", "J": "evidence",
            }, hidden=True,
                note="Legacy approval summary (rows 1-19 hidden). Superseded by the "
                     "per-decision-notice blocks below."),
            Block(22, 23, 30, DA_CONDITION_COLS,
                  context={"permit": "2308-36075 SDA", "permit_short": "TNAR"}),
            Block(33, 34, 45, DA_CONDITION_COLS,
                  context={"permit": "2309-36966 SDA", "permit_short": "NAC"}),
        ],
    ),
    SheetSpec("No Name Operational Works", "empty",
              purpose="Placeholder - no content (A1 only)."),
    SheetSpec(
        "External CHMP Conditions", "obligations", "CHMP",
        purpose="Jabree CHMP and Danggan Balun Partial Clearance Report requirements.",
        blocks=[Block(5, 6, 19, TEMPLATE_COLS)],
    ),
    SheetSpec(
        "RPP Exemption", "obligations", "RPP",
        purpose="Riverine protection permit exemption requirements (WSS/2013/726).",
        blocks=[Block(5, 6, 8, TEMPLATE_COLS)],
    ),
    SheetSpec(
        "Specification Requirements", "obligations", "SPEC",
        purpose="Subcontract / Schedule A5 / Schedule C1 / MD-13-320 / MD-15-3 requirements.",
        blocks=[Block(5, 6, 37, TEMPLATE_COLS)],
    ),
    SheetSpec("Reporting Matrix", "reference",
              purpose="Maps report content items (rows) to MONTHLY / POST "
                      "CONSTRUCTION reports (columns B-F) with an X."),
    SheetSpec(
        "Legislation", "obligations", "LEG",
        purpose="Legislative obligations (requirement + Act only; no other columns filled).",
        blocks=[Block(5, 6, 22, TEMPLATE_COLS)],
    ),
    SheetSpec("_Doc Ref", "reference",
              purpose="Source document register (sheet, document, reference, "
                      "version, date, comments). Red text = not yet provided."),
]

SHEETS_BY_NAME = {s.name: s for s in SHEETS}

# ---------------------------------------------------------------------------
# Cell cleaning
# ---------------------------------------------------------------------------

NBSP = "\u00a0"
# AI-assistant citation fragments pasted into Bridge Creek cells, e.g.
# "[2402-39300...A Decision | PDF]" or "[2402-39300...itions (1) | PDF]".
CITATION_RE = re.compile(r"\s*\[[^\[\]]*\.\.\.[^\[\]]*\|\s*PDF\]")


def clean_text(value: Any) -> Any:
    """Normalise whitespace in a cell value without changing its meaning.

    - non-breaking spaces become ordinary spaces
    - Windows/Mac line endings become ``\\n``
    - leading tabs and trailing whitespace on each line are removed
    - leading/trailing blank lines are removed
    - a value that is only whitespace becomes ``None``
    Non-string values are returned unchanged.
    """
    if not isinstance(value, str):
        return value
    text = value.replace(NBSP, " ").replace("\r\n", "\n").replace("\r", "\n")
    lines = [line.rstrip().lstrip("\t") for line in text.split("\n")]
    text = "\n".join(lines).strip()
    return text or None


def strip_citations(text: str | None) -> tuple[str | None, list[str]]:
    """Remove pasted citation fragments; return the cleaned text and what was removed."""
    if not text:
        return text, []
    found = [m.strip() for m in CITATION_RE.findall(text)]
    return (CITATION_RE.sub("", text).strip() or None), found


def resolve_hyperlink(target: str | None) -> str | None:
    """Turn a workbook hyperlink target into an absolute SharePoint URL.

    Relative targets such as ``../../../../:f:/s/DTI-20151005_QTMP/Ig...`` climb
    past the library root, so they resolve to the tenant host.
    """
    if not target:
        return None
    target = target.strip()
    if target.lower().startswith(("http://", "https://", "mailto:")):
        return target
    stripped = re.sub(r"^(\.\./)+", "", target)
    return f"{SHAREPOINT_HOST}/{stripped.lstrip('/')}"


def is_blank(value: Any) -> bool:
    return clean_text(value) is None


# Values that mean "information only" when found in the action/responsibility
# columns (or status on Marine Plants R35).
INFO_ONLY_MARKERS = {"note", "noted", "noted - see below"}

# Placeholders found in Links/Evidence columns that are not links.
LINK_PLACEHOLDERS = {
    "?", "..", "n/a - until completion", "n/a - no current records",
    "drawings?", "monthly reports?",
}


def is_info_only(row: dict[str, Any]) -> bool:
    for key in ("action", "responsibility", "status"):
        value = clean_text(row.get(key))
        if isinstance(value, str) and value.split("\n")[0].strip().lower() in INFO_ONLY_MARKERS:
            return True
    return False


# ---------------------------------------------------------------------------
# Row classification (shared by the audit and the transform)
# ---------------------------------------------------------------------------

def row_values(ws, r: int, columns: dict[str, str]) -> dict:
    return {fld: ws[f"{col}{r}"].value for col, fld in columns.items()}


def classify(ws, r: int, block: Block, header_texts: set[str]) -> tuple[str, str]:
    """Return (class, reason) for one row inside a block's range."""
    raw = row_values(ws, r, block.columns)
    cleaned = {k: clean_text(v) for k, v in raw.items()}
    filled = {k: v for k, v in cleaned.items() if v is not None}
    if not filled:
        has_nbsp = any(isinstance(v, str) and v.strip(" \u00a0") == "" and v for v in raw.values())
        return "blank", "only non-breaking spaces" if has_nbsp else ""
    if {str(v).strip() for v in filled.values()} <= header_texts:
        return "repeated-header", ""
    if block.hidden:
        if set(filled) == {"location"}:
            return "legacy-stub", f"only '{filled['location']}' in column A (hidden row)"
        return "legacy-hidden", "hidden legacy approval summary row"
    if is_info_only(cleaned):
        return "info-only", "Note/Noted in action, responsibility or status"
    if (filled.get("condition_no") and not filled.get("location") and not filled.get("stage")):
        return "sub-condition", f"condition {filled['condition_no']} continues the row above"
    req = filled.get("requirement")
    if req is None:
        return "flag", "no requirement text"
    if isinstance(req, str) and len(req) < 5:
        return "flag", f"requirement is a placeholder ({req!r})"
    return "obligation", ""


# ---------------------------------------------------------------------------
# Threaded comments (openpyxl only exposes the legacy fallback text, not the author)
# ---------------------------------------------------------------------------
_NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "tc": "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments",
}


def _rels(z: zipfile.ZipFile, path: str) -> dict[str, tuple[str, str]]:
    if path not in z.namelist():
        return {}
    root = ET.fromstring(z.read(path))
    return {r.get("Id"): (r.get("Type", ""), r.get("Target", "")) for r in root.findall("rel:Relationship", _NS)}


def load_threaded_comments(xlsx_path) -> dict[tuple[str, str], list[tuple[str, str]]]:
    """Return {(sheet name, cell ref): [(author, text), ...]} in thread order."""
    out: dict[tuple[str, str], list[tuple[str, str]]] = {}
    with zipfile.ZipFile(xlsx_path) as z:
        people = {}
        if "xl/persons/person.xml" in z.namelist():
            for p in ET.fromstring(z.read("xl/persons/person.xml")).findall("tc:person", _NS):
                people[p.get("id")] = " ".join(p.get("displayName", "").split())
        wb_rels = _rels(z, "xl/_rels/workbook.xml.rels")
        workbook = ET.fromstring(z.read("xl/workbook.xml"))
        for sheet in workbook.find("m:sheets", _NS):
            rid = sheet.get(f"{{{_NS['r']}}}id")
            target = wb_rels[rid][1]  # e.g. worksheets/sheet4.xml
            sheet_file = target.split("/")[-1]
            for rtype, rtarget in _rels(z, f"xl/worksheets/_rels/{sheet_file}.rels").values():
                if not rtype.endswith("/threadedComment"):
                    continue
                tc_path = "xl/" + rtarget.replace("../", "")
                for c in ET.fromstring(z.read(tc_path)).findall("tc:threadedComment", _NS):
                    text = " ".join((c.findtext("tc:text", "", _NS) or "").split())
                    out.setdefault((sheet.get("name"), c.get("ref")), []).append(
                        (people.get(c.get("personId"), "Unknown"), text))
    return out
