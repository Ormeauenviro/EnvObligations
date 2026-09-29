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
from dataclasses import dataclass, field
from typing import Any

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
