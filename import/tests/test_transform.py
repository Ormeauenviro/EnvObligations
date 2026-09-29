"""Tests for import/transform.py. Run with: python -m pytest import"""

from __future__ import annotations

import datetime as dt
import json
import re
import sys
from pathlib import Path

import pytest

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
import tp002_layout as L  # noqa: E402
import transform as T  # noqa: E402

IMPORT_DATE = dt.date(2026, 9, 29)


@pytest.fixture(scope="module")
def result(tmp_path_factory):
    out = tmp_path_factory.mktemp("output")
    t = T.run(T.DEFAULT_XLSX, out, IMPORT_DATE)
    return t, out


@pytest.fixture(scope="module")
def obs(result):
    return {o["Title"]: o for o in result[0].lists["Obligations"]}


def items(result, key):
    return result[0].lists[key]


# ---------------------------------------------------------------------------
# Counts
# ---------------------------------------------------------------------------
def test_list_counts(result):
    counts = {k: len(v) for k, v in result[0].lists.items()}
    assert counts == {
        "SourceDocuments": 27, "ProjectMilestones": 6, "RoleAssignments": 10, "Obligations": 225,
        "Occurrences": 21, "Tasks": 4, "Evidence": 71, "FollowUps": 5, "StatusHistory": 246,
        "ReportComponents": 26,
    }


def test_obligations_per_register(result):
    from collections import Counter
    c = Counter(o["SourceRegister"] for o in items(result, "Obligations"))
    assert c == {
        "CEMP and General Requirements": 61,  # 60 rows + row 11 split (D-05)
        "PFAS CMP Requirements": 19, "ADR WWBW Conditions": 7, "Bridge Creek DA Conditions": 42,
        "No Name DA Conditions": 11,          # row 7 becomes Tasks (D-06)
        "Marine Plants DA Conditions": 20, "External CHMP Conditions": 14, "RPP Exemption": 3,
        "Specification Requirements": 32, "Legislation": 16,
    }


# ---------------------------------------------------------------------------
# Marine Plants (the hardest sheet)
# ---------------------------------------------------------------------------
def test_marine_plants_ids_follow_decision_notice_blocks(obs):
    mp = sorted(k for k in obs if k.startswith("MP-"))
    assert mp == [
        "MP-NAC-001", "MP-NAC-002", "MP-NAC-002A", "MP-NAC-002B", "MP-NAC-002C", "MP-NAC-003", "MP-NAC-004",
        "MP-NAC-005", "MP-NAC-006", "MP-NAC-006B", "MP-NAC-007", "MP-NAC-007B",
        "MP-TNAR-001", "MP-TNAR-002", "MP-TNAR-003", "MP-TNAR-004", "MP-TNAR-005", "MP-TNAR-005B",
        "MP-TNAR-006", "MP-TNAR-006B",
    ]


def test_marine_plants_blocks_link_to_their_permit(obs):
    for k, o in obs.items():
        if k.startswith("MP-TNAR"):
            assert o["SourceDocument"]["key"] == "SDA-2308-36075"
            assert "2308-36075 SDA" in o["ExtraFields"]
        elif k.startswith("MP-NAC"):
            assert o["SourceDocument"]["key"] == "SDA-2309-36966"


def test_marine_plants_hidden_legacy_rows_not_imported(obs):
    rows = [int(re.search(r"!A(\d+)", o["SourceCell"]).group(1)) for o in obs.values()
            if o["SourceRegister"] == "Marine Plants DA Conditions"]
    assert min(rows) == 23 and not set(rows) & {6, 7, 8, 9, 21, 22, 32, 33}


def test_marine_plants_row23_realigned(obs):
    o = obs["MP-TNAR-001"]
    assert o["Status"] == "Underway"                 # from K "Not completed - ongoing"
    assert o["ComplianceNotes"] == "Compare to GIS as evidence"  # from L
    assert o["EvidenceRequired"] is None
    assert "Unmapped (J): Ongoing" in o["ExtraFields"]


def test_marine_plants_same_project_and_sub_conditions_inherit(obs):
    tnar_loc = obs["MP-TNAR-001"]["LocationActivity"]
    assert tnar_loc.startswith("46 Greyhound Road")
    for k in ("MP-TNAR-002", "MP-TNAR-005", "MP-TNAR-005B", "MP-TNAR-006B"):
        assert obs[k]["LocationActivity"] == tnar_loc, k
    nac_loc = obs["MP-NAC-001"]["LocationActivity"]
    assert nac_loc.startswith("Pacific Highway, 56 Prairie Road")
    for k in ("MP-NAC-002A", "MP-NAC-007B"):
        assert obs[k]["LocationActivity"] == nac_loc
        assert obs[k]["ProjectStage"] == "Construction"
    assert obs["MP-NAC-002A"]["ParentObligationID"] == "MP-NAC-002"
    assert obs["MP-NAC-002A"]["ConditionNo"] == "2a"
    assert obs["MP-TNAR-005B"]["ResponsibleRole"] == "Environmental Manager"


def test_marine_plants_noted_parent_is_info_only(obs):
    o = obs["MP-NAC-002"]
    assert o["InfoOnly"] and o["Status"] == T.NA_INFO and not o["Active"] and o["Owner"] is None


def test_marine_plants_notifications_split(result, obs):
    occ = {o["Title"]: o for o in items(result, "Occurrences")}
    for oid in ("MP-TNAR-002", "MP-NAC-003", "BC-DA-005"):
        assert occ[f"{oid}|COMMENCE"]["Status"] == "Complete"
        assert occ[f"{oid}|COMMENCE"]["EvidenceURL"]["url"].startswith("https://dtinfrastructurecomau.sharepoint.com/")
        assert occ[f"{oid}|COMPLETE"]["Status"] == "Not started"
        assert obs[oid]["Status"] == "Underway"           # rolled up (D-05)
        assert obs[oid]["CurrentOccurrenceKey"] == f"{oid}|COMPLETE"
        ev = [e for e in items(result, "Evidence") if e["ObligationID"] == oid]
        assert ev and ev[0]["OccurrenceKey"] == f"{oid}|COMMENCE"


def test_marine_plants_daf_quote_moved_to_notes(obs):
    for k in ("MP-TNAR-004", "MP-NAC-005"):
        assert "DAF correspondence: The Department believes" in obs[k]["ComplianceNotes"]
        assert obs[k]["EvidenceRequired"] is None


def test_marine_plants_source_documents(result):
    docs = {d["DocKey"]: d for d in items(result, "SourceDocuments")}
    nac = docs["SDA-2309-36966"]
    assert nac["DocLink"]["description"].startswith("2309-36966 SDA - Decision")
    assert nac["FolderLink"]["url"].startswith("https://")
    assert "2309-36996" in nac["Comments"]           # typo recorded (D-08)
    assert "hidden row 6" in docs["SDA-2308-36075"]["Comments"]


# ---------------------------------------------------------------------------
# Other sheets
# ---------------------------------------------------------------------------
def test_no_name_conditions_and_tasks(result, obs):
    nn = sorted(k for k in obs if k.startswith("NN-DA"))
    assert nn == [f"NN-DA-{i:03d}" for i in range(1, 12)]
    assert [obs[k]["ConditionNo"] for k in nn] == [str(i) for i in range(1, 12)]
    tasks = items(result, "Tasks")
    assert [t["Title"][:3] for t in tasks] == ["(a)", "(b)", "(c)", "(d)"]
    assert [t["Done"] for t in tasks] == [True, False, False, False]
    assert all(t["ObligationID"] == "NN-DA-001" for t in tasks)
    o = obs["NN-DA-010"]
    assert o["DueDate"] == "2026-07-31" and o["Status"] == "Complete"


def test_bridge_creek(obs):
    bc = [obs[f"BC-DA-{i:03d}"] for i in range(1, 15)]
    assert [o["ConditionNo"] for o in bc] == [str(i) for i in range(1, 15)]
    for o in obs.values():
        assert "...A Decision | PDF]" not in (o["Requirement"] or "")
        assert "| PDF]" not in (o["TimingDetail"] or "")
    assert "Citations" in obs["BC-DA-001"]["ExtraFields"]
    assert obs["BC-DA-001"]["SourceReference"].startswith("Condition 1, SARA Ref 2402-39300 SDA")  # hidden column D
    assert obs["BC-DA-001"]["Status"] == "Underway"
    assert obs["BC-DA-003"]["Status"] == "Not started"      # D-01b
    assert obs["BC-DA-008"]["Status"] == "Not started"
    assert obs["BC-DA-020"]["Status"] == "Not started"      # NBSP-only status
    assert "DAF correspondence" in obs["BC-DA-007"]["ComplianceNotes"]  # unheaded column N
    assert obs["BC-DA-017"]["SourceDocument"]["key"] == "PLAN-FISH-SALVAGE"


def test_cemp_split_and_report_components(result, obs):
    assert obs["CEMP-006"]["Frequency"] == "Weekly"
    assert obs["CEMP-006B"]["Frequency"] == "Monthly"
    assert obs["CEMP-006B"]["ComplianceAction"] == "CMP Monthly Audit Report"
    rc = {c["ReportName"]: c["ProducingObligationID"] for c in items(result, "ReportComponents")}
    assert rc == {"Monitoring data made available to TMR": "CEMP-005", "PFAS Monthly Compliance Report": "CEMP-004",
                  "CMP Monthly Audit Report": "CEMP-006B", "Compliance Register": "CEMP-039",
                  "CMP Close Out Report": "CEMP-003"}


def test_pfas_monthly_report_due_on_5th(obs):
    assert obs["CEMP-004"]["Frequency"] == "Monthly" and obs["CEMP-004"]["DueOffsetDays"] == 5


def test_roles_and_owners(obs):
    assert obs["CEMP-001"]["Owner"] == "DanielBlunt@dtinfrastructure.com.au"
    assert obs["PFAS-001"]["Owner"] == "SamMoraes@dtinfrastructure.com.au"
    nn = obs["NN-DA-003"]                                   # "Contractor / DT Infrastructure"
    assert nn["ResponsibleRole"] == "Project Manager" and nn["SupportingRoles"] == "Contractor"
    cemp35 = obs["CEMP-030"]                                # "Fauna Spotter-Catcher" only
    assert cemp35["ResponsibleRole"] == "Environmental and Sustainability Team"
    assert obs["BC-DA-012"]["ResponsibleRole"] == "Asset Owner / Operator" and obs["BC-DA-012"]["Owner"] is None


def test_followups_set_flag(result, obs):
    fu = {f["ObligationID"] for f in items(result, "FollowUps")}
    assert fu == {"RPP-001", "LEG-016", "ADR-001", "CEMP-036", "CEMP-023"}
    assert all(obs[k]["HasOpenFollowUp"] for k in fu)
    assert obs["RPP-001"]["Requirement"].startswith("\u26a0 Requirement text missing")


def test_every_obligation_has_import_history(result):
    hist = {(h["Title"], h["ItemType"]) for h in items(result, "StatusHistory")}
    for o in items(result, "Obligations"):
        assert (o["Title"], "Obligation") in hist


def test_milestone_occurrences_undated_until_milestones_known(result):
    occ = [o for o in items(result, "Occurrences") if o["Title"].split("|")[1] in ("Y1", "Y2", "Y5")]
    assert len(occ) == 15 and all(o["PeriodDue"] is None for o in occ)


def test_milestone_dates_flow_through(tmp_path, monkeypatch):
    """Filling a milestone date dates the dependent items."""
    cfg = tmp_path / "config"
    cfg.mkdir()
    (cfg / "role-assignments.csv").write_text((T.CONFIG / "role-assignments.csv").read_text(encoding="utf-8"), encoding="utf-8")
    ms = (T.CONFIG / "project-milestones.csv").read_text(encoding="utf-8").replace(
        "Practical completion,,", "Practical completion,30/06/2027,")
    (cfg / "project-milestones.csv").write_text(ms, encoding="utf-8")
    monkeypatch.setattr(T, "CONFIG", cfg)
    t = T.run(T.DEFAULT_XLSX, tmp_path / "out", IMPORT_DATE)
    obs = {o["Title"]: o for o in t.lists["Obligations"]}
    occ = {o["Title"]: o for o in t.lists["Occurrences"]}
    assert obs["CEMP-058"]["DueDate"] == "2027-07-28"            # practical completion + 4 weeks
    assert occ["BC-DA-034|Y1"]["PeriodDue"] == "2028-06-29"      # + 365 days
    assert obs["BC-DA-034"]["DueDate"] == "2028-06-29"           # rolled up from current occurrence


# ---------------------------------------------------------------------------
# Nothing is lost: every non-empty source cell appears somewhere in the output
# ---------------------------------------------------------------------------
def _collapse(text: str) -> str:
    return " ".join(str(text).split()).lower()


def _strings(value) -> list[str]:
    if isinstance(value, dict):
        return [s for v in value.values() for s in _strings(v)]
    if isinstance(value, list):
        return [s for v in value for s in _strings(v)]
    return [str(value)] if value is not None else []


def test_nothing_lost(result):
    t = result[0]
    by_row: dict[tuple[str, int], list[dict]] = {}
    for o in t.lists["Obligations"]:
        m = re.search(r"'(.+)'!A(\d+)", o["SourceCell"])
        by_row.setdefault((m.group(1), int(m.group(2))), []).append(o)
    missing = []
    for spec in L.SHEETS:
        if spec.kind != "obligations":
            continue
        ws = t.wb[spec.name]
        for block in spec.blocks:
            for r in range(block.first_row, block.last_row + 1):
                obs = by_row.get((spec.name, r))
                if not obs:
                    continue
                ids = {o["Title"] for o in obs}
                related = list(obs)
                for key in ("Evidence", "Tasks", "Occurrences"):
                    related += [i for i in t.lists[key] if i.get("ObligationID") in ids]
                blob = _collapse(" ".join(_strings(related)))
                for col in block.columns:
                    value = ws[f"{col}{r}"].value
                    if isinstance(value, (dt.date, dt.datetime)):
                        needle = value.date().isoformat()
                    else:
                        cleaned = L.clean_text(str(value) if value is not None else None)
                        if cleaned is None:
                            continue
                        needle = _collapse(L.strip_citations(cleaned)[0] or cleaned)
                    if needle in blob:
                        continue
                    # Multi-line cells may be stored line by line (roles joined with " / ", split rows).
                    lines = [_collapse(x) for x in str(L.clean_text(str(value))).split("\n") if x.strip()]
                    if len(lines) > 1 and all(x in blob for x in lines):
                        continue
                    missing.append(f"{spec.name}!{col}{r}: {needle[:60]!r}")
    assert not missing, "\n".join(missing[:20])


def test_output_is_deterministic(tmp_path):
    a, b = tmp_path / "a", tmp_path / "b"
    T.run(T.DEFAULT_XLSX, a, IMPORT_DATE)
    T.run(T.DEFAULT_XLSX, b, IMPORT_DATE)
    for f in a.iterdir():
        assert f.read_bytes() == (b / f.name).read_bytes(), f.name


# ---------------------------------------------------------------------------
# Layout helpers
# ---------------------------------------------------------------------------
def test_resolve_hyperlink():
    assert L.resolve_hyperlink("../../../../:f:/s/DTI/abc?e=1") == "https://dtinfrastructurecomau.sharepoint.com/:f:/s/DTI/abc?e=1"
    assert L.resolve_hyperlink("https://x.sharepoint.com/a") == "https://x.sharepoint.com/a"
    assert L.resolve_hyperlink(None) is None


def test_clean_text():
    assert L.clean_text("\u00a0") is None
    assert L.clean_text("\tLine one  \r\nLine two ") == "Line one\nLine two"
    assert L.strip_citations("Text. [2402-39300...A Decision | PDF]") == ("Text.", ["[2402-39300...A Decision | PDF]"])
