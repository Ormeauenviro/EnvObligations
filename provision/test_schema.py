"""Static checks on provision/list-schema.json. Run with: python -m pytest provision"""

import json
import re
from pathlib import Path

import pytest

SCHEMA = json.loads((Path(__file__).parent / "list-schema.json").read_text(encoding="utf-8"))
LISTS = SCHEMA["lists"]
TYPES = {"Text", "Note", "Choice", "DateOnly", "DateTime", "Number", "Boolean", "User", "UserMulti", "URL", "Lookup"}
# Built-in fields a view may reference.
BUILTIN = {"LinkTitle", "Title", "ID", "Created", "Modified", "Author", "Editor"}
LOOKUP_LIKE = {"User", "UserMulti", "Lookup"}
MAX_INDEXES = 20          # SharePoint limit per list
MAX_LOOKUPS_PER_VIEW = 12  # list view lookup threshold
# Names SharePoint already uses on a generic list; reusing them would clash.
RESERVED = {"ID", "Title", "Created", "Modified", "Author", "Editor", "Attachments", "ContentType", "Order", "GUID", "FileRef"}


def choices(field):
    c = field.get("choices")
    return SCHEMA["choiceSets"][c[1:]] if isinstance(c, str) and c.startswith("@") else c


@pytest.mark.parametrize("lst", LISTS, ids=[l["key"] for l in LISTS])
def test_fields_are_valid(lst):
    names = [f["name"] for f in lst["fields"]]
    assert len(names) == len(set(names)), "duplicate internal names"
    for f in lst["fields"]:
        assert re.fullmatch(r"[A-Z][A-Za-z0-9]{1,31}", f["name"]), f"bad internal name {f['name']}"
        assert f["name"] not in RESERVED, f"reserved name {f['name']}"
        assert f["type"] in TYPES, f"{f['name']}: unknown type {f['type']}"
        if f.get("unique"):
            assert f.get("indexed"), f"{f['name']}: unique requires indexed"
        if f["type"] == "Choice":
            opts = choices(f)
            assert opts and len(opts) == len(set(opts)), f"{f['name']}: bad choices"
            if "default" in f:
                assert f["default"] in opts, f"{f['name']}: default not in choices"
        if f["type"] == "Lookup":
            keys = [l["key"] for l in LISTS]
            assert f["lookupList"] in keys, f"{f['name']}: unknown lookup list"
            assert keys.index(f["lookupList"]) < keys.index(lst["key"]), "lookup target must be created first"
        if f["type"] == "Boolean":
            assert isinstance(f.get("default"), bool), f"{f['name']}: Boolean needs a default"


@pytest.mark.parametrize("lst", LISTS, ids=[l["key"] for l in LISTS])
def test_index_limit(lst):
    # Lookup columns are always indexed by the provisioning script.
    indexed = [f for f in lst["fields"] if f.get("indexed") or f["type"] == "Lookup"]
    total = len(indexed) + (1 if lst["titleField"].get("indexed") else 0)
    assert total <= MAX_INDEXES, f"{lst['key']}: {total} indexes"


@pytest.mark.parametrize("lst", LISTS, ids=[l["key"] for l in LISTS])
def test_views(lst):
    by_name = {f["name"]: f for f in lst["fields"]}
    assert sum(1 for v in lst["views"] if v.get("default")) == 1, "exactly one default view"
    assert any(v["title"] == "All Items" for v in lst["views"])
    for v in lst["views"]:
        for fld in v["fields"]:
            assert fld in by_name or fld in BUILTIN, f"view {v['title']}: unknown field {fld}"
        lookups = sum(1 for fld in v["fields"] if fld in by_name and by_name[fld]["type"] in LOOKUP_LIKE)
        lookups += sum(1 for fld in v["fields"] if fld in {"Author", "Editor"})
        assert lookups <= MAX_LOOKUPS_PER_VIEW
        # CAML must be well-formed and only reference known fields.
        from xml.etree import ElementTree as ET
        root = ET.fromstring(f"<Query>{v['query']}</Query>")
        for ref in root.iter("FieldRef"):
            name = ref.get("Name")
            assert name in by_name or name in BUILTIN, f"view {v['title']}: CAML references {name}"


def test_list_urls_unique_and_safe():
    urls = [l["url"] for l in LISTS]
    assert len(urls) == len(set(urls))
    for u in urls:
        assert re.fullmatch(r"Lists/[A-Za-z]+", u), u


def test_children_carry_obligation_id_for_delegation():
    for key in ("Occurrences", "Tasks", "Evidence", "FollowUps", "StatusHistory"):
        lst = next(l for l in LISTS if l["key"] == key)
        f = next(f for f in lst["fields"] if f["name"] == "ObligationID")
        assert f["type"] == "Text" and f["indexed"], key


def test_filterable_obligation_columns_are_indexed():
    """Every column the Register screen filters on must be indexed (the Occurrences list will pass 5,000 over time)."""
    ob = {f["name"]: f for f in next(l for l in LISTS if l["key"] == "Obligations")["fields"]}
    for name in ("SourceRegister", "ProjectStage", "Status", "Owner", "RiskSort", "Frequency",
                 "IsOverdue", "HasOpenFollowUp", "RequirementSummary", "DueDate"):
        assert ob[name].get("indexed"), name
    occ = {f["name"]: f for f in next(l for l in LISTS if l["key"] == "Occurrences")["fields"]}
    for name in ("ObligationID", "PeriodDue", "Status", "Assignee", "IsOverdue"):
        assert occ[name].get("indexed"), name


def test_generated_docs_are_current():
    """docs/list-schema.md and docs/manual-provisioning.md must match the schema (run build_schema_docs.py)."""
    import importlib.util
    spec = importlib.util.spec_from_file_location("build_schema_docs", Path(__file__).parent / "build_schema_docs.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    assert mod.main(check=True) == 0
