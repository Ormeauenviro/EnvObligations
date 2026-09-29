"""Render the schema-driven sections of docs/list-schema.md and docs/manual-provisioning.md.

The hand-written parts of both documents stay as they are; only the text between
``<!-- BEGIN GENERATED: name -->`` and ``<!-- END GENERATED: name -->`` is replaced.

Usage:
    python provision/build_schema_docs.py          # rewrite the generated sections
    python provision/build_schema_docs.py --check  # exit 1 if they are out of date
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from xml.etree import ElementTree as ET

REPO = Path(__file__).resolve().parent.parent
SCHEMA = json.loads((REPO / "provision" / "list-schema.json").read_text(encoding="utf-8"))
LISTS = SCHEMA["lists"]
BUILTIN_NAMES = {"LinkTitle": "(Title, linked)", "ID": "ID", "Created": "Created", "Modified": "Modified"}

TYPE_LABEL = {
    "Text": "Single line of text",
    "Note": "Multiple lines of text (plain)",
    "Choice": "Choice (drop-down)",
    "DateOnly": "Date (date only)",
    "DateTime": "Date and time",
    "Number": "Number (0 decimals)",
    "Boolean": "Yes/No",
    "User": "Person (single)",
    "UserMulti": "Person (multiple)",
    "URL": "Hyperlink",
    "Lookup": "Lookup",
}

# Exact settings to pick in classic List settings > Create column.
TYPE_UI = {
    "Text": "Single line of text; Maximum number of characters: 255",
    "Note": "Multiple lines of text; Number of lines: 6; Type of text: Plain text; Append Changes to Existing Text: No",
    "Choice": "Choice; Display choices using: Drop-Down Menu; Allow 'Fill-in' choices: No",
    "DateOnly": "Date and Time; Date and Time Format: Date Only; Display Format: Standard",
    "DateTime": "Date and Time; Date and Time Format: Date & Time; Display Format: Standard",
    "Number": "Number; Number of decimal places: 0",
    "Boolean": "Yes/No (check box)",
    "User": "Person or Group; Allow multiple selections: No; Allow selection of: People Only; Choose from: All Users",
    "UserMulti": "Person or Group; Allow multiple selections: Yes; Allow selection of: People Only; Choose from: All Users",
    "URL": "Hyperlink or Picture; Format URL as: Hyperlink",
    "Lookup": "Lookup (information already on this site)",
}


def list_by_key(key: str) -> dict:
    return next(l for l in LISTS if l["key"] == key)


def choices(field: dict) -> list[str]:
    c = field.get("choices")
    return SCHEMA["choiceSets"][c[1:]] if isinstance(c, str) and c.startswith("@") else (c or [])


def choice_ref(field: dict) -> str:
    c = field.get("choices")
    return f"[{c[1:]}](#choice-{c[1:].lower()})" if isinstance(c, str) and c.startswith("@") else ", ".join(c or [])


def default_text(field: dict) -> str:
    if "default" not in field:
        return ""
    d = field["default"]
    return ("Yes" if d else "No") if isinstance(d, bool) else str(d)


def is_indexed(field: dict) -> bool:
    return bool(field.get("indexed")) or field["type"] == "Lookup"


def display_of(lst: dict, name: str) -> str:
    if name in ("LinkTitle", "Title"):
        return lst["titleField"]["displayName"]
    for f in lst["fields"]:
        if f["name"] == name:
            return f["displayName"]
    return BUILTIN_NAMES.get(name, name)


# ---------------------------------------------------------------------------
# CAML -> plain English (for the manual view steps)
# ---------------------------------------------------------------------------
OPS = {"Eq": "is equal to", "Neq": "is not equal to", "Geq": "is greater than or equal to",
       "Leq": "is less than or equal to", "Gt": "is greater than", "Lt": "is less than"}


def _value_text(value_el: ET.Element) -> str:
    today = value_el.find("Today")
    if today is not None:
        off = int(today.get("OffsetDays", "0"))
        return "[Today]" if off == 0 else f"[Today]{off:+d}"
    if value_el.find("UserID") is not None:
        return "[Me]"
    if value_el.get("Type") == "Boolean":
        return "Yes" if value_el.text == "1" else "No"
    return value_el.text or ""


def _flatten(node: ET.Element, lst: dict, out: list[tuple[str, str]], joiner: str = "") -> None:
    if node.tag in ("And", "Or"):
        kids = list(node)
        _flatten(kids[0], lst, out, joiner)
        for k in kids[1:]:
            _flatten(k, lst, out, node.tag.upper())
        return
    field = display_of(lst, node.find("FieldRef").get("Name"))
    out.append((joiner, f"**{field}** {OPS[node.tag]} `{_value_text(node.find('Value'))}`"))


def describe_view(lst: dict, view: dict) -> list[str]:
    root = ET.fromstring(f"<Q>{view['query']}</Q>")
    steps = []
    where = root.find("Where")
    if where is not None:
        conds: list[tuple[str, str]] = []
        _flatten(list(where)[0], lst, conds)
        parts = [c if not j else f"{j} {c}" for j, c in conds]
        steps.append("Filter: Show items only when " + " ".join(parts))
    order = root.find("OrderBy")
    if order is not None:
        sorts = [f"**{display_of(lst, r.get('Name'))}** {'descending' if r.get('Ascending') == 'FALSE' else 'ascending'}"
                 for r in order.findall("FieldRef")]
        steps.append("Sort: " + ", then ".join(sorts))
    group = root.find("GroupBy")
    if group is not None:
        collapsed = "collapsed" if group.get("Collapse") == "TRUE" else "expanded"
        steps.append(f"Group By: **{display_of(lst, group.find('FieldRef').get('Name'))}** ({collapsed})")
    return steps


# ---------------------------------------------------------------------------
# list-schema.md
# ---------------------------------------------------------------------------
def render_schema() -> str:
    out: list[str] = []
    w = out.append
    for lst in LISTS:
        idx_count = sum(1 for f in lst["fields"] if is_indexed(f)) + (1 if lst["titleField"].get("indexed") else 0)
        lookups = [f for f in lst["fields"] if f["type"] in ("Lookup", "User", "UserMulti")]
        w(f"### {lst['title']}")
        w("")
        w(f"`{lst['url']}` · {lst['description']}")
        w("")
        w(f"Versioning on · attachments {'on' if lst.get('attachments') else 'off'} · "
          f"{idx_count} indexed columns (limit 20) · {len(lookups)} lookup/person columns (+ Created By/Modified By)")
        w("")
        w("| Display name | Internal name | Type | Req. | Index | Default | Choices / notes |")
        w("|---|---|---|---|---|---|---|")
        t = lst["titleField"]
        tidx = "Unique" if t.get("unique") else ("Yes" if t.get("indexed") else "")
        w(f"| {t['displayName']} | `Title` | Single line of text | Yes | {tidx} |  | Built-in Title column, renamed |")
        for f in lst["fields"]:
            idx = "Unique" if f.get("unique") else ("Yes" if is_indexed(f) else "")
            note = choice_ref(f) if f["type"] == "Choice" else ""
            if f["type"] == "Lookup":
                note = f"→ {list_by_key(f['lookupList'])['title']} (Title); restrict delete"
            if f.get("description"):
                note = f"{note}. {f['description']}" if note else f["description"]
            w(f"| {f['displayName']} | `{f['name']}` | {TYPE_LABEL[f['type']]} | {'Yes' if f.get('required') else ''} | {idx} | {default_text(f)} | {note} |")
        w("")
        w("| View | Default | Columns | Filter / sort / group |")
        w("|---|---|---|---|")
        for v in lst["views"]:
            cols = ", ".join(display_of(lst, n) for n in v["fields"])
            desc = "<br>".join(describe_view(lst, v)) or "–"
            w(f"| {v['title']} | {'✓' if v.get('default') else ''} | {cols} | {desc} |")
        w("")
    w("### Choice sets")
    w("")
    for name, values in SCHEMA["choiceSets"].items():
        w(f'<a id="choice-{name.lower()}"></a>**{name}**: ' + " · ".join(values))
        w("")
    return "\n".join(out).rstrip() + "\n"


# ---------------------------------------------------------------------------
# manual-provisioning.md
# ---------------------------------------------------------------------------
def render_checklists() -> str:
    out: list[str] = []
    w = out.append
    for n, lst in enumerate(LISTS, 1):
        w(f"### {n}. {lst['title']}")
        w("")
        slug = lst["url"].split("/", 1)[1]
        w(f"- [ ] **Create the list:** *New › List › Blank list*. Name it `{slug}` (no spaces, so the URL is `{lst['url']}`) "
          f"and click Create. Then, in *List settings › List name, description and navigation*, rename it to **{lst['title']}** "
          f"and paste the description: *{lst['description']}*")
        w(f"- [ ] **Versioning:** *List settings › Versioning settings*. Set *Create a version each time you edit an item* to **Yes** "
          f"and *Keep the following number of versions* to **500**.")
        w(f"- [ ] **Attachments:** *List settings › Advanced settings*. Set *Attachments to list items* to "
          f"**{'Enabled' if lst.get('attachments') else 'Disabled'}**.")
        t = lst["titleField"]
        extra = ""
        if t.get("unique"):
            extra = " Set *Enforce unique values* to **Yes** and accept the prompt to index it."
        elif t.get("indexed"):
            extra = " Index it (see the indexes step)."
        w(f"- [ ] **Title column:** click *Title*, rename it to **{t['displayName']}**, and leave *Require* on **Yes**.{extra}")
        w("- [ ] **Columns:** *List settings › Create column*. Type the **internal name** first, save, then edit the column "
          "and change *Column name* to the **display name** (details in step 3):")
        w("")
        w("  | # | Create as (internal) | Rename to (display) | Column type and settings | Required | Unique | Default |")
        w("  |---|---|---|---|---|---|---|")
        for i, f in enumerate(lst["fields"], 1):
            ui = TYPE_UI[f["type"]]
            if f["type"] == "Choice":
                ui += "; Choices (one per line):<br>" + "<br>".join(f"`{c}`" for c in choices(f))
            if f["type"] == "Lookup":
                ui += (f"; Get information from: **{list_by_key(f['lookupList'])['title']}**; In this column: **Title**; "
                       "Enforce relationship behaviour: **Yes › Restrict delete** (accept the prompt to index)")
            w(f"  | {i} | `{f['name']}` | {f['displayName']} | {ui} | {'Yes' if f.get('required') else 'No'} | "
              f"{'Yes' if f.get('unique') else 'No'} | {default_text(f) or '–'} |")
        w("")
        idx = (["Title (" + t["displayName"] + ")"] if t.get("indexed") else []) + \
              [f"{f['displayName']} (`{f['name']}`)" for f in lst["fields"] if is_indexed(f)]
        w(f"- [ ] **Indexes:** *List settings › Indexed columns › Create a new index*, one per column "
          f"(unique and lookup columns may already be indexed): {', '.join(idx) if idx else '*none*'}.")
        w("- [ ] **Views:** for *All Items*, use *List settings › Views › All Items*. For the others, use *Create view › Standard view*. "
          "Tick exactly the columns listed, in this order (use *Position from Left*). Set *Item Limit* to 100, "
          "*Display items in batches*.")
        w("")
        for v in lst["views"]:
            cols = ", ".join(display_of(lst, c) for c in v["fields"])
            steps = describe_view(lst, v)
            default = " · **Make this the default view**" if v.get("default") else ""
            w(f"  - **{v['title']}**{default}. Columns: {cols}.")
            for s in steps:
                w(f"    - {s}")
        w("")
    return "\n".join(out).rstrip() + "\n"


# ---------------------------------------------------------------------------
def replace_section(path: Path, name: str, body: str, check: bool) -> bool:
    text = path.read_text(encoding="utf-8")
    pattern = re.compile(rf"(<!-- BEGIN GENERATED: {name} -->\n).*?(<!-- END GENERATED: {name} -->)", re.S)
    if not pattern.search(text):
        raise SystemExit(f"{path}: markers for '{name}' not found")
    new = pattern.sub(lambda m: m.group(1) + "\n" + body + "\n" + m.group(2), text)
    if new == text:
        return False
    if check:
        print(f"{path.relative_to(REPO)}: generated section '{name}' is out of date")
        return True
    path.write_text(new, encoding="utf-8")
    print(f"Updated {path.relative_to(REPO)} ({name})")
    return True


def main(check: bool = False) -> int:
    stale = replace_section(REPO / "docs" / "list-schema.md", "lists", render_schema(), check)
    stale |= replace_section(REPO / "docs" / "manual-provisioning.md", "checklists", render_checklists(), check)
    return 1 if (check and stale) else 0


if __name__ == "__main__":
    sys.exit(main(check="--check" in sys.argv))
