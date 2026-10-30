"""
Analysis engine of TWB Inspector.

Reads a Tableau workbook (.twb or .twbx) directly from its XML and builds:
  * every field (raw fields, calculated fields, parameters, groups/bins/sets)
  * a dependency graph between fields (which field feeds which calculation)
  * where each field is used (worksheets, dashboards, actions)
  * an audit: unused fields, duplicate calculations, complexity flags

Only the Python standard library is needed here, so the GUI, the command-line
script and the notebook can all share the same logic.

Nothing is sent anywhere: the workbook is read locally and never connected to
its data sources.
"""

from __future__ import annotations

import os
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from dataclasses import dataclass, field as dc_field
from typing import Dict, List, Optional, Set, Tuple

PARAMETERS_DS = "Parameters"

# Thresholds for the static complexity checks
LONG_FORMULA_CHARS = 800
DEEP_CHAIN_DEPTH = 5
MANY_BRANCHES = 15
STRING_HEAVY_COUNT = 5

STATUS_SHEET = "Used on worksheet"
STATUS_DASHBOARD = "Used on dashboard / action"
STATUS_INDIRECT = "Used via other calculation"
STATUS_REPORTED = "Used (reported by Tableau)"
STATUS_JOIN = "Used only to join tables"
STATUS_UNUSED = "Unused"

_TABLE_CALC_FUNCS = (
    "WINDOW_", "RUNNING_", "LOOKUP(", "INDEX(", "RANK(", "RANK_", "TOTAL(",
    "FIRST(", "LAST(", "SIZE(", "PREVIOUS_VALUE(", "SCRIPT_",
)
_AGG_FUNCS = (
    "SUM(", "AVG(", "COUNT(", "COUNTD(", "MIN(", "MAX(", "MEDIAN(", "ATTR(",
    "STDEV(", "STDEVP(", "VAR(", "VARP(", "PERCENTILE(", "COLLECT(", "CORR(", "COVAR(",
)
_STRING_FUNCS = (
    "CONTAINS(", "FIND(", "FINDNTH(", "MID(", "LEFT(", "RIGHT(", "REPLACE(",
    "SPLIT(", "REGEXP_", "UPPER(", "LOWER(", "TRIM(", "LTRIM(", "RTRIM(",
    "STARTSWITH(", "ENDSWITH(", "SPACE(", "STR(",
)


# --------------------------------------------------------------------------- #
# Data model
# --------------------------------------------------------------------------- #

@dataclass
class Field:
    key: str                      # unique id: "<datasource>::<internal name>"
    datasource: str               # internal datasource name
    datasource_caption: str       # friendly datasource name
    internal_name: str            # e.g. "[Calculation_1238...]" or "[Sales]"
    name: str                     # friendly name, e.g. "Profit Margin %"
    kind: str                     # Raw | Calculated | Parameter | Group (groups & sets) | Bin
    datatype: str = ""
    role: str = ""
    hidden: bool = False
    description: str = ""
    formula: str = ""             # formula exactly as stored in the workbook
    formula_friendly: str = ""    # formula with ids replaced by friendly names
    calc_class: str = ""
    calc_kind: str = ""           # LOD | Table calculation | Aggregate | Row-level
    param_value: str = ""
    param_domain: str = ""
    inferred: bool = False        # referenced but never declared in the file
    source_table: str = ""        # table the raw column comes from, e.g. "Orders"
    source_column: str = ""       # column name in the source database/file
    object_id: str = ""           # logical table (relationship model) the column belongs to
    family: str = ""              # logical table name Tableau stores with the column
    defined_in: str = ""          # "Workbook" or "Published data source"
    deps: Set[str] = dc_field(default_factory=set)        # fields this one uses
    used_by: Set[str] = dc_field(default_factory=set)     # fields that use this one
    unresolved: List[str] = dc_field(default_factory=list)
    sheets_direct: Set[str] = dc_field(default_factory=set)
    sheets_reported: Set[str] = dc_field(default_factory=set)
    sheets_all: Set[str] = dc_field(default_factory=set)
    dashboards_direct: Set[str] = dc_field(default_factory=set)
    dashboards_all: Set[str] = dc_field(default_factory=set)
    in_actions: bool = False
    status: str = ""
    depth: int = 0                # longest chain of calculations above this field
    upstream_count: int = 0
    downstream_count: int = 0
    flags: List[str] = dc_field(default_factory=list)
    duplicate_group: str = ""

    @property
    def is_calc(self) -> bool:
        return self.kind == "Calculated"


@dataclass
class Worksheet:
    name: str
    fields_direct: Set[str] = dc_field(default_factory=set)
    fields_reported: Set[str] = dc_field(default_factory=set)
    dashboards: Set[str] = dc_field(default_factory=set)
    hidden: bool = False


@dataclass
class Dashboard:
    name: str
    is_story: bool = False
    sheets: Set[str] = dc_field(default_factory=set)
    fields_direct: Set[str] = dc_field(default_factory=set)


@dataclass
class Table:
    datasource: str               # internal datasource name
    datasource_caption: str
    name: str                     # table name as Tableau shows it, e.g. "Orders"
    physical: str = ""            # e.g. "[dbo].[Orders]" or "[Orders$]"
    table_type: str = "Table"     # Table | Custom SQL | Stored procedure | Extract | Published data source
    logical_table: str = ""       # logical table (relationship model) this table belongs to
    connection_type: str = ""     # sqlserver, snowflake, excel-direct, textscan, hyper ...
    server: str = ""
    database: str = ""            # database name or file path
    schema: str = ""
    sql: str = ""                 # Custom SQL text
    live_or_extract: str = ""
    joins: List[str] = dc_field(default_factory=list)          # physical joins involving this table
    relationships: List[str] = dc_field(default_factory=list)  # logical relationships of its logical table
    notes: List[str] = dc_field(default_factory=list)
    columns: List[str] = dc_field(default_factory=list)        # field keys of its columns
    object_id: str = ""           # logical table id in the relationship model
    published_source: str = ""    # name of the published data source it lives in
    site: str = ""                # Tableau Server/Cloud site
    project: str = ""             # Tableau Server/Cloud project

    @property
    def key(self) -> str:
        return f"{self.datasource}::table::{self.logical_table}::{self.name}"


@dataclass
class WorkbookLineage:
    source_file: str
    tableau_version: str
    datasources: Dict[str, str]                      # internal name -> caption
    fields: Dict[str, Field]
    worksheets: Dict[str, Worksheet]
    dashboards: Dict[str, Dashboard]
    duplicate_groups: Dict[str, List[str]]           # group id -> field keys
    warnings: List[str]
    tables: List[Table] = dc_field(default_factory=list)
    datasource_info: Dict[str, Dict[str, str]] = dc_field(default_factory=dict)

    # Convenience views ---------------------------------------------------- #
    def fields_of_kind(self, *kinds: str) -> List[Field]:
        return [f for f in self.fields.values() if f.kind in kinds]

    def unused(self, *kinds: str) -> List[Field]:
        kinds = kinds or ("Calculated", "Parameter", "Group", "Bin", "Raw")
        return [f for f in self.fields.values() if f.status == STATUS_UNUSED and f.kind in kinds]

    def summary(self) -> Dict[str, int]:
        calcs = self.fields_of_kind("Calculated")
        params = self.fields_of_kind("Parameter")
        return {
            "Datasources": len([d for d in self.datasources if d != PARAMETERS_DS]),
            "Published data sources": len([d for d in self.datasource_info.values() if d.get("published")]),
            "Tables": len(self.tables),
            "Custom SQL queries": len([t for t in self.tables if t.table_type == "Custom SQL"]),
            "Fields (total)": len(self.fields),
            "Raw fields": len(self.fields_of_kind("Raw")),
            "Calculated fields": len(calcs),
            "Parameters": len(params),
            "Groups / sets / bins": len(self.fields_of_kind("Group", "Bin")),
            "Worksheets": len(self.worksheets),
            "Dashboards": len([d for d in self.dashboards.values() if not d.is_story]),
            "Stories": len([d for d in self.dashboards.values() if d.is_story]),
            "Unused calculated fields": len(self.unused("Calculated")),
            "Unused workbook calculations": len([f for f in self.unused("Calculated")
                                                 if f.defined_in != "Published data source"]),
            "Unused parameters": len(self.unused("Parameter")),
            "Unused raw fields": len(self.unused("Raw")),
            "Duplicate calculation groups": len(self.duplicate_groups),
            "Calculations with flags": len([f for f in calcs if f.flags]),
            "Hidden fields": len([f for f in self.fields.values() if f.hidden]),
        }


# --------------------------------------------------------------------------- #
# Loading
# --------------------------------------------------------------------------- #

def load_workbook_xml(path: str) -> ET.Element:
    """Return the root <workbook> element of a .twb or .twbx file."""
    lower = path.lower()
    if lower.endswith(".twbx"):
        with zipfile.ZipFile(path, "r") as z:
            twbs = [n for n in z.namelist() if n.lower().endswith(".twb")]
            if not twbs:
                raise ValueError(f"No .twb file found inside packaged workbook: {path}")
            # Prefer the workbook at the root of the archive
            twbs.sort(key=lambda n: (n.count("/"), len(n)))
            with z.open(twbs[0]) as fh:
                return ET.parse(fh).getroot()
    if lower.endswith(".twb"):
        return ET.parse(path).getroot()
    raise ValueError(f"Unsupported file type (expected .twb or .twbx): {path}")


# --------------------------------------------------------------------------- #
# Formula scanning
# --------------------------------------------------------------------------- #

def _read_bracket(text: str, i: int) -> Tuple[str, int]:
    """text[i] == '['. Return (name, index after closing bracket). ']]' is an escaped ']'."""
    j = i + 1
    out = []
    n = len(text)
    while j < n:
        ch = text[j]
        if ch == "]":
            if j + 1 < n and text[j + 1] == "]":
                out.append("]")
                j += 2
                continue
            return "".join(out), j + 1
        out.append(ch)
        j += 1
    return "".join(out), n  # unterminated, be forgiving


def scan_formula(formula: str):
    """
    Split a Tableau formula into segments.

    Returns a list of tuples:
      ("text", str)                          ordinary text
      ("string", str)                        a string literal (kept verbatim)
      ("comment", str)                       a // or /* */ comment (kept verbatim)
      ("ref", qualifier or None, name, raw)  a field reference

    Field references inside string literals and comments are NOT treated as
    references, so a commented-out field does not count as a dependency.
    """
    segs = []
    buf = []
    i, n = 0, len(formula or "")
    f = formula or ""

    def flush():
        if buf:
            segs.append(("text", "".join(buf)))
            buf.clear()

    while i < n:
        ch = f[i]
        if ch in ("'", '"'):
            flush()
            j = i + 1
            while j < n:
                if f[j] == ch:
                    if j + 1 < n and f[j + 1] == ch:   # doubled quote = escaped quote
                        j += 2
                        continue
                    break
                j += 1
            segs.append(("string", f[i:j + 1]))
            i = j + 1
        elif f.startswith("//", i):
            flush()
            j = f.find("\n", i)
            j = n if j == -1 else j
            segs.append(("comment", f[i:j]))
            i = j
        elif f.startswith("/*", i):
            flush()
            j = f.find("*/", i + 2)
            j = n if j == -1 else j + 2
            segs.append(("comment", f[i:j]))
            i = j
        elif ch == "[":
            flush()
            name1, j = _read_bracket(f, i)
            if j + 1 < n and f[j] == "." and f[j + 1] == "[":
                name2, k = _read_bracket(f, j + 1)
                segs.append(("ref", name1, name2, f[i:k]))
                i = k
            else:
                segs.append(("ref", None, name1, f[i:j]))
                i = j
        else:
            buf.append(ch)
            i += 1
    flush()
    return segs


def _code_only(segs) -> str:
    """Formula text without comments and string contents, upper-cased (for keyword checks)."""
    parts = []
    for s in segs:
        if s[0] == "text":
            parts.append(s[1])
        elif s[0] == "ref":
            parts.append("[]")
        elif s[0] == "string":
            parts.append('""')
    return "".join(parts).upper()


def normalize_formula(formula: str) -> str:
    """Canonical form used to detect duplicate calculations (ignores comments, case and whitespace)."""
    segs = scan_formula(formula)
    parts = []
    for s in segs:
        if s[0] == "text":
            parts.append(re.sub(r"\s+", "", s[1]).upper())
        elif s[0] == "string":
            parts.append(s[1])
        elif s[0] == "ref":
            parts.append(s[3].upper())
    return "".join(parts)


def classify_calc(segs) -> str:
    code = _code_only(segs)
    if "{" in code:
        return "LOD"
    if any(fn in code.replace(" ", "") for fn in _TABLE_CALC_FUNCS):
        return "Table calculation"
    if any(fn in code.replace(" ", "") for fn in _AGG_FUNCS):
        return "Aggregate"
    return "Row-level"


def complexity_flags(f: Field, segs) -> List[str]:
    code = _code_only(segs)
    compact = code.replace(" ", "").replace("\n", "").replace("\r", "").replace("\t", "")
    flags = []
    lod_count = code.count("{")
    if lod_count > 1:
        flags.append(f"Nested/multiple LODs ({lod_count})")
    elif lod_count == 1:
        flags.append("LOD expression")
    if f.calc_kind == "Table calculation":
        flags.append("Table calculation (computed after the query)")
    if len(f.formula) > LONG_FORMULA_CHARS:
        flags.append(f"Long formula ({len(f.formula)} chars)")
    branches = len(re.findall(r"\b(ELSEIF|WHEN)\b", code))
    if branches >= MANY_BRANCHES:
        flags.append(f"Many IF/CASE branches ({branches})")
    string_ops = sum(compact.count(fn) for fn in _STRING_FUNCS)
    if string_ops >= STRING_HEAVY_COUNT:
        flags.append(f"String-heavy ({string_ops} string functions)")
    if "COUNTD(" in compact:
        flags.append("COUNTD (can be slow on large data)")
    if "RAWSQL" in compact:
        flags.append("Uses RAWSQL (database-specific SQL)")
    if "SCRIPT_" in compact:
        flags.append("Uses SCRIPT_ (external analytics service)")
    return flags


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #

def _strip(name: str) -> str:
    """'[Sales]' -> 'Sales', '[a]]b]' -> 'a]b'."""
    if name.startswith("[") and name.endswith("]"):
        return name[1:-1].replace("]]", "]")
    return name


def _bracket(name: str) -> str:
    return "[" + name.replace("]", "]]") + "]"


_INSTANCE_RE = re.compile(r"^(?:[a-z0-9]+:)+(.+):[a-z]{1,4}(?::\d+)?$")


def _instance_to_field_name(name: str) -> str:
    """'none:Category:nk' -> 'Category', 'pcto:usr:Calculation_1:qk' -> 'Calculation_1'."""
    m = _INSTANCE_RE.match(name)
    return m.group(1) if m else name


def _key(ds: str, internal: str) -> str:
    return f"{ds}::{internal}"


def _is_internal_column(name: str) -> bool:
    return "__tableau_internal_object_id__" in name or name.startswith("[:")


def _text_of_description(col: ET.Element) -> str:
    desc = col.find("desc")
    if desc is None:
        return ""
    return " ".join(t.strip() for t in desc.itertext() if t.strip())


# Matches "[datasource].[field or instance]" in sheet/dashboard XML attributes and text
_QUALIFIED_RE = re.compile(r"\[((?:[^\]]|\]\])+)\]\.\[((?:[^\]]|\]\])+)\]")


def _iter_strings(elem: ET.Element, skip_tags=("datasource-dependencies",)):
    """Yield every attribute value and text node under elem, skipping some subtrees."""
    stack = [elem]
    while stack:
        e = stack.pop()
        if e is not elem and e.tag in skip_tags:
            continue
        for v in e.attrib.values():
            yield v
        if e.text and "[" in e.text:
            yield e.text
        stack.extend(list(e))


# --------------------------------------------------------------------------- #
# Data source tables
# --------------------------------------------------------------------------- #

# Newer Tableau versions write some tags with a feature prefix, e.g.
# "_.fcp.ObjectModelEncapsulateLegacy.true...relation". Match on the tag ending.
def _tag_is(elem: ET.Element, name: str) -> bool:
    return elem.tag == name or elem.tag.endswith("..." + name)


def _children(elem: ET.Element, name: str) -> List[ET.Element]:
    return [c for c in elem if _tag_is(c, name)]


def _iter_skip(elem: ET.Element, name: str, skip: str = "extract"):
    """Like elem.iter(name) but does not descend into <extract> (the extract repeats the columns)."""
    stack = [elem]
    while stack:
        e = stack.pop()
        if e is not elem and _tag_is(e, skip):
            continue
        if _tag_is(e, name):
            yield e
        stack.extend(reversed(list(e)))


def format_expression(e: Optional[ET.Element]) -> str:
    """Turn Tableau's nested <expression op=...> tree into text such as "[A] = [B] AND [C] = [D]"."""
    if e is None:
        return ""
    op = e.attrib.get("op", "")
    kids = _children(e, "expression")
    if not kids:
        return op
    parts = [format_expression(k) for k in kids]
    if len(parts) == 2 or op.upper() in ("AND", "OR"):
        return f" {op} ".join(parts)
    return f"{op}({', '.join(parts)})"


_REL_TYPES = {"table": "Table", "text": "Custom SQL", "stored-proc": "Stored procedure"}


def _walk_relation(rel: ET.Element, out: List[dict], joins: List[str], note: str = "") -> None:
    """Collect leaf tables and join descriptions from a (possibly nested) <relation>."""
    rtype = rel.attrib.get("type", "")
    if rtype == "join":
        kids = _children(rel, "relation")
        clause = rel.find("clause")
        cond = format_expression(clause.find("expression")) if clause is not None else ""
        names = []
        for k in kids:
            start = len(out)
            _walk_relation(k, out, joins, note)
            names.extend(t["name"] for t in out[start:])
        kind = rel.attrib.get("join", "inner").capitalize()
        joins.append(f"{kind} join of {', '.join(names)}" + (f" on {cond}" if cond else ""))
        for t in out:
            if t["name"] in names:
                t.setdefault("joins", []).append(joins[-1])
        return
    if rtype in ("union", "collection", "pivot"):
        sub = "Union" if rtype == "union" else ("Pivot" if rtype == "pivot" else note)
        for k in _children(rel, "relation"):
            _walk_relation(k, out, joins, sub)
        if rtype == "union" and not _children(rel, "relation"):
            out.append({"name": rel.attrib.get("name", "Union"), "physical": "", "type": "Union",
                        "connection": rel.attrib.get("connection", ""), "sql": "", "note": "Union"})
        return
    if not rel.attrib.get("name"):
        return
    sql = (rel.text or "").strip() if rtype == "text" else ""
    physical = rel.attrib.get("table", "")
    if rtype == "stored-proc":
        physical = rel.attrib.get("name", "")
    out.append({"name": rel.attrib.get("name", ""), "physical": physical,
                "type": _REL_TYPES.get(rtype, rtype.capitalize() or "Table"),
                "connection": rel.attrib.get("connection", ""), "sql": sql, "note": note})


_OBJECT_ID_RE = re.compile(r"^(?P<cap>.+?) \((?P<phys>[^()]+)\)_[0-9A-Fa-f]{32}$")
_PROJECT_RE = re.compile(r"\s*\|\s*Project\s*:\s*(?P<project>.+)$", re.I)


def physical_from_object_id(object_id: str) -> str:
    """'ORDERS (PRD.ORDERS)_9F2C…' -> 'PRD.ORDERS' (Tableau stores the physical name in the id)."""
    m = _OBJECT_ID_RE.match(object_id.strip("[]"))
    return m.group("phys") if m else ""


def published_info(ds: ET.Element) -> Dict[str, str]:
    """Where a published data source lives: name, server, site, project, revision."""
    conn = next((c for c in ds if _tag_is(c, "connection")), None)
    if conn is None or conn.attrib.get("class") != "sqlproxy":
        return {}
    repo = next((c for c in ds if _tag_is(c, "repository-location")), None)
    repo_attr = dict(repo.attrib) if repo is not None else {}
    friendly = conn.attrib.get("server-ds-friendly-name") or ds.attrib.get("caption") or conn.attrib.get("dbname", "")
    m = _PROJECT_RE.search(friendly)
    name = _PROJECT_RE.sub("", friendly).strip() or friendly
    server = conn.attrib.get("server", "")
    if conn.attrib.get("port") not in (None, "", "80", "443"):
        server += ":" + conn.attrib["port"]
    return {
        "published": "yes",
        "name": name,
        "project": m.group("project").strip() if m else "",
        "server": server,
        "site": repo_attr.get("site", ""),
        "content_name": repo_attr.get("id") or conn.attrib.get("dbname", ""),
        "revision": repo_attr.get("revision", ""),
    }


def parse_datasource_tables(ds: ET.Element, ds_name: str, ds_caption: str) -> List[Table]:
    """All tables a datasource reads from, with connection details, joins and relationships."""
    conn = next((c for c in ds if _tag_is(c, "connection")), None)
    if conn is None:
        return []

    # Connection details: one per named connection (federated) or the connection itself
    named: Dict[str, Dict[str, str]] = {}
    for nc in conn.iter("named-connection"):
        inner = nc.find("connection")
        if inner is not None:
            named[nc.attrib.get("name", "")] = dict(inner.attrib)
    own = dict(conn.attrib)

    def conn_info(conn_name: str) -> Dict[str, str]:
        if conn_name and conn_name in named:
            return named[conn_name]
        if len(named) == 1:
            return next(iter(named.values()))
        return own

    extract = next((c for c in ds if _tag_is(c, "extract")), None)
    live_or_extract = "Extract" if extract is not None and extract.attrib.get("enabled", "true") == "true" else "Live"
    if own.get("class") in ("dataengine", "hyper"):
        live_or_extract = "Extract"

    # Logical tables (relationship model, Tableau 2020.2+)
    groups: List[Tuple[str, ET.Element, str]] = []  # (logical table caption, relation element, object id)
    relationships: List[Tuple[str, str, str]] = []
    graph = next((c for c in ds if _tag_is(c, "object-graph")), None)
    if graph is not None:
        objects = next((c for c in graph if _tag_is(c, "objects")), None)
        captions: Dict[str, str] = {}
        for obj in (list(objects) if objects is not None else []):
            if not _tag_is(obj, "object"):
                continue
            cap = obj.attrib.get("caption") or obj.attrib.get("id", "")
            captions[obj.attrib.get("id", "")] = cap
            # context='' is the live definition; context='extract' repeats it for the extract
            props_list = [p for p in obj if _tag_is(p, "properties")]
            main_props = [p for p in props_list if p.attrib.get("context", "") == ""] or props_list[:1]
            for props in main_props:
                for rel in _children(props, "relation"):
                    groups.append((cap, rel, obj.attrib.get("id", "")))
        rels_root = next((c for c in graph if _tag_is(c, "relationships")), None)
        for r in (list(rels_root) if rels_root is not None else []):
            a = r.find("first-end-point")
            b = r.find("second-end-point")
            if a is None or b is None:
                continue
            relationships.append((captions.get(a.attrib.get("object-id", ""), a.attrib.get("object-id", "")),
                                  captions.get(b.attrib.get("object-id", ""), b.attrib.get("object-id", "")),
                                  format_expression(r.find("expression"))))
    if not groups:
        # Older workbooks: the physical model sits directly under <connection>.
        # Prefer the legacy (non-collection) variant if Tableau wrote both.
        rels = _children(conn, "relation")
        legacy = [r for r in rels if ".true..." not in r.tag] or rels
        groups = [("", r, "") for r in legacy]

    pub = published_info(ds)
    if pub:
        # A published data source: the workbook knows its logical tables (and the physical
        # table hidden in each table's id), not the joins or SQL inside them.
        objects = [(cap, oid) for cap, _rel, oid in groups if oid] or [(pub["name"], "")]
        tables: List[Table] = []
        for cap, oid in dict.fromkeys(objects):
            rel_text = [f"{a} ↔ {b}" + (f" on {e}" if e else "") for a, b, e in relationships if cap in (a, b)]
            tables.append(Table(
                datasource=ds_name, datasource_caption=ds_caption, name=cap,
                physical=physical_from_object_id(oid), table_type="Published table" if oid else "Published data source",
                logical_table=cap if oid else "", connection_type="Tableau published data source",
                server=pub["server"], database=pub["content_name"], live_or_extract=live_or_extract,
                relationships=rel_text, object_id=oid, published_source=pub["name"],
                site=pub["site"], project=pub["project"],
                notes=["Joins and Custom SQL inside a published data source are stored on the server, not in the workbook"]))
        return tables

    tables = []
    seen = set()
    for logical, rel, object_id in groups:
        leaves: List[dict] = []
        joins: List[str] = []
        _walk_relation(rel, leaves, joins)
        rel_text = [f"{a} ↔ {b}" + (f" on {e}" if e else "") for a, b, e in relationships
                    if logical and logical in (a, b)]
        for leaf in leaves:
            ident = (logical, leaf["name"], leaf["physical"])
            if ident in seen:
                continue
            seen.add(ident)
            info = conn_info(leaf["connection"])
            cls = info.get("class", "")
            t = Table(datasource=ds_name, datasource_caption=ds_caption, name=leaf["name"],
                      physical=leaf["physical"], table_type=leaf["type"], logical_table=logical,
                      connection_type=cls, server=info.get("server", ""),
                      database=info.get("dbname") or info.get("filename", ""),
                      schema=info.get("schema", ""), sql=leaf["sql"], live_or_extract=live_or_extract,
                      joins=leaf.get("joins", []), relationships=rel_text, object_id=object_id)
            if leaf.get("note"):
                t.notes.append(leaf["note"])
            if cls in ("dataengine", "hyper"):
                t.live_or_extract = "Extract"
                t.table_type = "Extract"
            if cls == "sqlproxy":
                t.table_type = "Published data source"
                t.notes.append("Published data source: the real tables are defined on Tableau Server/Cloud")
            tables.append(t)
    if not tables and (own.get("class") or named):
        # Connection saved without table details (old or trimmed workbooks)
        info = own if own.get("class") != "federated" else conn_info("")
        if own.get("class") == "federated" and len(named) > 1:
            info = {k: ", ".join(sorted({v.get(k, "") for v in named.values()} - {""}))
                    for k in ("class", "server", "dbname", "filename")}
        tables.append(Table(datasource=ds_name, datasource_caption=ds_caption, name="(no table saved in file)",
                            table_type="Connection only", connection_type=info.get("class", ""),
                            server=info.get("server", ""), database=info.get("dbname") or info.get("filename", ""),
                            live_or_extract=live_or_extract,
                            notes=["The workbook stores the connection but not which tables it reads"]))
    return tables


# --------------------------------------------------------------------------- #
# Main analysis
# --------------------------------------------------------------------------- #

def analyze_workbook(path: str) -> WorkbookLineage:
    root = load_workbook_xml(path)
    warnings: List[str] = []
    version = root.attrib.get("source-build") or root.attrib.get("version", "")

    fields: Dict[str, Field] = {}
    datasources: Dict[str, str] = {}
    by_ds_name: Dict[str, Dict[str, str]] = defaultdict(dict)   # ds -> internal name -> key
    internal_captions: Dict[str, str] = {}                        # key -> caption for internal cols
    seg_cache: Dict[str, list] = {}

    tables: List[Table] = []
    datasource_info: Dict[str, Dict[str, str]] = {}
    ds_root = root.find("datasources")
    ds_elems = list(ds_root) if ds_root is not None else []

    # ---- 1. Collect fields from every datasource -------------------------- #
    for ds in ds_elems:
        if ds.tag != "datasource":
            continue
        ds_name = ds.attrib.get("name", "")
        ds_caption = ds.attrib.get("caption") or (PARAMETERS_DS if ds_name == PARAMETERS_DS else ds_name)
        datasources[ds_name] = ds_caption
        is_param_ds = ds_name == PARAMETERS_DS
        ds_pub = published_info(ds)
        is_published = bool(ds_pub)
        datasource_info[ds_name] = {"caption": ds_caption, "connection": "", **ds_pub}
        conn_el = next((c for c in ds if _tag_is(c, "connection")), None)
        if conn_el is not None:
            datasource_info[ds_name]["connection"] = conn_el.attrib.get("class", "")

        def add_field(internal: str, **kw) -> Field:
            k = _key(ds_name, internal)
            if k in fields:
                f = fields[k]
                for attr, val in kw.items():
                    if val and not getattr(f, attr):
                        setattr(f, attr, val)
                return f
            f = Field(key=k, datasource=ds_name, datasource_caption=ds_caption,
                      internal_name=internal, name=kw.pop("name", None) or _strip(internal),
                      kind=kw.pop("kind", "Raw"), **kw)
            fields[k] = f
            by_ds_name[ds_name][internal] = k
            return f

        # Raw columns known only from the connection metadata
        records = list(_iter_skip(ds, "metadata-record")) or list(ds.iter("metadata-record"))
        for rec in records:
            if rec.attrib.get("class") not in ("column", "measure"):
                continue
            local = (rec.findtext("local-name") or "").strip()
            if not local or _is_internal_column(local):
                continue
            is_calc = (rec.findtext("remote-type") or "").strip() == "-1"   # calculation defined in a published source
            f = add_field(local, kind="Calculated" if is_calc else "Raw",
                          name=(rec.findtext("caption") or "").strip() or None,
                          datatype=(rec.findtext("local-type") or "").strip())
            f.defined_in = f.defined_in or ("Published data source" if is_published else "")
            if (rec.findtext("hidden") or "").strip() == "true":
                f.hidden = True
            if not f.object_id:
                f.object_id = (rec.findtext("object-id") or "").strip().strip("[]")
            if not f.family:
                f.family = (rec.findtext("family") or "").strip()
            if not f.source_table and not is_calc:
                parent = _strip((rec.findtext("parent-name") or "").strip())
                f.source_table = "" if parent == "sqlproxy" else parent
                f.source_column = (rec.findtext("remote-name") or "").strip()

        # Declared columns (raw fields with settings, calculations, parameters, bins)
        for col in ds.findall("column"):
            internal = col.attrib.get("name", "")
            if not internal:
                continue
            if _is_internal_column(internal):
                if col.attrib.get("caption"):
                    internal_captions[_key(ds_name, internal)] = col.attrib["caption"]
                continue
            calc = col.find("calculation")
            calc_class = calc.attrib.get("class", "") if calc is not None else ""
            formula = calc.attrib.get("formula", "") if calc is not None else ""
            if is_param_ds or col.attrib.get("param-domain-type"):
                kind = "Parameter"
            elif calc_class == "tableau":
                kind = "Calculated"
            elif calc_class == "bin":
                kind = "Bin"
            elif calc_class == "categorical-bin":
                kind = "Group"
            elif calc is not None:
                kind = "Calculated"
            else:
                kind = "Raw"
            domain = ""
            if kind == "Parameter":
                members = [m.attrib.get("value", "") for m in col.findall("./members/member")]
                rng = col.find("range")
                if members:
                    domain = "List: " + ", ".join(members[:25]) + (" ..." if len(members) > 25 else "")
                elif rng is not None:
                    domain = "Range: " + ", ".join(f"{k}={v}" for k, v in rng.attrib.items())
                else:
                    domain = col.attrib.get("param-domain-type", "")
            f = add_field(
                internal,
                name=col.attrib.get("caption") or _strip(internal),
                kind=kind,
                datatype=col.attrib.get("datatype", ""),
                role=col.attrib.get("role", ""),
                description=_text_of_description(col),
                formula=formula,
                calc_class=calc_class,
                param_value=col.attrib.get("value", "") if kind == "Parameter" else "",
                param_domain=domain,
            )
            # Values from <column> win over metadata-only values
            f.kind = kind
            if col.attrib.get("caption"):
                f.name = col.attrib["caption"]
            if "hidden" in col.attrib:
                f.hidden = col.attrib["hidden"] == "true"
            if not f.defined_in and kind != "Parameter":
                f.defined_in = "Published data source" if (is_published and kind == "Raw") else "Workbook"
            if calc is not None and calc.attrib.get("column"):
                # categorical bins / groups point at their source column
                f.formula = f.formula or calc.attrib["column"]

        # Groups and sets
        for grp in ds.findall("group"):
            internal = grp.attrib.get("name", "")
            if not internal or _is_internal_column(internal):
                continue
            levels = sorted({gf.attrib["level"] for gf in grp.iter("groupfilter") if gf.attrib.get("level")})
            f = add_field(internal, name=grp.attrib.get("caption") or _strip(internal),
                          kind="Group", formula=" ".join(levels))
            f.hidden = grp.attrib.get("hidden", "false") == "true"

        if not is_param_ds:
            tables.extend(parse_datasource_tables(ds, ds_name, ds_caption))

    # Link raw columns to their table: by physical table name (parent-name), else by the
    # logical table id the column belongs to (relationship model, published data sources)
    by_object = {(t.datasource, t.object_id): t for t in tables if t.object_id}
    for f in fields.values():
        if f.kind != "Raw":
            continue
        named = [t for t in tables if t.datasource == f.datasource and t.name == f.source_table and f.source_table]
        target = named[0] if named else by_object.get((f.datasource, f.object_id))
        if target is None and f.family:
            fam = [t for t in tables if t.datasource == f.datasource and f.family in (t.name, t.logical_table)]
            target = fam[0] if len(fam) == 1 else None
        if target is not None:
            target.columns.append(f.key)
            f.source_table = target.name
    # A datasource with a single table owns all its raw columns
    by_ds_tables: Dict[str, List[Table]] = defaultdict(list)
    for t in tables:
        by_ds_tables[t.datasource].append(t)
    for dsn, tl in by_ds_tables.items():
        if len(tl) == 1 and not tl[0].columns:
            for f in fields.values():
                if f.datasource == dsn and f.kind == "Raw":
                    tl[0].columns.append(f.key)
                    f.source_table = f.source_table or tl[0].name

    # Column metadata of a published source can reveal Custom SQL inside it
    for t in tables:
        if t.table_type.startswith("Published") and t.columns:
            fams = {fields[k].family for k in t.columns}
            if any(fm.startswith("Custom SQL Query") for fm in fams):
                t.notes.insert(0, "Built on Custom SQL inside the published data source (SQL text is on the server)")

    ds_names = set(datasources)
    ds_by_caption = {cap: n for n, cap in datasources.items()}

    # ---- 2. Resolve dependencies between fields --------------------------- #
    def resolve(ds_name: str, qualifier: Optional[str], name: str, create: bool = True) -> Optional[str]:
        if qualifier is not None:
            if qualifier.startswith("__tableau_internal_object_id__"):
                return None
            if qualifier in ds_names:
                target_ds = qualifier
            elif qualifier in ds_by_caption:
                target_ds = ds_by_caption[qualifier]
            else:
                # Not a known datasource: treat "[a].[b]" as part of a name in the same datasource
                target_ds = ds_name
                name = f"{qualifier}].[{name}"
        else:
            target_ds = ds_name
        internal = _bracket(name)
        k = by_ds_name[target_ds].get(internal)
        if k:
            return k
        if not create:
            return None
        # Referenced but never declared (e.g. columns of a published data source).
        cap = datasources.get(target_ds, target_ds)
        f = Field(key=_key(target_ds, internal), datasource=target_ds, datasource_caption=cap,
                  internal_name=internal, name=_strip(internal),
                  kind="Parameter" if target_ds == PARAMETERS_DS else "Raw", inferred=True)
        fields[f.key] = f
        by_ds_name[target_ds][internal] = f.key
        return f.key

    for f in list(fields.values()):
        # a parameter's "formula" is just its current value, not a dependency
        if f.kind == "Parameter" or not f.formula:
            continue
        segs = scan_formula(f.formula)
        seg_cache[f.key] = segs
        for s in segs:
            if s[0] != "ref":
                continue
            _, qual, nm, raw = s
            target = resolve(f.datasource, qual, nm)
            if target is None:
                continue
            if target == f.key:
                continue
            f.deps.add(target)
            if fields[target].inferred and raw not in f.unresolved:
                f.unresolved.append(raw)

    for f in fields.values():
        for d in f.deps:
            fields[d].used_by.add(f.key)

    # ---- 3. Friendly formulas, calc kinds, flags --------------------------- #
    def friendly(f: Field) -> str:
        segs = seg_cache.get(f.key)
        if not segs:
            return f.formula
        out = []
        for s in segs:
            if s[0] != "ref":
                out.append(s[1])
                continue
            _, qual, nm, raw = s
            if qual and qual.startswith("__tableau_internal_object_id__"):
                cap = internal_captions.get(_key(f.datasource, raw))
                out.append(_bracket(cap) if cap else raw)
                continue
            target = resolve(f.datasource, qual, nm, create=False)
            if not target:
                out.append(raw)
                continue
            t = fields[target]
            if t.datasource != f.datasource and t.datasource != PARAMETERS_DS:
                out.append(_bracket(t.datasource_caption) + "." + _bracket(t.name))
            else:
                out.append(_bracket(t.name))
        return "".join(out)

    for f in fields.values():
        if f.kind == "Parameter":
            f.formula_friendly = ""
            continue
        f.formula_friendly = friendly(f)
        if f.kind == "Calculated":
            segs = seg_cache.get(f.key, [])
            f.calc_kind = classify_calc(segs)
            f.flags = complexity_flags(f, segs)

    # ---- 4. Worksheets ------------------------------------------------------ #
    worksheets: Dict[str, Worksheet] = {}
    ws_root = root.find("worksheets")

    def mark_qualified(strings, instance_map, sink: Set[str]):
        for s in strings:
            for q, nm in _QUALIFIED_RE.findall(s):
                q = q.replace("]]", "]")
                nm = nm.replace("]]", "]")
                if q not in ds_names:
                    continue
                col = instance_map.get((q, _bracket(nm)))
                if col is None:
                    col = _bracket(_instance_to_field_name(nm))
                k = by_ds_name[q].get(col)
                if k:
                    sink.add(k)

    for ws in (list(ws_root) if ws_root is not None else []):
        if ws.tag != "worksheet":
            continue
        w = Worksheet(name=ws.attrib.get("name", ""))
        worksheets[w.name] = w
        instance_map: Dict[Tuple[str, str], str] = {}
        for dep in ws.iter("datasource-dependencies"):
            dsn = dep.attrib.get("datasource", "")
            for col in dep.findall("column"):
                k = by_ds_name[dsn].get(col.attrib.get("name", ""))
                if k:
                    w.fields_reported.add(k)
            for ci in dep.findall("column-instance"):
                instance_map[(dsn, ci.attrib.get("name", ""))] = ci.attrib.get("column", "")
                k = by_ds_name[dsn].get(ci.attrib.get("column", ""))
                if k:
                    w.fields_direct.add(k)
        # shelves, filters, encodings, titles, tooltips, sorts ...
        mark_qualified(_iter_strings(ws), instance_map, w.fields_direct)
        for k in w.fields_direct:
            fields[k].sheets_direct.add(w.name)
        for k in w.fields_reported:
            fields[k].sheets_reported.add(w.name)

    # Hidden worksheets (only visible inside dashboards)
    win_root = root.find("windows")
    if win_root is not None:
        for win in win_root:
            if win.attrib.get("class") == "worksheet" and win.attrib.get("hidden") == "true":
                if win.attrib.get("name") in worksheets:
                    worksheets[win.attrib["name"]].hidden = True

    # ---- 5. Dashboards and stories ------------------------------------------ #
    dashboards: Dict[str, Dashboard] = {}
    db_root = root.find("dashboards")
    for db in (list(db_root) if db_root is not None else []):
        if db.tag != "dashboard":
            continue
        d = Dashboard(name=db.attrib.get("name", ""), is_story=db.attrib.get("type") == "storyboard")
        dashboards[d.name] = d
        for z in db.iter("zone"):
            zn = z.attrib.get("name")
            if zn and zn in worksheets:
                d.sheets.add(zn)
        mark_qualified(_iter_strings(db), {}, d.fields_direct)
    # stories can contain dashboards: inherit their sheets
    for d in dashboards.values():
        if d.is_story:
            db_elem = None
            for e in (list(db_root) if db_root is not None else []):
                if e.attrib.get("name") == d.name:
                    db_elem = e
            if db_elem is not None:
                for z in db_elem.iter("zone"):
                    zn = z.attrib.get("name")
                    if zn in dashboards and zn != d.name:
                        d.sheets |= dashboards[zn].sheets
    for d in dashboards.values():
        for s in d.sheets:
            worksheets[s].dashboards.add(d.name)
        for k in d.fields_direct:
            fields[k].dashboards_direct.add(d.name)

    # ---- 6. Actions (filter / highlight / parameter actions) ---------------- #
    act_root = root.find("actions")
    action_fields: Set[str] = set()
    if act_root is not None:
        mark_qualified(_iter_strings(act_root, skip_tags=()), {}, action_fields)
        for k in action_fields:
            fields[k].in_actions = True

    # ---- 7. Usage status ----------------------------------------------------- #
    def upstream(start: Set[str]) -> Set[str]:
        seen, stack = set(), list(start)
        while stack:
            k = stack.pop()
            if k in seen:
                continue
            seen.add(k)
            stack.extend(fields[k].deps)
        return seen

    def downstream(k0: str) -> Set[str]:
        seen, stack = set(), [k0]
        while stack:
            k = stack.pop()
            if k in seen:
                continue
            seen.add(k)
            stack.extend(fields[k].used_by)
        seen.discard(k0)
        return seen

    direct_sheet = {k for k, f in fields.items() if f.sheets_direct}
    direct_other = {k for k, f in fields.items() if f.dashboards_direct or f.in_actions}
    reachable = upstream(direct_sheet | direct_other)

    # Columns that link tables together (join clauses and relationships)
    join_keys: Set[str] = set()
    for t in tables:
        for text in t.joins + t.relationships:
            expr = text.split(" on ", 1)[1] if " on " in text else ""
            for seg in scan_formula(expr):
                if seg[0] != "ref":
                    continue
                _, qual, nm, _raw = seg
                if qual is None:                       # relationship: local field name
                    k = by_ds_name[t.datasource].get(_bracket(nm))
                    if k:
                        join_keys.add(k)
                else:                                  # join: [table].[database column]
                    for k2, f2 in fields.items():
                        if (f2.datasource == t.datasource and f2.source_table == qual
                                and (f2.source_column or _strip(f2.internal_name)) == nm):
                            join_keys.add(k2)

    for k, f in fields.items():
        if f.sheets_direct:
            f.status = STATUS_SHEET
        elif k in direct_other:
            f.status = STATUS_DASHBOARD
        elif k in reachable:
            f.status = STATUS_INDIRECT
        elif f.sheets_reported:
            f.status = STATUS_REPORTED
        elif k in join_keys:
            f.status = STATUS_JOIN
        else:
            f.status = STATUS_UNUSED

    # Worksheets/dashboards reached by each field (through everything downstream)
    for k, f in fields.items():
        down = downstream(k)
        f.downstream_count = len(down)
        sheets = set(f.sheets_direct) | set(f.sheets_reported)
        dbs = set(f.dashboards_direct)
        for d in down:
            sheets |= fields[d].sheets_direct
            dbs |= fields[d].dashboards_direct
        f.sheets_all = sheets
        for s in sheets:
            dbs |= worksheets[s].dashboards if s in worksheets else set()
        f.dashboards_all = dbs
        f.upstream_count = len(upstream({k})) - 1

    # Longest chain of calculations above each field
    depth_cache: Dict[str, int] = {}

    def depth(k: str, visiting: Set[str]) -> int:
        if k in depth_cache:
            return depth_cache[k]
        if k in visiting:
            warnings.append(f"Circular reference detected at {fields[k].name}")
            return 0
        visiting.add(k)
        deps = fields[k].deps
        d = 0 if not deps else 1 + max(depth(x, visiting) for x in deps)
        visiting.discard(k)
        depth_cache[k] = d
        return d

    for k, f in fields.items():
        f.depth = depth(k, set())

    # A calculation built on a table calculation is itself computed after the query,
    # and one built on an aggregate is itself aggregated. Walk from the bottom up.
    for f in sorted(fields.values(), key=lambda x: x.depth):
        if f.kind != "Calculated":
            continue
        kinds = {fields[d].calc_kind for d in f.deps}
        if f.calc_kind in ("Row-level", "Aggregate") and "Table calculation" in kinds:
            f.calc_kind = "Table calculation"
            f.flags.insert(0, "Table calculation (computed after the query)")
        elif f.calc_kind == "Row-level" and "Aggregate" in kinds:
            f.calc_kind = "Aggregate"

    for k, f in fields.items():
        if f.is_calc and f.depth >= DEEP_CHAIN_DEPTH:
            f.flags.append(f"Deep calculation chain (depth {f.depth})")
        if f.unresolved:
            f.flags.append("References fields not declared in the file: " + ", ".join(f.unresolved))

    # ---- 8. Duplicate calculations ------------------------------------------ #
    groups: Dict[str, List[str]] = defaultdict(list)
    for k, f in fields.items():
        if f.kind == "Calculated" and f.formula.strip():
            groups[normalize_formula(f.formula)].append(k)
    duplicate_groups: Dict[str, List[str]] = {}
    for i, (_, keys) in enumerate(sorted((g for g in groups.items() if len(g[1]) > 1),
                                         key=lambda g: fields[g[1][0]].name.lower()), start=1):
        gid = f"D{i}"
        duplicate_groups[gid] = sorted(keys, key=lambda x: fields[x].name.lower())
        for k in keys:
            others = [fields[o].name for o in keys if o != k]
            fields[k].duplicate_group = gid
            fields[k].flags.append("Same formula as: " + ", ".join(others))

    return WorkbookLineage(
        source_file=os.path.basename(path),
        tableau_version=version,
        datasources=datasources,
        fields=fields,
        worksheets=worksheets,
        dashboards=dashboards,
        duplicate_groups=duplicate_groups,
        warnings=warnings,
        tables=tables,
        datasource_info=datasource_info,
    )


def find_workbooks(folder: str) -> List[str]:
    """All .twb/.twbx files in a folder (not recursive), sorted by name."""
    if not os.path.isdir(folder):
        return []
    return sorted(
        os.path.join(folder, f) for f in os.listdir(folder)
        if f.lower().endswith((".twb", ".twbx")) and os.path.isfile(os.path.join(folder, f))
    )
