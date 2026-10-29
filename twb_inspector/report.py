"""
Turns the analysis from twb_inspector.core into the two outputs:
  * an Excel audit workbook (Summary, Fields, Unused, Duplicates, Flags,
    Dependencies, Worksheets)
  * a self-contained interactive HTML lineage diagram (works offline)
"""

from __future__ import annotations

import json
import os
import sys
from typing import Dict, List, Optional

import pandas as pd

from .excel import write_sheets
from .core import (
    PARAMETERS_DS, STATUS_DASHBOARD, STATUS_INDIRECT, STATUS_JOIN, STATUS_REPORTED, STATUS_SHEET,
    STATUS_UNUSED, Field, WorkbookLineage, analyze_workbook,
)

STATUS_COLORS = {
    STATUS_SHEET: "#C6EFCE",
    STATUS_JOIN: "#DDEBF7",
    STATUS_DASHBOARD: "#C6EFCE",
    STATUS_INDIRECT: "#DDEBF7",
    STATUS_REPORTED: "#FFF2CC",
    STATUS_UNUSED: "#FFC7CE",
}

_KIND_ORDER = {"Parameter": 0, "Calculated": 1, "Group": 2, "Bin": 3, "Raw": 4}


def resource_path(filename: str) -> str:
    """Path of a file shipped inside the twb_inspector package (works in the PyInstaller .exe too)."""
    if hasattr(sys, "_MEIPASS"):
        base = os.path.join(sys._MEIPASS, "twb_inspector")
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, filename)


def safe_base_name(source_file: str, max_len: int = 40) -> str:
    base = os.path.splitext(os.path.basename(source_file))[0]
    keep = "".join(c if c.isalnum() or c in " -_()" else "_" for c in base).strip()
    return (keep or "workbook")[:max_len]


def _join(values) -> str:
    return ", ".join(sorted(values, key=str.lower))


def _names(lineage: WorkbookLineage, keys) -> str:
    return _join(lineage.fields[k].name for k in keys)


def _sorted_fields(lineage: WorkbookLineage) -> List[Field]:
    return sorted(lineage.fields.values(),
                  key=lambda f: (_KIND_ORDER.get(f.kind, 9), f.datasource_caption.lower(), f.name.lower()))


# --------------------------------------------------------------------------- #
# Excel
# --------------------------------------------------------------------------- #

def build_tables(lineage: WorkbookLineage) -> Dict[str, pd.DataFrame]:
    fields = _sorted_fields(lineage)

    summary_rows = [("Workbook", lineage.source_file), ("Tableau version", lineage.tableau_version)]
    summary_rows += [(k, v) for k, v in lineage.summary().items()]
    for w in lineage.warnings:
        summary_rows.append(("Warning", w))
    summary = pd.DataFrame(summary_rows, columns=["Item", "Value"])

    field_rows = []
    for f in fields:
        field_rows.append({
            "Field_Name": f.name,
            "Type": f.kind,
            "Calc_Kind": f.calc_kind,
            "DataType": f.datatype,
            "Role": f.role,
            "Datasource": f.datasource_caption,
            "Source_Table": f.source_table,
            "Source_Column": f.source_column,
            "Calculation": f.formula_friendly if f.kind != "Parameter" else "",
            "Parameter_Value": f.param_value,
            "Parameter_Values_Allowed": f.param_domain,
            "Depends_On": _names(lineage, f.deps),
            "Used_By": _names(lineage, f.used_by),
            "Usage_Status": f.status,
            "Used_In_Report": "No" if f.status == STATUS_UNUSED else "Yes",
            "Worksheets_Direct": _join(f.sheets_direct),
            "Worksheets_All": _join(f.sheets_all),
            "Dashboards": _join(f.dashboards_all),
            "Upstream_Fields": f.upstream_count,
            "Downstream_Fields": f.downstream_count,
            "Chain_Depth": f.depth,
            "Hidden": "Yes" if f.hidden else "",
            "Description": f.description,
            "Flags": "; ".join(f.flags),
            "Field_ID": f.internal_name,
        })
    all_fields = pd.DataFrame(field_rows)

    unused_rows = []
    for f in fields:
        if f.status != STATUS_UNUSED:
            continue
        if f.kind == "Raw":
            advice = "Not used anywhere. Hide it in Tableau to keep extracts smaller."
        elif f.used_by:
            advice = "Only used by other unused calculations. Review together with: " + _names(lineage, f.used_by)
        else:
            advice = "Nothing uses this. Candidate to delete after checking."
        unused_rows.append({
            "Field_Name": f.name, "Type": f.kind, "Datasource": f.datasource_caption,
            "Source_Table": f.source_table,
            "Calculation": f.formula_friendly if f.kind != "Parameter" else f.param_value,
            "Used_By": _names(lineage, f.used_by), "Advice": advice, "Field_ID": f.internal_name,
        })
    unused = pd.DataFrame(unused_rows, columns=["Field_Name", "Type", "Datasource", "Source_Table", "Calculation",
                                                "Used_By", "Advice", "Field_ID"])

    dup_rows = []
    for gid, keys in lineage.duplicate_groups.items():
        for k in keys:
            f = lineage.fields[k]
            dup_rows.append({"Group": gid, "Field_Name": f.name, "Datasource": f.datasource_caption,
                             "Usage_Status": f.status, "Calculation": f.formula_friendly,
                             "Field_ID": f.internal_name})
    duplicates = pd.DataFrame(dup_rows, columns=["Group", "Field_Name", "Datasource", "Usage_Status",
                                                 "Calculation", "Field_ID"])

    flag_rows = []
    for f in fields:
        for fl in f.flags:
            flag_rows.append({"Field_Name": f.name, "Datasource": f.datasource_caption,
                              "Calc_Kind": f.calc_kind, "Flag": fl, "Usage_Status": f.status,
                              "Calculation": f.formula_friendly})
    flags = pd.DataFrame(flag_rows, columns=["Field_Name", "Datasource", "Calc_Kind", "Flag",
                                             "Usage_Status", "Calculation"])

    dep_rows = []
    for f in fields:
        for d in sorted(f.deps, key=lambda x: lineage.fields[x].name.lower()):
            src = lineage.fields[d]
            dep_rows.append({"From_Field": src.name, "From_Type": src.kind,
                             "From_Datasource": src.datasource_caption,
                             "To_Field": f.name, "To_Type": f.kind,
                             "To_Datasource": f.datasource_caption})
        for s in sorted(f.sheets_direct, key=str.lower):
            dep_rows.append({"From_Field": f.name, "From_Type": f.kind,
                             "From_Datasource": f.datasource_caption,
                             "To_Field": s, "To_Type": "Worksheet", "To_Datasource": ""})
    dependencies = pd.DataFrame(dep_rows, columns=["From_Field", "From_Type", "From_Datasource",
                                                   "To_Field", "To_Type", "To_Datasource"])

    ws_rows = []
    for w in sorted(lineage.worksheets.values(), key=lambda w: w.name.lower()):
        calcs = [k for k in w.fields_direct if lineage.fields[k].kind == "Calculated"]
        ws_rows.append({
            "Worksheet": w.name,
            "Dashboards": _join(w.dashboards),
            "Hidden_Tab": "Yes" if w.hidden else "",
            "Not_On_Any_Dashboard": "Yes" if not w.dashboards else "",
            "Fields_Used_Directly": len(w.fields_direct),
            "Calculations_Used_Directly": _names(lineage, calcs),
            "All_Fields_Used_Directly": _names(lineage, w.fields_direct),
        })
    for d in sorted(lineage.dashboards.values(), key=lambda d: d.name.lower()):
        ws_rows.append({
            "Worksheet": f"[{'Story' if d.is_story else 'Dashboard'}] {d.name}",
            "Dashboards": "",
            "Hidden_Tab": "",
            "Not_On_Any_Dashboard": "",
            "Fields_Used_Directly": len(d.fields_direct),
            "Calculations_Used_Directly": "",
            "All_Fields_Used_Directly": "Sheets: " + _join(d.sheets)
                                        + ((" | Controls/filters: " + _names(lineage, d.fields_direct))
                                           if d.fields_direct else ""),
        })
    worksheets = pd.DataFrame(ws_rows)

    table_rows = []
    for t in sorted(lineage.tables, key=lambda t: (t.datasource_caption.lower(), t.logical_table.lower(), t.name.lower())):
        cols = [lineage.fields[k] for k in t.columns]
        used = [c for c in cols if c.status != STATUS_UNUSED]
        unused_cols = [c for c in cols if c.status == STATUS_UNUSED]
        notes = list(t.notes)
        if cols and not used:
            notes.append("No column of this table is used in the workbook")
        table_rows.append({
            "Datasource": t.datasource_caption,
            "Logical_Table": t.logical_table,
            "Table": t.name,
            "Physical_Table": t.physical,
            "Table_Type": t.table_type,
            "Connection_Type": t.connection_type,
            "Server": t.server,
            "Database_or_File": t.database,
            "Schema": t.schema,
            "Live_or_Extract": t.live_or_extract,
            "Columns": len(cols),
            "Columns_Used": len(used),
            "Unused_Columns": _join(c.name for c in unused_cols),
            "Joins": "; ".join(t.joins),
            "Relationships": "; ".join(t.relationships),
            "Custom_SQL": t.sql,
            "Notes": "; ".join(notes),
            "Column_Names": _join(c.name for c in cols),
        })
    tables = pd.DataFrame(table_rows, columns=[
        "Datasource", "Logical_Table", "Table", "Physical_Table", "Table_Type", "Connection_Type", "Server",
        "Database_or_File", "Schema", "Live_or_Extract", "Columns", "Columns_Used", "Unused_Columns", "Joins",
        "Relationships", "Custom_SQL", "Notes", "Column_Names"])

    return {"summary": summary, "tables": tables, "fields": all_fields, "unused": unused, "duplicates": duplicates,
            "flags": flags, "dependencies": dependencies, "worksheets": worksheets}


def write_excel(lineage: WorkbookLineage, path: str) -> str:
    t = build_tables(lineage)
    status_hl = {"column": "Usage_Status", "rules": STATUS_COLORS}
    sheets = [
        {"df": t["summary"], "name": "Summary", "widths": [36, 60],
         "color": "#B7DEE8", "wrap": False, "footer": "Summary"},
        {"df": t["tables"], "name": "Data Sources & Tables",
         "widths": [24, 18, 22, 26, 14, 14, 22, 30, 10, 10, 9, 9, 30, 40, 40, 50, 30, 50],
         "color": "#FFE699", "autofilter": True, "footer": "Data sources and tables"},
        {"df": t["fields"], "name": "All Fields",
         "widths": [28, 12, 16, 10, 10, 22, 18, 18, 60, 14, 30, 30, 30, 26, 10, 30, 30, 26, 10, 10, 10, 8, 30, 40, 30],
         "color": "#fff0b3", "autofilter": True, "highlight": status_hl, "footer": "All fields"},
        {"df": t["unused"], "name": "Unused", "widths": [28, 12, 22, 18, 60, 30, 50, 30],
         "color": "#FFC7CE", "autofilter": True, "footer": "Unused fields"},
        {"df": t["duplicates"], "name": "Duplicate Calcs", "widths": [8, 28, 22, 26, 60, 30],
         "color": "#FCE4D6", "autofilter": True, "highlight": status_hl, "footer": "Duplicate calculations"},
        {"df": t["flags"], "name": "Review Flags", "widths": [28, 22, 16, 45, 26, 60],
         "color": "#E2EFDA", "autofilter": True, "highlight": status_hl, "footer": "Review flags"},
        {"df": t["dependencies"], "name": "Dependencies", "widths": [28, 12, 22, 28, 12, 22],
         "color": "#DDEBF7", "autofilter": True, "footer": "Dependencies"},
        {"df": t["worksheets"], "name": "Worksheets", "widths": [30, 30, 10, 12, 10, 50, 60],
         "color": "#EDEDED", "autofilter": True, "footer": "Worksheets and dashboards"},
    ]
    return write_sheets(sheets, path)


# --------------------------------------------------------------------------- #
# HTML lineage diagram
# --------------------------------------------------------------------------- #

def build_graph(lineage: WorkbookLineage) -> Dict:
    ids: Dict[str, int] = {}
    nodes: List[Dict] = []

    for f in _sorted_fields(lineage):
        ids[f.key] = len(nodes)
        nodes.append({
            "id": len(nodes), "label": f.name, "kind": f.kind, "calcKind": f.calc_kind,
            "ds": f.datasource_caption if f.datasource != PARAMETERS_DS else "Parameters",
            "status": f.status, "unused": f.status == STATUS_UNUSED,
            "formula": f.formula_friendly if f.kind != "Parameter" else "",
            "paramValue": f.param_value, "paramDomain": f.param_domain,
            "datatype": f.datatype, "role": f.role, "hidden": f.hidden, "inferred": f.inferred,
            "description": f.description, "flags": f.flags, "dup": f.duplicate_group,
            "sheets": sorted(f.sheets_all, key=str.lower),
            "dashboards": sorted(f.dashboards_all, key=str.lower),
            "level": f.depth + 1, "fieldId": f.internal_name,
            "table": f.source_table, "sourceColumn": f.source_column,
        })

    table_ids: Dict[int, int] = {}
    for i, t in enumerate(sorted(lineage.tables, key=lambda t: (t.datasource_caption.lower(), t.name.lower()))):
        table_ids[id(t)] = len(nodes)
        used = sum(1 for k in t.columns if lineage.fields[k].status != STATUS_UNUSED)
        flags = list(t.notes)
        if t.columns and not used:
            flags.append("No column of this table is used in the workbook")
        nodes.append({
            "id": len(nodes), "label": t.name, "kind": "Table", "level": 0, "ds": t.datasource_caption,
            "tableType": t.table_type, "physical": t.physical, "logical": t.logical_table,
            "connection": t.connection_type, "server": t.server, "database": t.database, "schema": t.schema,
            "liveOrExtract": t.live_or_extract, "joins": t.joins, "relationships": t.relationships,
            "sql": t.sql, "columns": len(t.columns), "columnsUsed": used, "flags": flags,
            "unused": bool(t.columns) and not used,
        })

    max_level = max([n["level"] for n in nodes] or [0])
    ws_ids: Dict[str, int] = {}
    for w in sorted(lineage.worksheets.values(), key=lambda w: w.name.lower()):
        ws_ids[w.name] = len(nodes)
        nodes.append({"id": len(nodes), "label": w.name, "kind": "Worksheet", "level": max_level + 1,
                      "dashboards": sorted(w.dashboards, key=str.lower), "hidden": w.hidden,
                      "flags": [] if w.dashboards else ["Not on any dashboard"]})
    db_ids: Dict[str, int] = {}
    for d in sorted(lineage.dashboards.values(), key=lambda d: (d.is_story, d.name.lower())):
        db_ids[d.name] = len(nodes)
        nodes.append({"id": len(nodes), "label": d.name, "kind": "Story" if d.is_story else "Dashboard",
                      "level": max_level + (3 if d.is_story else 2),
                      "sheets": sorted(d.sheets, key=str.lower), "flags": []})

    edges = []
    for t in lineage.tables:
        for k in t.columns:
            edges.append({"from": table_ids[id(t)], "to": ids[k]})
    for f in lineage.fields.values():
        for d in f.deps:
            edges.append({"from": ids[d], "to": ids[f.key]})
        for s in f.sheets_direct:
            edges.append({"from": ids[f.key], "to": ws_ids[s], "use": True})
        for db in f.dashboards_direct:
            edges.append({"from": ids[f.key], "to": db_ids[db], "use": True})
    for d in lineage.dashboards.values():
        for s in d.sheets:
            edges.append({"from": ws_ids[s], "to": db_ids[d.name], "use": True})

    return {"nodes": nodes, "edges": edges}


def write_html(lineage: WorkbookLineage, path: str, visjs_path: Optional[str] = None) -> str:
    visjs_path = visjs_path or resource_path("vis-network.min.js")
    with open(visjs_path, "r", encoding="utf-8") as fh:
        visjs = fh.read()
    graph = build_graph(lineage)
    data_json = json.dumps({
        "workbook": lineage.source_file,
        "summary": lineage.summary(),
        "graph": graph,
    }, ensure_ascii=False).replace("</", "<\\/")
    title = safe_base_name(lineage.source_file, 80)
    icon = ""
    try:
        import base64
        with open(resource_path(os.path.join("assets", "favicon.png")), "rb") as fh:
            icon = "data:image/png;base64," + base64.b64encode(fh.read()).decode("ascii")
    except OSError:
        pass
    html = (HTML_TEMPLATE
            .replace("__ICON_LINK__", f'<link rel="icon" type="image/png" href="{icon}">' if icon else "")
            .replace("__ICON_IMG__", f'<img src="{icon}" alt="" width="26" height="26">' if icon else "")
            .replace("__TITLE__", _html_escape(title))
            .replace("/*__VISJS__*/", visjs.replace("</script", "<\\/script"))
            .replace("/*__DATA__*/null", data_json))
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(html)
    return path


def _html_escape(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


# --------------------------------------------------------------------------- #
# One call for everything
# --------------------------------------------------------------------------- #

def process_workbook(path: str, output_dir: str, excel: bool = True, html: bool = True) -> Dict:
    """Analyze one workbook and write the requested outputs. Returns paths and the analysis."""
    os.makedirs(output_dir, exist_ok=True)
    lineage = analyze_workbook(path)
    base = safe_base_name(path)
    result = {"lineage": lineage, "excel": None, "html": None}
    if excel:
        result["excel"] = write_excel(lineage, os.path.join(output_dir, f"{base}_Calculations_audit.xlsx"))
    if html:
        result["html"] = write_html(lineage, os.path.join(output_dir, f"{base}_lineage_diagram.html"))
    return result


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>__TITLE__ · TWB Inspector</title>
__ICON_LINK__
<script type="text/javascript">
/*__VISJS__*/
</script>
<style>
  :root {
    --bg: #f4f5f7; --panel: #ffffff; --text: #1f2328; --muted: #656d76; --line: #d8dee4;
    --accent: #0b6bcb; --raw: #8ab4f8; --calc: #f6ad55; --param: #b794f4; --group: #68d391;
    --ws: #4fd1c5; --table: #fbb6ce; --db: #4a5568; --unused: #e2e8f0; --bad: #c53030;
  }
  * { box-sizing: border-box; }
  body { margin: 0; font-family: "Segoe UI", Arial, sans-serif; background: var(--bg); color: var(--text);
         height: 100vh; display: flex; flex-direction: column; }
  header { padding: 10px 16px; background: var(--panel); border-bottom: 1px solid var(--line);
           display: flex; flex-wrap: wrap; align-items: center; gap: 8px 16px; }
  header h1 { font-size: 17px; margin: 0; display: flex; align-items: center; gap: 10px; }
  header h1 img { border-radius: 6px; }
  .brand { margin-left: auto; font-size: 12px; color: var(--muted); }
  .chips { display: flex; flex-wrap: wrap; gap: 6px; }
  .chip { font-size: 12px; padding: 3px 9px; border-radius: 999px; background: var(--bg); border: 1px solid var(--line); }
  .chip.warn { border-color: #f5c2c7; background: #fff5f5; color: var(--bad); }
  main { flex: 1; display: flex; min-height: 0; }
  aside { width: 300px; background: var(--panel); border-right: 1px solid var(--line); display: flex; flex-direction: column; min-height: 0; }
  #detail { position: relative; width: 380px; border-left: 1px solid var(--line); border-right: 0; overflow: auto; padding: 14px; display: none; }
  #detail.open { display: block; }
  .section { padding: 12px; border-bottom: 1px solid var(--line); }
  .section h3 { margin: 0 0 8px; font-size: 12px; text-transform: uppercase; letter-spacing: .04em; color: var(--muted); }
  input[type=search], select { width: 100%; padding: 7px 9px; border: 1px solid var(--line); border-radius: 6px; font-size: 13px; }
  select { margin-top: 6px; }
  label.toggle { display: flex; align-items: center; gap: 6px; font-size: 13px; margin: 4px 0; cursor: pointer; }
  #list { flex: 1; overflow: auto; }
  .item { padding: 7px 12px; border-bottom: 1px solid #f0f2f4; cursor: pointer; font-size: 13px; display: flex; gap: 8px; align-items: center; }
  .item:hover, .item.active { background: #eef5ff; }
  .item .dot { width: 10px; height: 10px; border-radius: 3px; flex: none; }
  .item .nm { flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .item .tag { font-size: 11px; color: var(--muted); }
  .item.unused .nm { color: var(--muted); text-decoration: line-through; }
  #count { font-size: 12px; color: var(--muted); margin-top: 6px; }
  #graph { flex: 1; position: relative; background: #fff; min-width: 0; }
  #network { position: absolute; inset: 0; }
  .toolbar { position: absolute; top: 10px; left: 10px; z-index: 5; display: flex; gap: 6px; flex-wrap: wrap; }
  button { font: inherit; font-size: 13px; padding: 6px 12px; border-radius: 6px; border: 1px solid var(--line);
           background: var(--panel); color: var(--text); cursor: pointer; }
  button:hover { border-color: var(--accent); color: var(--accent); }
  button.primary { background: var(--accent); color: #fff; border-color: var(--accent); }
  button.primary:hover { color: #fff; opacity: .9; }
  .legend { position: absolute; top: 10px; right: 10px; z-index: 5; background: rgba(255,255,255,.95);
            border: 1px solid var(--line); border-radius: 8px; padding: 8px 10px; font-size: 12px; display: flex; flex-wrap: wrap; gap: 6px 12px; max-width: 55%; justify-content: flex-end; }
  .legend span { display: inline-flex; align-items: center; gap: 5px; }
  .sw { width: 12px; height: 12px; border-radius: 3px; display: inline-block; border: 1px solid rgba(0,0,0,.2); }
  .sw.dash { background: var(--unused); border: 1px dashed #718096; }
  #detail h2 { font-size: 18px; margin: 0 24px 6px 0; word-break: break-word; }
  #detail .close { position: absolute; right: 12px; top: 10px; border: 0; font-size: 18px; padding: 2px 8px; }
  .badges { display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 10px; }
  .badge { font-size: 12px; padding: 2px 8px; border-radius: 4px; background: var(--bg); border: 1px solid var(--line); }
  .badge.unused { background: #fff5f5; color: var(--bad); border-color: #f5c2c7; }
  .badge.ok { background: #f0fff4; color: #276749; border-color: #c6f6d5; }
  dl { margin: 0; font-size: 13px; }
  dt { color: var(--muted); font-size: 12px; margin-top: 12px; }
  dd { margin: 3px 0 0; word-break: break-word; }
  pre { background: #f6f8fa; border: 1px solid var(--line); border-radius: 6px; padding: 10px; white-space: pre-wrap;
        word-break: break-word; font-size: 12.5px; font-family: Consolas, "Courier New", monospace; margin: 4px 0 0; max-height: 320px; overflow: auto; }
  .links { display: flex; flex-wrap: wrap; gap: 4px; }
  .links a { font-size: 12px; padding: 2px 7px; border-radius: 4px; background: #eef5ff; color: var(--accent); text-decoration: none; cursor: pointer; }
  .links a:hover { text-decoration: underline; }
  ul.flags { margin: 4px 0 0; padding-left: 18px; color: #975a16; }
  .actions { display: flex; gap: 6px; margin-top: 14px; flex-wrap: wrap; }
  .empty { color: var(--muted); font-size: 13px; }
  @media (max-width: 900px) { aside { width: 230px; } #detail { width: 300px; } }
</style>
</head>
<body>
<header>
  <h1>__ICON_IMG__<span>__TITLE__ · Calculation lineage</span></h1>
  <div class="chips" id="chips"></div>
  <span class="brand">Made with TWB Inspector</span>
</header>
<main>
  <aside>
    <div class="section">
      <input type="search" id="search" placeholder="Search fields, sheets, formulas…">
      <select id="filter">
        <option value="all">All items</option>
        <option value="Calculated">Calculated fields</option>
        <option value="Parameter">Parameters</option>
        <option value="Raw">Raw fields</option>
        <option value="Group">Groups, sets &amp; bins</option>
        <option value="sheets">Worksheets &amp; dashboards</option>
        <option value="Table">Data source tables</option>
        <option value="unused">Unused only</option>
        <option value="flagged">With complexity flags</option>
        <option value="dup">Duplicate calculations</option>
      </select>
      <div id="count"></div>
    </div>
    <div class="section">
      <h3>Show on diagram</h3>
      <label class="toggle"><input type="checkbox" id="tTables" checked> Data source tables</label>
      <label class="toggle"><input type="checkbox" id="tRaw" checked> Raw fields</label>
      <label class="toggle"><input type="checkbox" id="tUnusedRaw"> Unused raw fields</label>
      <label class="toggle"><input type="checkbox" id="tUnused" checked> Unused calculations &amp; parameters</label>
      <label class="toggle"><input type="checkbox" id="tSheets" checked> Worksheets &amp; dashboards</label>
    </div>
    <div id="list"></div>
  </aside>
  <section id="graph">
    <div class="toolbar">
      <button onclick="network.fit({animation:true})">Fit to screen</button>
      <button id="clearBtn" onclick="clearTrace()" style="display:none">Clear trace</button>
    </div>
    <div id="network"></div>
    <div class="legend">
      <span><i class="sw" style="background:var(--table)"></i>Table</span>
      <span><i class="sw" style="background:var(--raw)"></i>Raw field</span>
      <span><i class="sw" style="background:var(--calc)"></i>Calculation</span>
      <span><i class="sw" style="background:var(--param)"></i>Parameter</span>
      <span><i class="sw" style="background:var(--group)"></i>Group / set / bin</span>
      <span><i class="sw" style="background:var(--ws)"></i>Worksheet</span>
      <span><i class="sw" style="background:var(--db)"></i>Dashboard / story</span>
      <span><i class="sw dash"></i>Unused</span>
    </div>
  </section>
  <aside id="detail"></aside>
</main>

<script type="text/javascript">
var DATA = /*__DATA__*/null;
var G = DATA.graph;
var ALL = G.nodes;
var byId = {}; ALL.forEach(function (n) { byId[n.id] = n; });
var parents = {}, children = {};
ALL.forEach(function (n) { parents[n.id] = []; children[n.id] = []; });
G.edges.forEach(function (e) { parents[e.to].push(e.from); children[e.from].push(e.to); });

var COLORS = { Table: '#fbb6ce', Raw: '#8ab4f8', Calculated: '#f6ad55', Parameter: '#b794f4', Group: '#68d391', Bin: '#68d391',
               Worksheet: '#4fd1c5', Dashboard: '#4a5568', Story: '#2d3748' };
var BORDERS = { Table: '#b83280', Raw: '#3b78e7', Calculated: '#c05621', Parameter: '#6b46c1', Group: '#2f855a', Bin: '#2f855a',
                Worksheet: '#2c7a7b', Dashboard: '#1a202c', Story: '#1a202c' };
var isField = function (n) { return ['Worksheet', 'Dashboard', 'Story', 'Table'].indexOf(n.kind) < 0; };

// ---------- header chips ----------
(function () {
  var s = DATA.summary, el = document.getElementById('chips');
  var show = [['Tables', 0], ['Calculated fields', 0], ['Parameters', 0], ['Worksheets', 0], ['Dashboards', 0],
              ['Unused calculated fields', 1], ['Unused parameters', 1], ['Duplicate calculation groups', 1]];
  show.forEach(function (p) {
    var c = document.createElement('span');
    c.className = 'chip' + (p[1] && s[p[0]] > 0 ? ' warn' : '');
    c.textContent = p[0] + ': ' + s[p[0]];
    el.appendChild(c);
  });
})();

// ---------- node styling ----------
function styleNode(n, dim) {
  var bg = COLORS[n.kind] || '#cbd5e0', border = BORDERS[n.kind] || '#718096', font = (n.kind === 'Dashboard' || n.kind === 'Story') ? '#ffffff' : '#1a202c';
  var o = { id: n.id, label: n.label, level: n.level, shape: 'box', margin: 9,
            widthConstraint: { maximum: 220 }, borderWidth: 1.5, shapeProperties: { borderDashes: false },
            font: { size: 14, color: font, face: 'Segoe UI, Arial' } };
  if (n.kind === 'Parameter') o.shape = 'ellipse';
  if (n.kind === 'Table') { o.label = '\u25A6 ' + n.label + (n.tableType && n.tableType !== 'Table' ? '\n(' + n.tableType + ')' : ''); o.font.bold = true; }
  if (n.kind === 'Calculated' && (n.calcKind === 'LOD' || n.calcKind === 'Table calculation')) {
    o.label = n.label + (n.calcKind === 'LOD' ? '  {LOD}' : '  (table calc)');
  }
  if (n.unused) { bg = '#edf2f7'; border = '#718096'; font = '#4a5568'; o.shapeProperties = { borderDashes: [5, 4] }; }
  if (n.flags && n.flags.length && isField(n)) o.borderWidth = 3;
  if (dim) { bg = '#f7fafc'; border = '#e2e8f0'; font = '#cbd5e0'; }
  o.color = { background: bg, border: border, highlight: { background: bg, border: '#0b6bcb' }, hover: { background: bg, border: '#0b6bcb' } };
  o.font.color = font;
  return o;
}

// ---------- visibility ----------
function visible(n) {
  var tT = document.getElementById('tTables').checked, tRaw = document.getElementById('tRaw').checked, tUR = document.getElementById('tUnusedRaw').checked,
      tU = document.getElementById('tUnused').checked, tS = document.getElementById('tSheets').checked;
  if (n.kind === 'Worksheet' || n.kind === 'Dashboard' || n.kind === 'Story') return tS;
  if (n.kind === 'Table') return tT;
  if (n.kind === 'Raw') return tRaw && (!n.unused || tUR);
  if (n.unused) return tU;
  return true;
}

var nodes = new vis.DataSet([]), edges = new vis.DataSet([]);
var network = new vis.Network(document.getElementById('network'), { nodes: nodes, edges: edges }, {
  layout: { hierarchical: { enabled: true, direction: 'LR', sortMethod: 'directed', levelSeparation: 270,
                            nodeSpacing: 70, treeSpacing: 90, shakeTowards: 'roots' } },
  physics: { enabled: false },
  edges: { arrows: { to: { enabled: true, scaleFactor: 0.5 } }, color: { color: '#a0aec0', highlight: '#0b6bcb', hover: '#0b6bcb' },
           smooth: { type: 'cubicBezier', forceDirection: 'horizontal', roundness: 0.4 }, width: 1.2 },
  interaction: { hover: true, tooltipDelay: 150, navigationButtons: true, keyboard: { enabled: true, bindToWindow: false }, multiselect: false }
});

var traced = null;   // set of ids currently highlighted

function rebuild(fit) {
  var vis_ = {}; ALL.forEach(function (n) { if (visible(n)) vis_[n.id] = true; });
  var ns = [], es = [];
  ALL.forEach(function (n) { if (vis_[n.id]) ns.push(styleNode(n, traced && !traced.all[n.id])); });
  G.edges.forEach(function (e, i) {
    if (!vis_[e.from] || !vis_[e.to]) return;
    var dim = traced && !((traced.up[e.from] && traced.up[e.to]) || (traced.down[e.from] && traced.down[e.to]));
    es.push({ id: i, from: e.from, to: e.to, dashes: !!e.use,
              color: dim ? { color: '#edf2f7' } : undefined, width: dim ? 0.6 : (traced ? 2 : 1.2) });
  });
  nodes.clear(); edges.clear(); nodes.add(ns); edges.add(es);
  if (fit !== false) setTimeout(function () { network.fit(); }, 30);
}

['tTables', 'tRaw', 'tUnusedRaw', 'tUnused', 'tSheets'].forEach(function (id) {
  document.getElementById(id).addEventListener('change', function () { rebuild(); });
});

// ---------- lineage tracing ----------
function walk(start, map) {
  var seen = {}, stack = [start];
  while (stack.length) { var x = stack.pop(); if (seen[x]) continue; seen[x] = true; stack = stack.concat(map[x]); }
  return seen;
}
function trace(id) {
  var up = walk(id, parents), down = walk(id, children), all = {};
  Object.keys(up).forEach(function (k) { all[k] = true; });
  Object.keys(down).forEach(function (k) { all[k] = true; });
  traced = { all: all, up: up, down: down };
  // make sure everything in the lineage is visible
  var n = byId[id];
  ALL.forEach(function (m) {
    if (!all[m.id] || visible(m)) return;
    if (m.kind === 'Table') document.getElementById('tTables').checked = true;
    else if (m.kind === 'Raw') { document.getElementById('tRaw').checked = true; if (m.unused) document.getElementById('tUnusedRaw').checked = true; }
    else if (m.unused) document.getElementById('tUnused').checked = true;
    else document.getElementById('tSheets').checked = true;
  });
  document.getElementById('clearBtn').style.display = '';
  rebuild(false);
  network.fit({ nodes: Object.keys(all).map(Number).filter(function (k) { return nodes.get(k); }), animation: true });
  network.selectNodes([id]);
}
function clearTrace() {
  traced = null; document.getElementById('clearBtn').style.display = 'none'; rebuild();
}

// ---------- detail panel ----------
function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
function linkList(ids) {
  if (!ids.length) return '<span class="empty">None</span>';
  ids = ids.slice().sort(function (a, b) { return byId[a].label.localeCompare(byId[b].label); });
  return '<div class="links">' + ids.map(function (i) { return '<a data-id="' + i + '">' + esc(byId[i].label) + '</a>'; }).join('') + '</div>';
}
function nameList(arr) { return arr && arr.length ? esc(arr.join(', ')) : '<span class="empty">None</span>'; }

function showDetail(id) {
  var n = byId[id], el = document.getElementById('detail');
  var fieldParents = parents[id].filter(function (i) { return isField(byId[i]); });
  var fieldChildren = children[id].filter(function (i) { return isField(byId[i]); });
  var h = '<button class="close" onclick="closeDetail()" aria-label="Close">×</button>';
  h += '<h2>' + esc(n.label) + '</h2><div class="badges"><span class="badge">' + esc(n.kind) + '</span>';
  if (n.calcKind) h += '<span class="badge">' + esc(n.calcKind) + '</span>';
  if (n.status) h += '<span class="badge ' + (n.unused ? 'unused' : 'ok') + '">' + esc(n.status) + '</span>';
  if (n.kind === 'Table') h += '<span class="badge">' + esc(n.tableType) + '</span><span class="badge">' + esc(n.liveOrExtract) + '</span>';
  if (n.hidden) h += '<span class="badge">Hidden</span>';
  if (n.inferred) h += '<span class="badge">Not declared in file</span>';
  h += '</div><dl>';
  if (isField(n)) {
    h += '<dt>Datasource</dt><dd>' + esc(n.ds) + (n.datatype ? ' · ' + esc(n.datatype) : '') + (n.role ? ' · ' + esc(n.role) : '') + '</dd>';
    if (n.table) h += '<dt>Source table / column</dt><dd>' + esc(n.table) + (n.sourceColumn ? ' · ' + esc(n.sourceColumn) : '') + '</dd>';
    if (n.formula) h += '<dt>Formula</dt><dd><pre>' + esc(n.formula) + '</pre></dd>';
    if (n.kind === 'Parameter') h += '<dt>Current value</dt><dd>' + esc(n.paramValue) + '</dd><dt>Allowed values</dt><dd>' + esc(n.paramDomain) + '</dd>';
    if (n.description) h += '<dt>Description</dt><dd>' + esc(n.description) + '</dd>';
    h += '<dt>Depends on</dt><dd>' + linkList(fieldParents) + '</dd>';
    h += '<dt>Used by</dt><dd>' + linkList(fieldChildren) + '</dd>';
    h += '<dt>Worksheets (incl. through other calculations)</dt><dd>' + nameList(n.sheets) + '</dd>';
    h += '<dt>Dashboards</dt><dd>' + nameList(n.dashboards) + '</dd>';
    h += '<dt>Field ID</dt><dd><code>' + esc(n.fieldId) + '</code></dd>';
  } else if (n.kind === 'Table') {
    h += '<dt>Datasource</dt><dd>' + esc(n.ds) + (n.logical && n.logical !== n.label ? ' · logical table ' + esc(n.logical) : '') + '</dd>';
    if (n.physical) h += '<dt>Physical table</dt><dd><code>' + esc(n.physical) + '</code></dd>';
    h += '<dt>Connection</dt><dd>' + esc([n.connection, n.server, n.database, n.schema].filter(Boolean).join(' · ')) + '</dd>';
    if (n.sql) h += '<dt>Custom SQL</dt><dd><pre>' + esc(n.sql) + '</pre></dd>';
    if (n.joins && n.joins.length) h += '<dt>Joins</dt><dd>' + n.joins.map(esc).join('<br>') + '</dd>';
    if (n.relationships && n.relationships.length) h += '<dt>Relationships</dt><dd>' + n.relationships.map(esc).join('<br>') + '</dd>';
    h += '<dt>Columns (' + n.columnsUsed + ' of ' + n.columns + ' used)</dt><dd>' + linkList(children[id]) + '</dd>';
  } else if (n.kind === 'Worksheet') {
    h += '<dt>Fields used directly</dt><dd>' + linkList(parents[id]) + '</dd>';
    h += '<dt>On dashboards</dt><dd>' + linkList(children[id]) + '</dd>';
  } else {
    h += '<dt>Contains</dt><dd>' + linkList(parents[id]) + '</dd>';
  }
  if (n.flags && n.flags.length) h += '<dt>Flags</dt><dd><ul class="flags">' + n.flags.map(function (f) { return '<li>' + esc(f) + '</li>'; }).join('') + '</ul></dd>';
  h += '</dl><div class="actions"><button class="primary" onclick="trace(' + id + ')">Trace lineage</button>';
  if (n.formula || n.sql) h += '<button onclick="copyFormula(' + id + ', this)">Copy ' + (n.sql ? 'SQL' : 'formula') + '</button>';
  h += '</div>';
  el.innerHTML = h; el.classList.add('open');
  el.querySelectorAll('a[data-id]').forEach(function (a) { a.onclick = function () { focusNode(+a.dataset.id); }; });
  document.querySelectorAll('.item').forEach(function (it) { it.classList.toggle('active', +it.dataset.id === id); });
}
function closeDetail() { document.getElementById('detail').classList.remove('open'); network.unselectAll(); }
function copyFormula(id, btn) {
  var t = byId[id].formula || byId[id].sql;
  try { navigator.clipboard.writeText(t).then(function () { btn.textContent = 'Copied'; }); }
  catch (e) { btn.textContent = 'Copy not available'; }
}
function focusNode(id) {
  var n = byId[id];
  if (!visible(n)) {
    if (n.kind === 'Table') document.getElementById('tTables').checked = true;
    else if (n.kind === 'Raw') { document.getElementById('tRaw').checked = true; if (n.unused) document.getElementById('tUnusedRaw').checked = true; }
    else if (n.unused) document.getElementById('tUnused').checked = true;
    else document.getElementById('tSheets').checked = true;
    rebuild(false);
  }
  showDetail(id);
  network.selectNodes([id]);
  network.focus(id, { scale: 1.0, animation: true });
}

network.on('click', function (p) { if (p.nodes.length) showDetail(p.nodes[0]); });
network.on('doubleClick', function (p) { if (p.nodes.length) trace(p.nodes[0]); });

// ---------- sidebar list ----------
function renderList() {
  var q = document.getElementById('search').value.trim().toLowerCase(), f = document.getElementById('filter').value;
  var items = ALL.filter(function (n) {
    if (f === 'sheets' && (isField(n) || n.kind === 'Table')) return false;
    if (f === 'Table' && n.kind !== 'Table') return false;
    if (f === 'unused' && !n.unused) return false;
    if (f === 'flagged' && !(n.flags && n.flags.length && isField(n))) return false;
    if (f === 'dup' && !n.dup) return false;
    if (f === 'Group' && n.kind !== 'Group' && n.kind !== 'Bin') return false;
    if (['Calculated', 'Parameter', 'Raw'].indexOf(f) >= 0 && n.kind !== f) return false;
    if (!q) return true;
    return (n.label + ' ' + (n.formula || '') + ' ' + (n.ds || '') + ' ' + (n.physical || '') + ' ' + (n.sql || '') + ' ' + (n.table || '')).toLowerCase().indexOf(q) >= 0;
  });
  var el = document.getElementById('list');
  el.innerHTML = items.slice(0, 600).map(function (n) {
    return '<div class="item' + (n.unused ? ' unused' : '') + '" data-id="' + n.id + '"><i class="dot" style="background:' + (COLORS[n.kind] || '#ccc') + '"></i>' +
           '<span class="nm" title="' + esc(n.label) + '">' + esc(n.label) + '</span><span class="tag">' + esc(n.kind === 'Calculated' ? (n.calcKind || 'Calc') : (n.kind === 'Table' ? n.tableType : n.kind)) + '</span></div>';
  }).join('');
  document.getElementById('count').textContent = items.length + ' item' + (items.length === 1 ? '' : 's') + (items.length > 600 ? ' (showing first 600)' : '');
  el.querySelectorAll('.item').forEach(function (it) { it.onclick = function () { focusNode(+it.dataset.id); }; });
}
document.getElementById('search').addEventListener('input', renderList);
document.getElementById('filter').addEventListener('change', renderList);

rebuild();
renderList();
</script>
</body>
</html>
"""
