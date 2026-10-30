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

def unused_advice(lineage: WorkbookLineage, f: Field) -> str:
    """What to do with an unused field, depending on its type, where it is defined and whether it is hidden."""
    published = f.defined_in == "Published data source"
    if f.kind == "Raw":
        if f.hidden:
            return "Already hidden. Nothing to do."
        if published:
            return ("Not used in this workbook. It comes from a published data source, so other workbooks may "
                    "need it: hide it in this workbook to keep the Data pane clean.")
        return "Not used anywhere. Hide it in Tableau to keep the Data pane clean and extracts smaller."
    if published:
        return "Defined in the published data source, so other workbooks may use it. Nothing to change here."
    if f.used_by:
        return "Only used by other unused calculations. Review together with: " + _names(lineage, f.used_by)
    if f.hidden:
        return "Hidden and unused: probably safe to delete after checking."
    return "Nothing uses this. Candidate to delete after checking."


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
            "Defined_In": f.defined_in,
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
        advice = unused_advice(lineage, f)
        unused_rows.append({
            "Field_Name": f.name, "Type": f.kind, "Hidden": "Yes" if f.hidden else "",
            "Defined_In": f.defined_in, "Datasource": f.datasource_caption,
            "Source_Table": f.source_table,
            "Calculation": f.formula_friendly if f.kind != "Parameter" else f.param_value,
            "Used_By": _names(lineage, f.used_by), "Advice": advice, "Field_ID": f.internal_name,
        })
    unused = pd.DataFrame(unused_rows, columns=["Field_Name", "Type", "Hidden", "Defined_In", "Datasource",
                                                "Source_Table", "Calculation",
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
            "Published_On": " / ".join(x for x in (t.server, t.site, t.project) if x),
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
        "Datasource", "Published_On", "Logical_Table", "Table", "Physical_Table", "Table_Type", "Connection_Type", "Server",
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
         "widths": [24, 26, 18, 22, 26, 14, 14, 22, 30, 10, 10, 9, 9, 30, 40, 40, 50, 30, 50],
         "color": "#FFE699", "autofilter": True, "footer": "Data sources and tables"},
        {"df": t["fields"], "name": "All Fields",
         "widths": [28, 12, 18, 16, 10, 10, 22, 18, 18, 60, 14, 30, 30, 30, 26, 10, 30, 30, 26, 10, 10, 10, 8, 30, 40, 30],
         "color": "#fff0b3", "autofilter": True, "highlight": status_hl, "footer": "All fields"},
        {"df": t["unused"], "name": "Unused", "widths": [28, 12, 8, 18, 22, 18, 50, 30, 60, 30],
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
    """Nodes and edges for the HTML diagram. Levels (columns) are computed in the browser."""
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
            "definedIn": f.defined_in, "description": f.description, "flags": f.flags,
            "dup": f.duplicate_group, "sheets": sorted(f.sheets_all, key=str.lower),
            "dashboards": sorted(f.dashboards_all, key=str.lower), "fieldId": f.internal_name,
            "table": f.source_table, "sourceColumn": f.source_column, "tableNode": None,
        })

    table_node: Dict[int, int] = {}
    for t in sorted(lineage.tables, key=lambda t: (t.datasource_caption.lower(), t.name.lower())):
        table_node[id(t)] = len(nodes)
        used = sum(1 for k in t.columns if lineage.fields[k].status != STATUS_UNUSED)
        flags = list(t.notes)
        if t.columns and not used:
            flags.append("No column of this table is used in the workbook")
        nodes.append({
            "id": len(nodes), "label": t.name, "kind": "Table", "ds": t.datasource_caption,
            "tableType": t.table_type, "physical": t.physical, "logical": t.logical_table,
            "connection": t.connection_type, "server": t.server, "database": t.database, "schema": t.schema,
            "site": t.site, "project": t.project, "publishedSource": t.published_source,
            "liveOrExtract": t.live_or_extract, "joins": t.joins, "relationships": t.relationships,
            "sql": t.sql, "columns": len(t.columns), "columnsUsed": used, "flags": flags,
            "unused": bool(t.columns) and not used,
        })
        for k in t.columns:
            nodes[ids[k]]["tableNode"] = table_node[id(t)]

    ws_ids: Dict[str, int] = {}
    for w in sorted(lineage.worksheets.values(), key=lambda w: w.name.lower()):
        ws_ids[w.name] = len(nodes)
        nodes.append({"id": len(nodes), "label": w.name, "kind": "Worksheet",
                      "dashboards": sorted(w.dashboards, key=str.lower), "hidden": w.hidden,
                      "flags": [] if w.dashboards else ["Not on any dashboard"]})
    db_ids: Dict[str, int] = {}
    for d in sorted(lineage.dashboards.values(), key=lambda d: (d.is_story, d.name.lower())):
        db_ids[d.name] = len(nodes)
        nodes.append({"id": len(nodes), "label": d.name, "kind": "Story" if d.is_story else "Dashboard",
                      "sheets": sorted(d.sheets, key=str.lower), "flags": []})

    edges = []
    for t in lineage.tables:
        for k in t.columns:
            edges.append({"from": table_node[id(t)], "to": ids[k]})
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
    --bg: #f4f5f7; --panel: #ffffff; --text: #1f2328; --muted: #5b6573; --line: #d8dee4;
    --accent: #0b6bcb; --navy: #14213d; --bad: #c53030; --warn: #975a16;
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
  #detail { position: relative; width: 390px; border-left: 1px solid var(--line); border-right: 0; overflow: auto; padding: 14px; display: none; }
  #detail.open { display: block; }
  .section { padding: 12px; border-bottom: 1px solid var(--line); }
  .section h3 { margin: 0 0 8px; font-size: 12px; text-transform: uppercase; letter-spacing: .04em; color: var(--muted); }
  input[type=search], select { width: 100%; padding: 7px 9px; border: 1px solid var(--line); border-radius: 6px; font-size: 13px; background: #fff; }
  select { margin-top: 6px; }
  .hint { font-size: 12px; color: var(--muted); margin-top: 6px; }
  label.toggle { display: flex; align-items: center; gap: 6px; font-size: 13px; margin: 5px 0; cursor: pointer; }
  label.toggle.sub { margin-left: 22px; }
  label.toggle .cnt { margin-left: auto; font-size: 12px; color: var(--muted); }
  label.toggle.disabled { opacity: .45; cursor: default; }
  #list { flex: 1; overflow: auto; }
  .grp { position: sticky; top: 0; background: #f7f8fa; padding: 6px 12px; font-size: 11px; font-weight: 600; text-transform: uppercase;
         letter-spacing: .04em; color: var(--muted); border-bottom: 1px solid var(--line); }
  .item { padding: 6px 12px; border-bottom: 1px solid #f0f2f4; cursor: pointer; font-size: 13px; display: flex; gap: 8px; align-items: center; }
  .item:hover, .item.active { background: #eef5ff; }
  .item .dot { width: 10px; height: 10px; border-radius: 3px; flex: none; }
  .item .nm { flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .item .tag { font-size: 11px; color: var(--muted); white-space: nowrap; }
  .item.unused .nm { color: var(--muted); text-decoration: line-through; }
  #graph { flex: 1; position: relative; background: #fff; min-width: 0; }
  #network { position: absolute; inset: 0; }
  .emptymsg { position: absolute; top: 45%; left: 50%; transform: translate(-50%, -50%); display: none; font-size: 15px;
              color: var(--muted); background: var(--panel); border: 1px dashed var(--line); border-radius: 8px; padding: 14px 18px; }
  .emptymsg.on { display: block; }
  .toolbar { position: absolute; top: 10px; left: 10px; z-index: 5; display: flex; gap: 6px; flex-wrap: wrap; align-items: center; }
  .toolbar .sep { width: 1px; height: 24px; background: var(--line); margin: 0 2px; }
  button { font: inherit; font-size: 13px; padding: 6px 11px; border-radius: 6px; border: 1px solid var(--line);
           background: var(--panel); color: var(--text); cursor: pointer; }
  button:hover { border-color: var(--accent); color: var(--accent); }
  button.icon { width: 32px; padding: 6px 0; text-align: center; font-weight: 600; }
  button.primary { background: var(--accent); color: #fff; border-color: var(--accent); }
  button.primary:hover { color: #fff; opacity: .9; }
  .banner { position: absolute; top: 54px; left: 10px; z-index: 5; display: none; align-items: center; gap: 10px; background: #fffbea;
            border: 1px solid #f3d27a; border-radius: 8px; padding: 6px 8px 6px 12px; font-size: 13px; }
  .banner.on { display: flex; }
  details.legend { position: absolute; top: 10px; right: 10px; z-index: 5; background: rgba(255,255,255,.96);
                   border: 1px solid var(--line); border-radius: 8px; padding: 6px 10px; font-size: 12px; max-width: 360px; }
  details.legend summary { cursor: pointer; font-weight: 600; color: var(--muted); }
  .lg { display: grid; grid-template-columns: 1fr 1fr; gap: 5px 14px; margin-top: 8px; }
  .lg span { display: inline-flex; align-items: center; gap: 6px; }
  .sw { width: 12px; height: 12px; border-radius: 3px; display: inline-block; border: 1px solid rgba(0,0,0,.2); }
  .sw.dash { background: #edf2f7; border: 1px dashed #718096; }
  #detail h2 { font-size: 18px; margin: 0 24px 6px 0; word-break: break-word; }
  #detail .close { position: absolute; right: 12px; top: 10px; border: 0; font-size: 18px; padding: 2px 8px; }
  .badges { display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 10px; }
  .badge { font-size: 12px; padding: 2px 8px; border-radius: 4px; background: var(--bg); border: 1px solid var(--line); }
  .badge.unused { background: #fff5f5; color: var(--bad); border-color: #f5c2c7; }
  .badge.ok { background: #f0fff4; color: #276749; border-color: #c6f6d5; }
  .badge.hid { background: #f3f0ff; color: #553c9a; border-color: #d6ccfa; }
  dl { margin: 0; font-size: 13px; }
  dt { color: var(--muted); font-size: 12px; margin-top: 12px; }
  dd { margin: 3px 0 0; word-break: break-word; }
  pre { background: #f6f8fa; border: 1px solid var(--line); border-radius: 6px; padding: 10px; white-space: pre-wrap;
        word-break: break-word; font-size: 12.5px; font-family: Consolas, "Courier New", monospace; margin: 4px 0 0; max-height: 320px; overflow: auto; }
  .links { display: flex; flex-wrap: wrap; gap: 4px; }
  .links a { font-size: 12px; padding: 2px 7px; border-radius: 4px; background: #eef5ff; color: var(--accent); text-decoration: none; cursor: pointer; }
  .links a:hover { text-decoration: underline; }
  ul.flags { margin: 4px 0 0; padding-left: 18px; color: var(--warn); }
  .actions { display: flex; gap: 6px; margin-top: 14px; flex-wrap: wrap; }
  .empty { color: var(--muted); font-size: 13px; }
  @media (max-width: 1000px) { aside { width: 240px; } #detail { width: 310px; } }
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
      <input type="search" id="search" placeholder="Search fields, tables, formulas…">
      <select id="filter">
        <option value="all">All items</option>
        <option value="Calculated">Calculations</option>
        <option value="Parameter">Parameters</option>
        <option value="Table">Tables</option>
        <option value="Raw">Columns (raw fields)</option>
        <option value="Group">Groups, sets &amp; bins</option>
        <option value="sheets">Worksheets &amp; dashboards</option>
        <option value="unused">Unused only</option>
        <option value="hidden">Hidden fields</option>
        <option value="flagged">With review flags</option>
        <option value="dup">Duplicate calculations</option>
      </select>
      <div class="hint" id="count"></div>
    </div>
    <div class="section">
      <h3>Show on diagram</h3>
      <label class="toggle"><input type="checkbox" id="tTables" checked> Tables <span class="cnt" id="cTables"></span></label>
      <label class="toggle"><input type="checkbox" id="tCols"> Columns <span class="cnt" id="cCols"></span></label>
      <label class="toggle sub" id="lUnusedCols"><input type="checkbox" id="tUnusedCols"> include unused <span class="cnt" id="cUnusedCols"></span></label>
      <label class="toggle"><input type="checkbox" id="tCalcs" checked> Calculations, parameters &amp; groups <span class="cnt" id="cCalcs"></span></label>
      <label class="toggle sub" id="lUnused"><input type="checkbox" id="tUnused" checked> include unused <span class="cnt" id="cUnused"></span></label>
      <label class="toggle"><input type="checkbox" id="tSheets" checked> Worksheets &amp; dashboards <span class="cnt" id="cSheets"></span></label>
      <div class="hint">With Columns off, each table stands in for its columns.</div>
    </div>
    <div id="list"></div>
  </aside>
  <section id="graph">
    <div class="toolbar">
      <button onclick="network.fit({animation:true})" title="Fit the diagram to the screen">Fit</button>
      <button class="icon" onclick="zoom(1/1.3)" title="Zoom out">−</button>
      <button class="icon" onclick="zoom(1.3)" title="Zoom in">+</button>
      <span class="sep"></span>
      <button id="dirBtn" onclick="toggleDirection()" title="Switch between left-to-right and top-to-bottom">Top → bottom</button>
      <button onclick="exportPng()" title="Download the current view as a PNG image">Export PNG</button>
      <button id="clearBtn" onclick="clearTrace()" style="display:none">Clear trace</button>
    </div>
    <div class="banner" id="focusBar">Showing only the lineage of <b id="focusName"></b> <button onclick="setFocus(false)">Back to full view</button></div>
    <div id="network"></div>
    <div id="emptyMsg" class="emptymsg">Nothing selected. Tick something under <b>Show on diagram</b>.</div>
    <details class="legend" id="legend" open>
      <summary>Legend</summary>
      <div class="lg">
        <span><i class="sw" style="background:#fbb6ce"></i>Table</span>
        <span><i class="sw" style="background:#8ab4f8"></i>Column (raw field)</span>
        <span><i class="sw" style="background:#f6ad55"></i>Calculation</span>
        <span><i class="sw" style="background:#b794f4"></i>Parameter</span>
        <span><i class="sw" style="background:#68d391"></i>Group / set / bin</span>
        <span><i class="sw" style="background:#4fd1c5"></i>Worksheet</span>
        <span><i class="sw" style="background:#4a5568"></i>Dashboard / story</span>
        <span><i class="sw dash"></i>Unused</span>
        <span><svg width="28" height="8" aria-hidden="true"><line x1="0" y1="4" x2="28" y2="4" stroke="#7a8699" stroke-width="2"/></svg>feeds into</span>
        <span><svg width="28" height="8" aria-hidden="true"><line x1="0" y1="4" x2="28" y2="4" stroke="#7a8699" stroke-width="2" stroke-dasharray="5 4"/></svg>shown on sheet</span>
      </div>
    </details>
  </section>
  <aside id="detail"></aside>
</main>

<script type="text/javascript">
var DATA = /*__DATA__*/null;
var TITLE = document.title.split(' · ')[0];
var G = DATA.graph, ALL = G.nodes;
var byId = {}; ALL.forEach(function (n) { byId[n.id] = n; });
var parents = {}, children = {};
ALL.forEach(function (n) { parents[n.id] = []; children[n.id] = []; });
G.edges.forEach(function (e) { parents[e.to].push(e.from); children[e.from].push(e.to); });

var COLORS = { Table: '#fbb6ce', Raw: '#8ab4f8', Calculated: '#f6ad55', Parameter: '#b794f4', Group: '#68d391', Bin: '#68d391',
               Worksheet: '#4fd1c5', Dashboard: '#4a5568', Story: '#2d3748' };
var BORDERS = { Table: '#b83280', Raw: '#3b78e7', Calculated: '#c05621', Parameter: '#6b46c1', Group: '#2f855a', Bin: '#2f855a',
                Worksheet: '#2c7a7b', Dashboard: '#1a202c', Story: '#1a202c' };
var isSheetKind = function (n) { return n.kind === 'Worksheet' || n.kind === 'Dashboard' || n.kind === 'Story'; };
var isField = function (n) { return !isSheetKind(n) && n.kind !== 'Table'; };
var $ = function (id) { return document.getElementById(id); };
var state = { traced: null, focus: false, focusId: null, dir: 'LR' };

// ---------- header chips ----------
(function () {
  var s = DATA.summary, el = $('chips');
  [['Published data sources', 0], ['Tables', 0], ['Calculated fields', 0], ['Parameters', 0], ['Worksheets', 0], ['Dashboards', 0],
   ['Unused calculated fields', 1], ['Unused parameters', 1], ['Duplicate calculation groups', 1], ['Hidden fields', 0]].forEach(function (p) {
    if (s[p[0]] === undefined || (p[0] === 'Published data sources' && !s[p[0]])) return;
    var c = document.createElement('span');
    c.className = 'chip' + (p[1] && s[p[0]] > 0 ? ' warn' : '');
    c.textContent = p[0] + ': ' + s[p[0]];
    if (p[0] === 'Unused calculated fields' && s['Unused workbook calculations'] !== s[p[0]])
      c.textContent += ' (' + s['Unused workbook calculations'] + ' in workbook, ' + (s[p[0]] - s['Unused workbook calculations']) + ' from published source)';
    el.appendChild(c);
  });
})();

// ---------- toggle counts ----------
(function () {
  var cnt = function (f) { return ALL.filter(f).length; };
  $('cTables').textContent = cnt(function (n) { return n.kind === 'Table'; });
  $('cCols').textContent = cnt(function (n) { return n.kind === 'Raw' && !n.unused; });
  $('cUnusedCols').textContent = '+' + cnt(function (n) { return n.kind === 'Raw' && n.unused; });
  $('cCalcs').textContent = cnt(function (n) { return isField(n) && n.kind !== 'Raw' && !n.unused; });
  $('cUnused').textContent = '+' + cnt(function (n) { return isField(n) && n.kind !== 'Raw' && n.unused; });
  $('cSheets').textContent = cnt(isSheetKind);
  if (!ALL.some(function (n) { return n.kind === 'Table'; })) { $('tCols').checked = true; }
})();

// ---------- visibility ----------
function on(id) { return $(id).checked; }
function baseVisible(n) {
  if (isSheetKind(n)) return on('tSheets');
  if (n.kind === 'Table') return on('tTables');
  if (n.kind === 'Raw') return on('tCols') && (!n.unused || on('tUnusedCols'));
  if (!on('tCalcs')) return false;                       // calculations, parameters, groups, bins
  return n.unused ? on('tUnused') : true;
}
function visible(n) {
  if (!baseVisible(n)) return false;
  return !(state.focus && state.traced && !state.traced.all[n.id]);
}
function updateToggleState() {
  $('tUnusedCols').disabled = !on('tCols');
  $('lUnusedCols').classList.toggle('disabled', !on('tCols'));
  $('tUnused').disabled = !on('tCalcs');
  $('lUnused').classList.toggle('disabled', !on('tCalcs'));
}

// ---------- node style ----------
function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
function safe(s) { return String(s).replace(/</g, '‹').replace(/>/g, '›'); }
function tagsOf(n) {
  var t = [];
  if (n.kind === 'Calculated' && n.calcKind === 'LOD') t.push('LOD');
  if (n.kind === 'Calculated' && n.calcKind === 'Table calculation') t.push('table calc');
  if (n.kind === 'Table') {
    t.push(n.columns + ' col' + (n.columns === 1 ? '' : 's') + (n.columns ? ' · ' + n.columnsUsed + ' used' : ''));
    if (n.tableType === 'Custom SQL') t.push('Custom SQL');
    if (n.publishedSource) t.push('published');
  }
  if (n.hidden && n.kind !== 'Worksheet') t.push('hidden');
  if (n.kind === 'Worksheet' && n.hidden) t.push('hidden tab');
  return t;
}
function styleNode(n, dim, level) {
  var bg = COLORS[n.kind] || '#cbd5e0', border = BORDERS[n.kind] || '#718096',
      font = (n.kind === 'Dashboard' || n.kind === 'Story') ? '#ffffff' : '#1a202c';
  var tags = tagsOf(n);
  var label = (n.kind === 'Table' ? '▦ ' : '') + safe(n.label) + (tags.length ? '\n<i>' + safe(tags.join(' · ')) + '</i>' : '');
  var o = { id: n.id, label: label, level: level, shape: n.kind === 'Parameter' ? 'ellipse' : 'box', margin: 9,
            widthConstraint: { maximum: 230 }, borderWidth: n.kind === 'Table' ? 2.5 : 1.5,
            shapeProperties: { borderDashes: false },
            font: { size: 14, face: 'Segoe UI, Arial', multi: 'html', ital: { size: 11, color: font === '#ffffff' ? '#e2e8f0' : '#4a5568' } } };
  if (n.unused) { bg = '#edf2f7'; border = '#718096'; font = '#4a5568'; o.shapeProperties = { borderDashes: [5, 4] }; }
  if (n.flags && n.flags.length && isField(n)) o.borderWidth = 3;
  if (dim) { bg = '#f7fafc'; border = '#e2e8f0'; font = '#cbd5e0'; o.font.ital.color = '#d5dbe3'; }
  o.color = { background: bg, border: border, highlight: { background: bg, border: '#0b6bcb' }, hover: { background: bg, border: '#0b6bcb' } };
  o.font.color = font;
  return o;
}

// ---------- graph for the current view ----------
var nodes = new vis.DataSet([]), edges = new vis.DataSet([]);
function hierOptions() {
  var lr = state.dir === 'LR';
  return { layout: { hierarchical: { enabled: true, direction: state.dir, sortMethod: 'directed', shakeTowards: 'roots',
                                     levelSeparation: lr ? 270 : 150, nodeSpacing: lr ? 64 : 200, treeSpacing: 90 } },
           edges: { smooth: { type: 'cubicBezier', forceDirection: lr ? 'horizontal' : 'vertical', roundness: 0.4 } } };
}
var network = new vis.Network($('network'), { nodes: nodes, edges: edges }, {
  physics: { enabled: false },
  edges: { arrows: { to: { enabled: true, scaleFactor: 0.5 } }, color: { color: '#a0aec0', highlight: '#0b6bcb', hover: '#0b6bcb', inherit: false }, width: 1.2,
           smooth: { type: 'cubicBezier', forceDirection: 'horizontal', roundness: 0.4 } },
  interaction: { hover: true, tooltipDelay: 150, navigationButtons: false, keyboard: false, hideEdgesOnDrag: true }
});
network.setOptions(hierOptions());   // merged into the options above (not replacing them)

function currentGraph() {
  var vis_ = {};
  ALL.forEach(function (n) { if (visible(n)) vis_[n.id] = true; });
  var rep = function (id) {
    if (vis_[id]) return id;
    var n = byId[id];
    if (n.kind === 'Raw' && n.tableNode !== null && vis_[n.tableNode]) return n.tableNode;   // collapsed column -> its table
    return null;
  };
  var seen = {}, es = [];
  G.edges.forEach(function (e) {
    var a = rep(e.from), b = rep(e.to);
    if (a === null || b === null || a === b) return;
    var k = a + '>' + b;
    if (seen[k]) { if (!e.use) seen[k].use = false; return; }
    seen[k] = { from: a, to: b, use: !!e.use, origFrom: e.from, origTo: e.to };
    es.push(seen[k]);
  });
  // columns: longest path over the visible edges; sheets and dashboards get their own columns
  var level = {}, indeg = {}, out = {};
  Object.keys(vis_).forEach(function (id) { level[id] = 0; indeg[id] = 0; out[id] = []; });
  es.forEach(function (e) { indeg[e.to]++; out[e.from].push(e.to); });
  var queue = Object.keys(vis_).filter(function (id) { return indeg[id] === 0; });
  while (queue.length) {
    var x = queue.shift();
    out[x].forEach(function (y) { level[y] = Math.max(level[y], level[x] + 1); if (--indeg[y] === 0) queue.push(y); });
  }
  var maxField = 0;
  Object.keys(vis_).forEach(function (id) { if (!isSheetKind(byId[id])) maxField = Math.max(maxField, level[id]); });
  Object.keys(vis_).forEach(function (id) {
    var k = byId[id].kind;
    if (k === 'Worksheet') level[id] = maxField + 1;
    if (k === 'Dashboard') level[id] = maxField + 2;
    if (k === 'Story') level[id] = maxField + 3;
  });
  return { visible: vis_, edges: es, level: level };
}

function rebuild(fit) {
  updateToggleState();
  var g = currentGraph(), t = state.traced, ns = [], es = [];
  ALL.forEach(function (n) {
    if (!g.visible[n.id]) return;
    ns.push(styleNode(n, t && !state.focus && !t.all[n.id], g.level[n.id]));
  });
  g.edges.forEach(function (e, i) {
    var on_ = !t || state.focus || (t.up[e.from] && t.up[e.to]) || (t.down[e.from] && t.down[e.to]);
    es.push({ id: i, from: e.from, to: e.to, dashes: e.use,
              color: on_ ? undefined : { color: '#edf2f7' }, width: !on_ ? 0.6 : (t ? 2 : 1.2) });
  });
  nodes.clear(); edges.clear(); nodes.add(ns); edges.add(es);
  $('emptyMsg').classList.toggle('on', ns.length === 0);
  if (fit !== false) setTimeout(function () { network.fit(); }, 30);
}
['tTables', 'tCols', 'tUnusedCols', 'tCalcs', 'tUnused', 'tSheets'].forEach(function (id) {
  $(id).addEventListener('change', function () { rebuild(); });
});

// ---------- toolbar ----------
function zoom(f) { network.moveTo({ scale: network.getScale() * f, animation: { duration: 200 } }); }
function toggleDirection() {
  state.dir = state.dir === 'LR' ? 'UD' : 'LR';
  $('dirBtn').textContent = state.dir === 'LR' ? 'Top → bottom' : 'Left → right';
  network.setOptions(hierOptions());
  rebuild();
}
function exportPng() {
  var c = document.querySelector('#network canvas');
  if (!c) return;
  var o = document.createElement('canvas'); o.width = c.width; o.height = c.height;
  var x = o.getContext('2d'); x.fillStyle = '#ffffff'; x.fillRect(0, 0, o.width, o.height); x.drawImage(c, 0, 0);
  var a = document.createElement('a');
  a.download = TITLE.replace(/[^\w\- ]+/g, '_') + '_lineage.png';
  a.href = o.toDataURL('image/png'); document.body.appendChild(a); a.click(); a.remove();
}

// ---------- lineage tracing ----------
function walk(start, map) {
  var seen = {}, stack = [start];
  while (stack.length) { var x = stack.pop(); if (seen[x]) continue; seen[x] = true; Array.prototype.push.apply(stack, map[x]); }
  return seen;
}
function reveal(n) {
  if (n.kind === 'Table') $('tTables').checked = true;
  else if (n.kind === 'Raw') {
    $('tCols').checked = true;
    if (n.unused) $('tUnusedCols').checked = true;
  }
  else if (isSheetKind(n)) $('tSheets').checked = true;
  else {
    $('tCalcs').checked = true;
    if (n.unused) $('tUnused').checked = true;
  }
}
function trace(id) {
  var up = walk(id, parents), down = walk(id, children), all = {};
  Object.keys(up).forEach(function (k) { all[k] = true; });
  Object.keys(down).forEach(function (k) { all[k] = true; });
  state.traced = { all: all, up: up, down: down, id: id };
  var n = byId[id];
  if (!baseVisible(n)) reveal(n);
  $('clearBtn').style.display = '';
  rebuild(false);
  var vis_ = Object.keys(all).map(Number).filter(function (k) { return nodes.get(k); });
  network.fit({ nodes: vis_.length ? vis_ : undefined, animation: true });
  network.selectNodes(nodes.get(id) ? [id] : []);
  showDetail(id);
}
function setFocus(onOff) {
  if (onOff && !state.traced) return;
  state.focus = onOff;
  $('focusBar').classList.toggle('on', onOff);
  if (onOff) $('focusName').textContent = byId[state.traced.id].label;
  rebuild();
  if (state.traced && nodes.get(state.traced.id)) network.selectNodes([state.traced.id]);
}
function clearTrace() {
  state.traced = null; state.focus = false;
  $('focusBar').classList.remove('on'); $('clearBtn').style.display = 'none';
  rebuild();
}

// ---------- detail panel ----------
function linkList(ids) {
  if (!ids.length) return '<span class="empty">None</span>';
  ids = ids.slice().sort(function (a, b) { return byId[a].label.localeCompare(byId[b].label); });
  return '<div class="links">' + ids.map(function (i) { return '<a data-id="' + i + '">' + esc(byId[i].label) + '</a>'; }).join('') + '</div>';
}
function nameList(arr) { return arr && arr.length ? esc(arr.join(', ')) : '<span class="empty">None</span>'; }
function showDetail(id) {
  var n = byId[id], el = $('detail');
  var fieldParents = parents[id].filter(function (i) { return byId[i].kind !== 'Worksheet'; });
  var fieldChildren = children[id].filter(function (i) { return isField(byId[i]); });
  var h = '<button class="close" onclick="closeDetail()" aria-label="Close">×</button>';
  h += '<h2>' + esc(n.label) + '</h2><div class="badges"><span class="badge">' + esc(n.kind === 'Raw' ? 'Column' : n.kind) + '</span>';
  if (n.calcKind) h += '<span class="badge">' + esc(n.calcKind) + '</span>';
  if (n.status) h += '<span class="badge ' + (n.unused ? 'unused' : 'ok') + '">' + esc(n.status) + '</span>';
  if (n.hidden) h += '<span class="badge hid">' + (n.kind === 'Worksheet' ? 'Hidden tab' : 'Hidden') + '</span>';
  if (n.definedIn === 'Published data source') h += '<span class="badge">From published data source</span>';
  if (n.inferred) h += '<span class="badge">Not declared in file</span>';
  if (n.kind === 'Table') h += '<span class="badge">' + esc(n.tableType) + '</span><span class="badge">' + esc(n.liveOrExtract) + '</span>';
  h += '</div><dl>';
  if (isField(n)) {
    h += '<dt>Data source</dt><dd>' + esc(n.ds) + (n.datatype ? ' · ' + esc(n.datatype) : '') + (n.role ? ' · ' + esc(n.role) : '') + '</dd>';
    if (n.table) h += '<dt>Source table / column</dt><dd>' + esc(n.table) + (n.sourceColumn ? ' · ' + esc(n.sourceColumn) : '') + '</dd>';
    if (n.formula) h += '<dt>Formula</dt><dd><pre>' + esc(n.formula) + '</pre></dd>';
    else if (n.kind === 'Calculated') h += '<dt>Formula</dt><dd class="empty">Defined in the published data source; not saved in this workbook</dd>';
    if (n.kind === 'Parameter') h += '<dt>Current value</dt><dd>' + esc(n.paramValue) + '</dd><dt>Allowed values</dt><dd>' + esc(n.paramDomain) + '</dd>';
    if (n.description) h += '<dt>Description</dt><dd>' + esc(n.description) + '</dd>';
    h += '<dt>Depends on</dt><dd>' + linkList(fieldParents) + '</dd>';
    h += '<dt>Used by</dt><dd>' + linkList(fieldChildren) + '</dd>';
    h += '<dt>Worksheets (incl. through other calculations)</dt><dd>' + nameList(n.sheets) + '</dd>';
    h += '<dt>Dashboards</dt><dd>' + nameList(n.dashboards) + '</dd>';
    h += '<dt>Field ID</dt><dd><code>' + esc(n.fieldId) + '</code></dd>';
  } else if (n.kind === 'Table') {
    h += '<dt>Data source</dt><dd>' + esc(n.publishedSource || n.ds) + '</dd>';
    if (n.publishedSource) h += '<dt>Published on</dt><dd>' + esc([n.server, n.site && 'site ' + n.site, n.project && 'project ' + n.project].filter(Boolean).join(' · ')) + '</dd>';
    if (n.physical) h += '<dt>Physical table</dt><dd><code>' + esc(n.physical) + '</code></dd>';
    if (!n.publishedSource) h += '<dt>Connection</dt><dd>' + esc([n.connection, n.server, n.database, n.schema].filter(Boolean).join(' · ')) + '</dd>';
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
  if (n.flags && n.flags.length) h += '<dt>Notes</dt><dd><ul class="flags">' + n.flags.map(function (f) { return '<li>' + esc(f) + '</li>'; }).join('') + '</ul></dd>';
  h += '</dl><div class="actions"><button class="primary" onclick="trace(' + id + ')">Trace lineage</button>';
  h += '<button onclick="trace(' + id + ');setFocus(true)">Show only this lineage</button>';
  if (n.formula || n.sql) h += '<button onclick="copyText(' + id + ', this)">Copy ' + (n.sql ? 'SQL' : 'formula') + '</button>';
  h += '</div>';
  el.innerHTML = h; el.classList.add('open');
  el.querySelectorAll('a[data-id]').forEach(function (a) { a.onclick = function () { focusNode(+a.dataset.id); }; });
  document.querySelectorAll('.item').forEach(function (it) { it.classList.toggle('active', +it.dataset.id === id); });
}
function closeDetail() { $('detail').classList.remove('open'); network.unselectAll(); }
function copyText(id, btn) {
  var t = byId[id].formula || byId[id].sql;
  try { navigator.clipboard.writeText(t).then(function () { btn.textContent = 'Copied'; }); }
  catch (e) { btn.textContent = 'Copy not available'; }
}
function focusNode(id) {
  var n = byId[id];
  if (state.focus && state.traced && !state.traced.all[id]) setFocus(false);
  if (!baseVisible(n)) { reveal(n); rebuild(false); }
  showDetail(id);
  if (nodes.get(id)) { network.selectNodes([id]); network.focus(id, { scale: 1.0, animation: true }); }
}
network.on('click', function (p) { if (p.nodes.length) showDetail(p.nodes[0]); });
network.on('doubleClick', function (p) { if (p.nodes.length) trace(p.nodes[0]); });

// ---------- sidebar list ----------
var GROUPS = [['Dashboard', 'Dashboards'], ['Story', 'Stories'], ['Worksheet', 'Worksheets'], ['Calculated', 'Calculations'],
              ['Parameter', 'Parameters'], ['Group', 'Groups & sets'], ['Bin', 'Bins'], ['Table', 'Tables'], ['Raw', 'Columns']];
var lastMatches = [];
function matches() {
  var q = $('search').value.trim().toLowerCase(), f = $('filter').value;
  return ALL.filter(function (n) {
    if (f === 'sheets' && !isSheetKind(n)) return false;
    if (f === 'unused' && !n.unused) return false;
    if (f === 'hidden' && !(n.hidden && n.kind !== 'Worksheet')) return false;
    if (f === 'flagged' && !(n.flags && n.flags.length && isField(n))) return false;
    if (f === 'dup' && !n.dup) return false;
    if (f === 'Group' && n.kind !== 'Group' && n.kind !== 'Bin') return false;
    if (['Calculated', 'Parameter', 'Raw', 'Table'].indexOf(f) >= 0 && n.kind !== f) return false;
    if (!q) return true;
    return [n.label, n.formula, n.ds, n.physical, n.sql, n.table].join(' ').toLowerCase().indexOf(q) >= 0;
  });
}
function renderList() {
  var items = matches(), el = $('list'), h = '', shown = 0;
  lastMatches = items;
  GROUPS.forEach(function (g) {
    var grp = items.filter(function (n) { return n.kind === g[0]; });
    if (!grp.length) return;
    h += '<div class="grp">' + g[1] + ' (' + grp.length + ')</div>';
    grp.slice(0, 400).forEach(function (n) {
      shown++;
      var tag = n.kind === 'Calculated' ? (n.calcKind || '') : (n.kind === 'Table' ? n.tableType : '');
      if (n.hidden && n.kind !== 'Worksheet') tag = (tag ? tag + ' · ' : '') + 'hidden';
      h += '<div class="item' + (n.unused ? ' unused' : '') + '" data-id="' + n.id + '"><i class="dot" style="background:' + (COLORS[n.kind] || '#ccc') + '"></i>' +
           '<span class="nm" title="' + esc(n.label) + '">' + esc(n.label) + '</span><span class="tag">' + esc(tag) + '</span></div>';
    });
  });
  el.innerHTML = h || '<div class="section empty">Nothing matches.</div>';
  $('count').textContent = items.length + ' item' + (items.length === 1 ? '' : 's') + ($('search').value ? ' · Enter opens the first' : '');
  el.querySelectorAll('.item').forEach(function (it) { it.onclick = function () { focusNode(+it.dataset.id); }; });
  // highlight matches on the diagram while searching
  if ($('search').value.trim()) network.selectNodes(items.map(function (n) { return n.id; }).filter(function (i) { return nodes.get(i); }));
  else network.unselectAll();
}
$('search').addEventListener('input', renderList);
$('search').addEventListener('keydown', function (e) {
  if (e.key === 'Enter' && lastMatches.length) {
    var first = GROUPS.map(function (g) { return lastMatches.filter(function (n) { return n.kind === g[0]; })[0]; }).filter(Boolean)[0];
    if (first) focusNode(first.id);
  }
});
$('filter').addEventListener('change', renderList);

if (ALL.length > 120) $('legend').open = false;   // more room on big workbooks
rebuild();
renderList();
</script>
</body>
</html>
"""
