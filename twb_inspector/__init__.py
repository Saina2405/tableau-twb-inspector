"""
TWB Inspector: find unused fields, duplicate calculations, lineage and
data source tables in Tableau workbooks (.twb / .twbx).

    from twb_inspector import analyze_workbook, process_workbook

    lineage = analyze_workbook("Sales.twbx")        # analysis only
    process_workbook("Sales.twbx", "outputs")       # + Excel audit and HTML diagram
"""

__version__ = "1.1.0"
APP_NAME = "TWB Inspector"

from .core import analyze_workbook, find_workbooks, WorkbookLineage  # noqa: E402
from .report import process_workbook, build_tables, write_excel, write_html  # noqa: E402

__all__ = ["analyze_workbook", "find_workbooks", "WorkbookLineage", "process_workbook",
           "build_tables", "write_excel", "write_html", "__version__", "APP_NAME"]
