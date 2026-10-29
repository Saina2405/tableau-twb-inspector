"""
Writes formatted Excel workbooks from pandas DataFrames (one sheet per DataFrame).
"""

from typing import Dict, List

import pandas as pd

DEFAULT_WIDTHS = [15, 12, 15, 50, 20, 25, 30]


def write_sheets(sheets: List[Dict], excel_path: str) -> str:
    """
    Create an Excel file with one formatted sheet per entry in ``sheets``.

    Each entry is a dict with:
      df          pandas DataFrame to write (required)
      name        sheet name
      widths      list of column widths
      color       header fill colour (default '#B7DEE8')
      autofilter  True to add filter buttons to the header row
      highlight   {'column': <column name>, 'rules': {cell text: fill colour}}
      wrap        False to turn off text wrapping (e.g. a summary sheet)
      footer      footer text for printing
    """
    with pd.ExcelWriter(excel_path, engine="xlsxwriter") as writer:
        for cfg in sheets:
            df = cfg["df"]
            name = cfg.get("name", "Sheet")
            widths = cfg.get("widths", DEFAULT_WIDTHS)

            df.to_excel(writer, index=False, sheet_name=name)
            book = writer.book
            ws = writer.sheets[name]

            header_fmt = book.add_format({
                "bold": True, "bg_color": cfg.get("color", "#B7DEE8"), "border": 1,
                "align": "center", "valign": "vcenter", "font_size": 11,
            })
            data_fmt = book.add_format({
                "border": 1, "text_wrap": cfg.get("wrap", True), "valign": "top", "font_size": 10,
            })

            for col, width in enumerate(widths[:len(df.columns)]):
                ws.set_column(col, col, width, data_fmt)
            for col, value in enumerate(df.columns):
                ws.write(0, col, value, header_fmt)

            last_row = max(len(df), 1)
            last_col = max(len(df.columns) - 1, 0)

            if cfg.get("autofilter") and len(df.columns):
                ws.autofilter(0, 0, last_row, last_col)

            highlight = cfg.get("highlight")
            if highlight and highlight.get("column") in df.columns and len(df):
                col = list(df.columns).index(highlight["column"])
                for text, fill in highlight.get("rules", {}).items():
                    ws.conditional_format(1, col, last_row, col, {
                        "type": "cell", "criteria": "==", "value": f'"{text}"',
                        "format": book.add_format({"bg_color": fill}),
                    })

            # Print setup: A4 landscape, one page wide, header row repeated
            ws.set_paper(9)
            ws.set_landscape()
            ws.fit_to_pages(1, 0)
            ws.set_margins(left=0.5, right=0.5, top=0.75, bottom=0.75)
            ws.set_footer(f"&C{cfg.get('footer', 'Page &P of &N')}")
            ws.repeat_rows(0)
            ws.freeze_panes(1, 0)
    return excel_path
