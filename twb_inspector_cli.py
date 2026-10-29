"""
TWB Inspector - command line.

Usage
-----
  python twb_inspector_cli.py
      -> analyses every .twb/.twbx in the 'inputs' folder, writes to 'outputs'

  python twb_inspector_cli.py path/to/book.twbx other.twb some_folder -o my_outputs
  python twb_inspector_cli.py --no-html --no-open
  python twb_inspector_cli.py --help

For each workbook it writes:
  <name>_Calculations_audit.xlsx   Summary, Data Sources & Tables, All Fields, Unused,
                                   Duplicate Calcs, Review Flags, Dependencies, Worksheets
  <name>_lineage_diagram.html      interactive lineage diagram (works offline)
"""

import argparse
import os
import sys
import webbrowser

from twb_inspector import __version__, find_workbooks, process_workbook

HERE = os.path.dirname(os.path.abspath(__file__))


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        prog="twb_inspector_cli.py",
        description="TWB Inspector: unused fields, lineage and data sources for Tableau workbooks.")
    parser.add_argument("--version", action="version", version=f"TWB Inspector {__version__}")
    parser.add_argument("workbooks", nargs="*",
                        help=".twb/.twbx files or folders (default: the 'inputs' folder)")
    parser.add_argument("-o", "--output", default=os.path.join(HERE, "outputs"), help="output folder")
    parser.add_argument("--no-excel", action="store_true", help="skip the Excel audit")
    parser.add_argument("--no-html", action="store_true", help="skip the HTML lineage diagram")
    parser.add_argument("--no-open", action="store_true", help="do not open the diagram in the browser")
    args = parser.parse_args(argv)

    targets = args.workbooks or [os.path.join(HERE, "inputs")]
    files = []
    for t in targets:
        if os.path.isdir(t):
            files.extend(find_workbooks(t))
        elif os.path.isfile(t):
            files.append(t)
        else:
            print(f"Not found: {t}")
    if not files:
        print("No .twb or .twbx files found. Put a workbook in the 'inputs' folder or pass a path.")
        return 1

    failures = 0
    for path in files:
        print(f"\nProcessing {os.path.basename(path)} ...")
        try:
            res = process_workbook(path, args.output, excel=not args.no_excel, html=not args.no_html)
        except Exception as exc:
            failures += 1
            print(f"  ERROR: {exc}")
            continue
        for k, v in res["lineage"].summary().items():
            print(f"  {k:<32}{v}")
        for w in res["lineage"].warnings:
            print(f"  WARNING: {w}")
        if res["excel"]:
            print(f"  Excel:   {res['excel']}")
        if res["html"]:
            print(f"  Diagram: {res['html']}")
            if not args.no_open:
                webbrowser.open("file:///" + os.path.realpath(res["html"]).replace("\\", "/"))

    print(f"\nDone. {len(files) - failures} processed, {failures} failed.")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
