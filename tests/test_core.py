"""Tests for the TWB Inspector engine. Run from the project folder:  python -m unittest -v"""

import os
import shutil
import tempfile
import unittest
import zipfile

from twb_inspector.core import (
    STATUS_DASHBOARD, STATUS_INDIRECT, STATUS_SHEET, STATUS_UNUSED,
    analyze_workbook, normalize_formula, scan_formula,
)

HERE = os.path.dirname(os.path.abspath(__file__))
SAMPLES = os.path.join(HERE, "samples")
SAMPLE = os.path.join(SAMPLES, "sample_edge_cases.twb")


def by_name(lineage, name):
    return next(f for f in lineage.fields.values() if f.name == name)


class FormulaScanTest(unittest.TestCase):
    def test_refs_ignore_strings_and_comments(self):
        segs = scan_formula('IF [A] > 1 THEN "[B]" END // [C]\n + [D].[E]')
        refs = [(s[1], s[2]) for s in segs if s[0] == "ref"]
        self.assertEqual(refs, [(None, "A"), ("D", "E")])

    def test_escaped_bracket(self):
        segs = scan_formula("[Calc]]x] + 1")
        self.assertEqual([s[2] for s in segs if s[0] == "ref"], ["Calc]x"])

    def test_normalize_ignores_spacing_case_and_comments(self):
        self.assertEqual(normalize_formula("{ FIXED [Region] : sum([Sales]) } // note"),
                         normalize_formula("{FIXED [Region]:SUM([Sales])}"))


class WorkbookAnalysisTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.L = analyze_workbook(SAMPLE)

    def test_similar_ids_are_not_confused(self):
        # Calculation_1 must not be replaced inside Calculation_12 / _13
        self.assertEqual(by_name(self.L, "Twelve").formula_friendly.split("//")[0].strip(), "[Base] + [Profit]")
        self.assertIn("[Twelve] > [Target]", by_name(self.L, "Top").formula_friendly)

    def test_intermediate_calculation_is_in_lineage(self):
        base = by_name(self.L, "Base")
        self.assertEqual(base.status, STATUS_INDIRECT)
        self.assertIn("Sheet A", base.sheets_all)
        self.assertIn("Main", base.dashboards_all)

    def test_comment_reference_is_not_a_dependency(self):
        twelve = by_name(self.L, "Twelve")
        self.assertNotIn("Region", {self.L.fields[d].name for d in twelve.deps})

    def test_usage_statuses(self):
        self.assertEqual(by_name(self.L, "Top").status, STATUS_SHEET)
        self.assertEqual(by_name(self.L, "Target").status, STATUS_DASHBOARD)
        self.assertEqual(by_name(self.L, "Unused Param").status, STATUS_UNUSED)
        self.assertEqual(by_name(self.L, "Orphan").status, STATUS_UNUSED)
        self.assertEqual(by_name(self.L, "Unused Col").status, STATUS_UNUSED)

    def test_duplicates_and_flags(self):
        self.assertEqual(len(self.L.duplicate_groups), 1)
        self.assertIn("LOD expression", by_name(self.L, "Orphan").flags)

    def test_cross_datasource_reference(self):
        uses = by_name(self.L, "Uses odd")
        self.assertEqual({self.L.fields[d].name for d in uses.deps}, {"Odd ] name", "Returned"})
        self.assertEqual(self.L.fields[next(d for d in uses.deps if self.L.fields[d].name == "Returned")].datasource,
                         "federated.returns")

    def test_dashboard_and_hidden_sheet(self):
        self.assertEqual(self.L.worksheets["Sheet A"].dashboards, {"Main"})
        self.assertTrue(self.L.worksheets["Sheet A"].hidden)
        self.assertEqual(self.L.worksheets["Sheet B"].dashboards, set())

    def test_twbx_gives_same_result(self):
        tmp = tempfile.mkdtemp()
        try:
            twbx = os.path.join(tmp, "packed.twbx")
            with zipfile.ZipFile(twbx, "w") as z:
                z.write(SAMPLE, "packed.twb")
                z.writestr("Data/extract.hyper", b"")
            L2 = analyze_workbook(twbx)
            self.assertEqual(L2.summary(), self.L.summary())
        finally:
            shutil.rmtree(tmp)


class DataSourceTablesTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.L = analyze_workbook(os.path.join(SAMPLES, "sample_tables.twb"))
        cls.T = {t.name: t for t in cls.L.tables}

    def test_all_tables_found_and_extract_copy_ignored(self):
        self.assertEqual(set(self.T), {"Orders", "Customers", "Targets"})

    def test_connection_details(self):
        t = self.T["Orders"]
        self.assertEqual((t.physical, t.connection_type, t.server, t.database, t.live_or_extract),
                         ("[dbo].[Orders]", "sqlserver", "sql01.company.local", "DWH", "Extract"))
        self.assertEqual(t.logical_table, "Orders")

    def test_join_relationship_and_custom_sql(self):
        self.assertIn("Left join of Orders, Customers on [Orders].[CustomerID] = [Customers].[CustomerID]",
                      self.T["Customers"].joins)
        self.assertIn("Orders ↔ Targets on [Region] = [region (Targets)]", self.T["Targets"].relationships)
        self.assertEqual(self.T["Targets"].table_type, "Custom SQL")
        self.assertIn("FROM dbo.targets", self.T["Targets"].sql)

    def test_columns_linked_to_tables(self):
        names = {t: sorted(self.L.fields[k].name for k in tb.columns) for t, tb in self.T.items()}
        self.assertEqual(names["Customers"], ["CustomerID (Customers)", "Region"])
        self.assertEqual(by_name(self.L, "target").source_table, "Targets")
        self.assertEqual(by_name(self.L, "Order ID").source_column, "OrderID")

    def test_join_keys_are_not_unused(self):
        from twb_inspector.core import STATUS_JOIN
        self.assertEqual(by_name(self.L, "CustomerID (Customers)").status, STATUS_JOIN)
        self.assertEqual(by_name(self.L, "region (Targets)").status, STATUS_JOIN)
        self.assertEqual(by_name(self.L, "Order ID").status, STATUS_UNUSED)


class DemoWorkbookTest(unittest.TestCase):
    """The demo in examples/ is used in the README; keep it meaningful."""

    @classmethod
    def setUpClass(cls):
        root = os.path.dirname(HERE)
        cls.L = analyze_workbook(os.path.join(root, "examples", "Coffee Shop Sales (demo).twb"))

    def test_summary(self):
        s = self.L.summary()
        self.assertEqual((s["Tables"], s["Custom SQL queries"], s["Calculated fields"]), (4, 1, 10))
        self.assertEqual((s["Unused calculated fields"], s["Unused parameters"], s["Duplicate calculation groups"]),
                         (2, 1, 1))

    def test_kind_inherited_from_table_calculation(self):
        self.assertEqual(by_name(self.L, "Is Top N Product").calc_kind, "Table calculation")


class OutputsTest(unittest.TestCase):
    def test_excel_and_html_are_written(self):
        from twb_inspector.report import process_workbook
        tmp = tempfile.mkdtemp()
        try:
            res = process_workbook(SAMPLE, tmp)
            self.assertTrue(os.path.getsize(res["excel"]) > 0)
            with open(res["html"], encoding="utf-8") as fh:
                html = fh.read()
            self.assertIn("vis.Network", html)
            self.assertNotIn("/*__DATA__*/", html)
            with zipfile.ZipFile(res["excel"]) as z:          # .xlsx is a zip; no extra package needed
                book = z.read("xl/workbook.xml").decode("utf-8")
            for sheet in ("Summary", "Data Sources &amp; Tables", "All Fields", "Unused"):
                self.assertIn(f'name="{sheet}"', book)
        finally:
            shutil.rmtree(tmp)


if __name__ == "__main__":
    unittest.main()
