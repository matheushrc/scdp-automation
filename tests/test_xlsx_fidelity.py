import tempfile
import unittest
from copy import copy
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook

from scdp_automation import xlsx_output
from tests.test_xlsx_output import make_trip
from tests.xlsx_fixtures import reference_fixture


class FidelityTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.reference = self.root / "reference.xlsx"
        reference_fixture(self.reference)

    def test_fresh_extraction_preserves_template_sheets_without_old_trips(self):
        template = self.root / "template.xlsx"
        xlsx_output.import_reference_workbook(self.reference, template)
        before = template.read_bytes()
        output = self.root / "fresh.xlsx"
        with patch.object(xlsx_output, "DEFAULT_REFERENCE", template):
            xlsx_output.publish_workbook([make_trip("888888/26")], output)

        self.assertEqual(template.read_bytes(), before)
        workbook = load_workbook(output)
        source = load_workbook(template)
        self.addCleanup(workbook.close)
        self.addCleanup(source.close)
        self.assertEqual(workbook["BASE VIAGENS"]["A2"].value, "888888/26")
        self.assertEqual(workbook["BASE VIAGENS"].max_row, 2)
        self.assertIsNone(workbook["BASE VIAGENS"]["Q2"].value)
        for name in ("APOIO", "RESUMO GASTOS"):
            actual, expected = workbook[name], source[name]
            self.assertEqual(actual.max_row, expected.max_row)
            self.assertEqual(actual.max_column, expected.max_column)
            self.assertEqual(
                set(map(str, actual.merged_cells)), set(map(str, expected.merged_cells))
            )
            for row in expected:
                for cell in row:
                    copied = actual[cell.coordinate]
                    self.assertEqual(
                        copied.value, cell.value, f"{name}!{cell.coordinate}"
                    )
                    self.assertEqual(
                        copied._style, cell._style, f"{name}!{cell.coordinate}"
                    )

    def test_fresh_extraction_does_not_reuse_matching_template_trip_metadata(self):
        template = self.root / "template.xlsx"
        xlsx_output.import_reference_workbook(self.reference, template)
        workbook = load_workbook(template)
        workbook["BASE VIAGENS"]["O2"] = "Old template description"
        workbook["BASE VIAGENS"]["N2"] = "Old verification"
        workbook.save(template)
        workbook.close()
        output = self.root / "fresh.xlsx"
        with patch.object(xlsx_output, "DEFAULT_REFERENCE", template):
            xlsx_output.publish_workbook([make_trip("999001/26")], output)
        workbook = load_workbook(output)
        self.addCleanup(workbook.close)
        for column in ("N", "O", "Q", "R"):
            self.assertIsNone(workbook["BASE VIAGENS"][f"{column}2"].value)

    def test_import_fills_all_three_sheets_and_preserves_original_presentation(self):
        path = self.root / "rebuilt.xlsx"
        before = self.reference.read_bytes()
        xlsx_output.import_reference_workbook(self.reference, path)
        self.assertEqual(self.reference.read_bytes(), before)
        workbook = load_workbook(path)
        source = load_workbook(self.reference)
        self.addCleanup(workbook.close)
        self.addCleanup(source.close)
        self.assertEqual(workbook.loaded_theme, source.loaded_theme)
        self.assertEqual(
            workbook.sheetnames, ["BASE VIAGENS", "APOIO", "RESUMO GASTOS"]
        )
        base = workbook["BASE VIAGENS"]
        self.assertEqual(base.max_row, 3)
        self.assertEqual(
            [base.cell(r, 17).value for r in (2, 3)], ["AGRONOMIA", "PPGEL +"]
        )
        support = workbook["APOIO"]
        row = next(
            r
            for r in range(2, support.max_row + 1)
            if support.cell(r, 1).value == "AGRONOMIA"
        )
        self.assertEqual(
            [support.cell(row, c).value for c in (4, 5, 7, 8)], [400, 600, 200, 150]
        )
        self.assertEqual(
            support.cell(row, 6).value, '=IF(AND(ISNUMBER(D3),ISNUMBER(E3)),D3+E3,"")'
        )
        summary = workbook["RESUMO GASTOS"]
        original = source["RESUMO GASTOS"]
        self.assertEqual(
            set(map(str, summary.merged_cells)), set(map(str, original.merged_cells))
        )
        for coordinate in ("B2", "C2", "B5", "C7", "F8", "J28", "M41", "B58"):
            with self.subTest(coordinate=coordinate):
                self.assertEqual(
                    copy(summary[coordinate].font), copy(original[coordinate].font)
                )
                self.assertEqual(
                    copy(summary[coordinate].fill), copy(original[coordinate].fill)
                )
                for side in ("left", "right", "top", "bottom", "diagonal"):
                    actual = getattr(summary[coordinate].border, side)
                    expected = getattr(original[coordinate].border, side)
                    self.assertEqual(
                        (actual.style, actual.color) if actual else (None, None),
                        (expected.style, expected.color) if expected else (None, None),
                    )
                self.assertEqual(
                    copy(summary[coordinate].alignment),
                    copy(original[coordinate].alignment),
                )
                self.assertEqual(
                    summary[coordinate].number_format,
                    original[coordinate].number_format,
                )
        self.assertEqual(summary.column_dimensions["B"].width, 44)
        self.assertEqual(summary.row_dimensions[5].height, 32)
        self.assertIn("SUMIFS", summary["F8"].value)
        self.assertEqual(
            summary["G8"].value, '=IF(AND(ISNUMBER(E8),ISNUMBER(F8)),E8-F8,"Pendente")'
        )
        self.assertIn("Apoio", summary["J8"].value)
        self.assertEqual(
            summary["M8"].value, '=IF(AND(ISNUMBER(G8),ISNUMBER(K8)),G8+K8,"Pendente")'
        )
        self.assertIn('COUNTIFS(ViagensC,"Cancelada",ViagensN,""', summary["F68"].value)

    def test_refresh_preserves_budget_transport_and_classification(self):
        current = self.root / "current.xlsx"
        candidate = self.root / "candidate.xlsx"
        xlsx_output.import_reference_workbook(self.reference, current)
        workbook = load_workbook(current)
        workbook["APOIO"]["D3"] = 12345
        workbook["APOIO"]["G3"] = 321
        workbook["APOIO"]["H3"] = 123
        workbook.save(current)
        workbook.close()
        before = current.read_bytes()
        xlsx_output.build_candidate(
            [make_trip("999002/26-1C"), make_trip("999001/26"), make_trip("999003/26")],
            current,
            candidate,
        )
        self.assertEqual(current.read_bytes(), before)
        refreshed = load_workbook(candidate)
        self.addCleanup(refreshed.close)
        self.assertEqual(
            [refreshed["BASE VIAGENS"].cell(r, 17).value for r in (2, 3, 4)],
            ["PPGEL +", "AGRONOMIA", None],
        )
        self.assertEqual(
            [refreshed["APOIO"].cell(3, c).value for c in (4, 7, 8)], [12345, 321, 123]
        )

    def test_summary_groups_supplements_without_double_counting_budget(self):
        path = self.root / "rebuilt.xlsx"
        xlsx_output.import_reference_workbook(self.reference, path)
        workbook = load_workbook(path)
        self.addCleanup(workbook.close)
        formula = workbook["RESUMO GASTOS"]["F28"].value
        self.assertIn("SUMIFS(", formula)
        self.assertIn("SUMPRODUCT(", formula)
        self.assertIn("ApoioA", formula)
        category_formula = workbook["APOIO"]["L21"].value
        self.assertIn("ViagensC", category_formula)
        self.assertIn("ViagensM", category_formula)
        self.assertNotIn("SUMIFS", workbook["RESUMO GASTOS"]["E28"].value)

    def test_manual_summary_is_preserved_but_bad_transport_prevents_publication(self):
        path = self.root / "rebuilt.xlsx"
        xlsx_output.import_reference_workbook(self.reference, path)
        for coordinate, value in (("F8", "=0"), ("J8", 0)):
            with self.subTest(coordinate=coordinate):
                workbook = load_workbook(path)
                workbook["RESUMO GASTOS"][coordinate] = value
                xlsx_output._check_workbook_structure(workbook)
                self.assertEqual(workbook["RESUMO GASTOS"][coordinate].value, value)
                workbook.close()
        workbook = load_workbook(path)
        workbook["APOIO"]["G3"] = "inválido"
        with self.assertRaisesRegex(ValueError, "numérico"):
            xlsx_output._check_workbook_structure(workbook)
        workbook.close()

    def test_default_first_publish_uses_reference_and_available_codes(self):
        path = self.root / "fresh.xlsx"
        with patch.object(xlsx_output, "DEFAULT_REFERENCE", self.reference):
            xlsx_output.publish_workbook([make_trip("999001/26")], path)
        workbook = load_workbook(path)
        self.addCleanup(workbook.close)
        self.assertEqual(workbook["BASE VIAGENS"]["Q2"].value, "AGRONOMIA")
        self.assertEqual(workbook["RESUMO GASTOS"]["B5"].value, "CURSOS DE GRADUAÇÃO")

    def test_reference_cannot_be_used_as_output(self):
        before = self.reference.read_bytes()
        with self.assertRaisesRegex(ValueError, "referência"):
            xlsx_output.import_reference_workbook(self.reference, self.reference)
        self.assertEqual(self.reference.read_bytes(), before)

    def test_new_first_row_does_not_inherit_previous_classification(self):
        current = self.root / "current.xlsx"
        candidate = self.root / "candidate.xlsx"
        xlsx_output.import_reference_workbook(self.reference, current)
        xlsx_output.build_candidate(
            [make_trip("999003/26"), make_trip("999002/26-1C"), make_trip("999001/26")],
            current,
            candidate,
        )
        workbook = load_workbook(candidate)
        self.addCleanup(workbook.close)
        self.assertEqual(
            [workbook["BASE VIAGENS"].cell(r, 17).value for r in (2, 3, 4)],
            [None, "PPGEL +", "AGRONOMIA"],
        )

    def test_invalid_rateio_cannot_silently_drop_or_duplicate_spending(self):
        path = self.root / "rebuilt.xlsx"
        xlsx_output.import_reference_workbook(self.reference, path)
        workbook = load_workbook(path)
        self.addCleanup(workbook.close)
        workbook["APOIO"]["K21"] = 0.1
        with self.assertRaisesRegex(ValueError, "100%"):
            xlsx_output._check_workbook_structure(workbook)

    def test_split_budget_without_trips_still_requires_rateio(self):
        path = self.root / "rebuilt.xlsx"
        xlsx_output.import_reference_workbook(self.reference, path)
        workbook = load_workbook(path)
        self.addCleanup(workbook.close)
        for row in range(2, workbook["BASE VIAGENS"].max_row + 1):
            workbook["BASE VIAGENS"].cell(row, 17).value = None
        for col in (9, 10, 11):
            workbook["APOIO"].cell(21, col).value = None
        workbook["APOIO"]["E21"] = 900
        with self.assertRaisesRegex(ValueError, "rateio"):
            xlsx_output._check_workbook_structure(workbook)
