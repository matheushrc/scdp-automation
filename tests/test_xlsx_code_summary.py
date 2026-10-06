import shutil
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator

from scdp_automation.xlsx_output import import_reference_workbook
from scdp_automation.xlsx_recalculate import recalculate_workbook
from tests.support.workbooks import reference_fixture


@unittest.skipUnless(shutil.which("libreoffice"), "LibreOffice necessário")
class CodeSummaryTests(unittest.TestCase):
    def test_new_category_uses_code_without_group_or_hidden_metadata(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            reference, output = root / "reference.xlsx", root / "output.xlsx"
            reference_fixture(reference)
            import_reference_workbook(reference, output)
            book = load_workbook(output)
            support, summary = book["APOIO"], book["RESUMO GASTOS"]
            self.assertNotIn("Grupo no resumo", [c.value for c in support[1]])
            for r in range(2, support.max_row + 1):
                if support.cell(r, 1).value:
                    for c in (4, 5, 7):
                        if support.cell(r, c).value is None:
                            support.cell(r, c).value = 0
            expected_budget = (
                sum(
                    support.cell(r, c).value or 0
                    for r in range(2, support.max_row + 1)
                    for c in (4, 5)
                )
                + 10000
            )
            row = support.max_row + 3
            for col, value in (
                (1, "TESTE"),
                (2, "Categoria teste"),
                (3, "SEG 1 GRADUAÇÃO"),
                (4, 10000),
                (5, 0),
                (7, 0),
            ):
                support.cell(row, col, value)
            summary["B70"] = "TESTE"
            for col in ("C", "E", "F", "G", "I", "J", "K", "M"):
                summary[f"{col}70"] = Translator(
                    summary[f"{col}9"].value, origin=f"{col}9"
                ).translate_formula(f"{col}70")
            base = book["BASE VIAGENS"]
            base.append(
                [
                    "111111/26",
                    "Pessoa teste",
                    "Concluída",
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    125,
                    None,
                    None,
                    None,
                    None,
                    None,
                    "TESTE",
                ]
            )
            book.save(output)
            book.close()
            recalculate_workbook(output)
            cached = load_workbook(output, data_only=True)
            summary = cached["RESUMO GASTOS"]
            self.assertEqual(summary["E70"].value, 10000)
            self.assertEqual(summary["F70"].value, 125)
            self.assertEqual(summary["M70"].value, 9875)
            self.assertEqual(summary["C64"].value, 0)
            self.assertEqual(summary["C68"].value, expected_budget)
            self.assertEqual(summary["F68"].value, 437)
            cached.close()

    def test_direction_subcodes_and_rateio_keep_their_totals(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            reference, output = root / "reference.xlsx", root / "output.xlsx"
            reference_fixture(reference)
            import_reference_workbook(reference, output)
            book = load_workbook(output)
            support, base = book["APOIO"], book["BASE VIAGENS"]
            self.assertNotIn("Grupo no resumo", [c.value for c in support[1]])
            row = next(
                r
                for r in range(2, support.max_row + 1)
                if support.cell(r, 1).value == "DIREÇÃO - AGAS"
            )
            support.cell(row, 4).value = 10
            support.cell(row, 5).value = 20
            support.cell(row, 7).value = 5
            base.append(
                [
                    "111111/26",
                    "Pessoa teste",
                    "Concluída",
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    75,
                    None,
                    None,
                    None,
                    None,
                    None,
                    "DIREÇÃO - AGAS",
                ]
            )
            book.save(output)
            book.close()
            recalculate_workbook(output)
            cached = load_workbook(output, data_only=True)
            summary = cached["RESUMO GASTOS"]
            self.assertEqual(summary["C41"].value, 1030)
            self.assertEqual(summary["F41"].value, 75)
            self.assertEqual(summary["J41"].value, 205)
            self.assertEqual(summary["F27"].value, 52)
            self.assertEqual(summary["F28"].value, 52)
            self.assertEqual(summary["F32"].value, 52)
            cached.close()

    def test_migration_preserves_manual_formulas_and_support_references(self):
        from scdp_automation.xlsx_output import build_candidate
        from scdp_automation.xlsx_reference import read_reference
        from tests.support.workbooks import restore_legacy_group_layout

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            reference, current, candidate = (
                root / "reference.xlsx",
                root / "current.xlsx",
                root / "candidate.xlsx",
            )
            reference_fixture(reference)
            import_reference_workbook(reference, current)
            book = load_workbook(current)
            restore_legacy_group_layout(book)
            summary = book["RESUMO GASTOS"]
            summary["E8"] = "=SUMIF(ApoioI,$B8,ApoioD)+10"
            summary["M9"] = "=COUNT(APOIO!M3:M4)"
            book.save(current)
            book.close()
            source_bytes = current.read_bytes()
            build_candidate(read_reference(reference).trips, current, candidate)
            self.assertEqual(current.read_bytes(), source_bytes)
            book = load_workbook(candidate)
            self.assertTrue(book["RESUMO GASTOS"]["E8"].value.endswith("+10"))
            self.assertEqual(book["RESUMO GASTOS"]["M9"].value, "=COUNT(APOIO!L3:L4)")
            self.assertNotIn("ApoioI", book["RESUMO GASTOS"]["E8"].value)
            book.close()
            recalculate_workbook(candidate)
            cached = load_workbook(candidate, data_only=True)
            self.assertEqual(cached["RESUMO GASTOS"]["E8"].value, 410)
            cached.close()

    def test_direction_code_and_renamed_category_enter_totals_once(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            reference, output = root / "reference.xlsx", root / "output.xlsx"
            reference_fixture(reference)
            import_reference_workbook(reference, output)
            book = load_workbook(output)
            support = book["APOIO"]
            for r in range(2, support.max_row + 1):
                if support.cell(r, 1).value:
                    for col in (4, 5, 7):
                        if support.cell(r, col).value is None:
                            support.cell(r, col).value = 0
            expected = sum(
                support.cell(r, c).value or 0
                for r in range(2, support.max_row + 1)
                for c in (4, 5)
            )
            row = next(
                r
                for r in range(2, support.max_row + 1)
                if support.cell(r, 1).value == "DIREÇÃO"
            )
            base = book["BASE VIAGENS"]
            base.append(
                [
                    "111111/26",
                    "Pessoa teste",
                    "Concluída",
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    75,
                    None,
                    None,
                    None,
                    None,
                    None,
                    "DIREÇÃO",
                ]
            )
            book["RESUMO GASTOS"]["B41"] = "DIREÇÃO"
            book.save(output)
            book.close()
            recalculate_workbook(output)
            cached = load_workbook(output, data_only=True)
            self.assertEqual(cached["RESUMO GASTOS"]["C68"].value, expected)
            self.assertEqual(cached["RESUMO GASTOS"]["F45"].value, 387)
            cached.close()
            book = load_workbook(output)
            book["APOIO"].cell(row, 2).value = "OUTROS"
            book["RESUMO GASTOS"]["B41"] = "OUTROS"
            book.save(output)
            book.close()
            recalculate_workbook(output)
            cached = load_workbook(output, data_only=True)
            self.assertEqual(cached["RESUMO GASTOS"]["C68"].value, expected)
            self.assertEqual(cached["RESUMO GASTOS"]["F45"].value, 387)
            cached.close()

    def test_migration_rejects_unsupported_references_without_changing_source(self):
        from scdp_automation.xlsx_output import build_candidate
        from scdp_automation.xlsx_reference import read_reference
        from tests.support.workbooks import restore_legacy_group_layout

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            reference, current, candidate = (
                root / "reference.xlsx",
                root / "current.xlsx",
                root / "candidate.xlsx",
            )
            reference_fixture(reference)
            for formula in (
                '=COUNTIF(ApoioI,"Agronomia")',
                "=APOIO!I3",
                "=INDEX(APOIO!I3:J3,1,1)",
            ):
                with self.subTest(formula=formula):
                    current.unlink(missing_ok=True)
                    candidate.unlink(missing_ok=True)
                    import_reference_workbook(reference, current)
                    book = load_workbook(current)
                    restore_legacy_group_layout(book)
                    book["RESUMO GASTOS"]["E70"] = formula
                    book.save(current)
                    book.close()
                    before = current.read_bytes()
                    with self.assertRaisesRegex(ValueError, "grupo antigo"):
                        build_candidate(
                            read_reference(reference).trips, current, candidate
                        )
                    self.assertEqual(current.read_bytes(), before)
