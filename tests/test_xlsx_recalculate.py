import shutil
import tempfile
import unittest
from copy import copy
from pathlib import Path

from openpyxl import load_workbook

from scdp_automation.xlsx_output import (
    _check_workbook_structure,
    import_reference_workbook,
)
from scdp_automation.xlsx_recalculate import recalculate_workbook
from tests.xlsx_fixtures import reference_fixture


@unittest.skipUnless(shutil.which("libreoffice"), "LibreOffice unavailable")
class RecalculationTests(unittest.TestCase):
    def test_visible_results_recalculate_while_original_formulas_and_styles_survive(
        self,
    ):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            reference = root / "reference.xlsx"
            output = root / "output.xlsx"
            reference_fixture(reference)
            import_reference_workbook(reference, output)
            workbook = load_workbook(output)
            font = copy(workbook["RESUMO GASTOS"]["B2"].font)
            formula = workbook["RESUMO GASTOS"]["F8"].value
            workbook.close()
            recalculate_workbook(output)
            cached = load_workbook(output, data_only=True)
            self.assertEqual(
                [
                    cached["RESUMO GASTOS"][c].value
                    for c in ("F8", "G8", "J8", "K8", "M8", "F28", "F68")
                ],
                [156, 244, 200, 400, 644, 52, 312],
            )
            self.assertEqual(cached["APOIO"]["F3"].value, 1000)
            cached.close()
            workbook = load_workbook(output)
            _check_workbook_structure(workbook)
            self.assertEqual(workbook["RESUMO GASTOS"]["F8"].value, formula)
            self.assertEqual(copy(workbook["RESUMO GASTOS"]["B2"].font), font)
            workbook["APOIO"]["D3"] = 500
            workbook["APOIO"]["D21"] = 100
            workbook["APOIO"]["E21"] = 100
            workbook["APOIO"]["G21"] = 20
            workbook["APOIO"]["H21"] = 10
            workbook.save(output)
            workbook.close()
            recalculate_workbook(output)
            cached = load_workbook(output, data_only=True)
            self.assertEqual(cached["RESUMO GASTOS"]["G8"].value, 344)
            self.assertEqual(cached["APOIO"]["F3"].value, 1100)
            for coordinate, expected in (
                ("C28", 1066.6666666666667),
                ("E28", 433.3333333333333),
                ("I28", 633.3333333333333),
                ("J28", 206.6666666666667),
                ("M28", 808),
            ):
                self.assertAlmostEqual(
                    cached["RESUMO GASTOS"][coordinate].value, expected, places=8
                )
            cached.close()

    def test_cancellations_new_categories_and_manual_groups_recalculate(self):
        from scdp_automation.xlsx_code_summary import code_group_formulas
        from scdp_automation.xlsx_output import build_candidate
        from tests.test_xlsx_output import make_trip

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            reference, output = root / "reference.xlsx", root / "output.xlsx"
            reference_fixture(reference)
            import_reference_workbook(reference, output)
            workbook = load_workbook(output)
            base, support = workbook["BASE VIAGENS"], workbook["APOIO"]
            base["C2"] = "Cancelada"
            workbook.save(output)
            workbook.close()
            recalculate_workbook(output)
            cached = load_workbook(output, data_only=True)
            self.assertEqual(cached["RESUMO GASTOS"]["F8"].value, "Pendente")
            self.assertEqual(cached["RESUMO GASTOS"]["G8"].value, "Pendente")
            cached.close()
            workbook = load_workbook(output)
            base, support = workbook["BASE VIAGENS"], workbook["APOIO"]
            base["Q2"] = "Não"
            row = support.max_row + 4
            support.cell(row, 1, "NOVO")
            support.cell(row, 2, "Nova categoria")
            support.cell(row, 3, "SEG 1 GRADUAÇÃO")
            support.cell(row, 4, 10)
            support.cell(row, 5, 20)
            summary = workbook["RESUMO GASTOS"]
            summary["B69"] = "NOVO"
            support.cell(row, 7, 0)
            for coordinate, formula in code_group_formulas(69).items():
                summary[coordinate] = formula
            row = base.max_row + 4
            base.cell(row, 1, "111111/26")
            base.cell(row, 3, "Concluída")
            base.cell(row, 11, 12)
            base.cell(row, 16, "NOVO")
            workbook.save(output)
            workbook.close()
            recalculate_workbook(output)
            cached = load_workbook(output, data_only=True)
            self.assertEqual(cached["RESUMO GASTOS"]["F8"].value, 0)
            self.assertEqual(cached["RESUMO GASTOS"]["C8"].value, 1000)
            cached.close()
            workbook = load_workbook(output)
            workbook["BASE VIAGENS"]["Q2"] = "Sim"
            group_row = workbook["APOIO"].max_row + 3
            support = workbook["APOIO"]
            for col, value in (
                (1, "GRUPO NOVO"),
                (2, "Grupo novo"),
                (3, "SEG 1 GRADUAÇÃO"),
                (4, 100),
                (5, 200),
                (7, 0),
            ):
                support.cell(group_row, col, value)
            summary = workbook["RESUMO GASTOS"]
            summary["B70"] = "Grupo novo"
            summary["O70"] = "Grupo adicional"
            for coordinate, formula in code_group_formulas(70).items():
                summary[coordinate] = formula
            before = [
                (c.coordinate, c.value, c.style_id) for row in summary for c in row
            ]
            workbook.save(output)
            workbook.close()
            recalculate_workbook(output)
            cached = load_workbook(output, data_only=True)
            self.assertEqual(cached["RESUMO GASTOS"]["F8"].value, 156)
            self.assertEqual(cached["RESUMO GASTOS"]["C70"].value, 300)
            self.assertEqual(cached["RESUMO GASTOS"]["C64"].value, 0)
            cached.close()
            candidate = root / "candidate.xlsx"
            build_candidate(
                [
                    make_trip("999001/26"),
                    make_trip("999002/26-1C"),
                    make_trip("111111/26"),
                ],
                output,
                candidate,
            )
            workbook = load_workbook(candidate)
            after = [
                (c.coordinate, c.value, c.style_id)
                for row in workbook["RESUMO GASTOS"]
                for c in row
            ]
            self.assertEqual(after, before)
            workbook.close()
