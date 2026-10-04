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
            self.assertEqual(cached["APOIO"]["F3"].value, 600)
            cached.close()
            workbook = load_workbook(output)
            _check_workbook_structure(workbook)
            self.assertEqual(workbook["RESUMO GASTOS"]["F8"].value, formula)
            self.assertEqual(copy(workbook["RESUMO GASTOS"]["B2"].font), font)
            workbook["APOIO"]["D3"] = 500
            workbook["APOIO"]["D21"] = 100
            workbook["APOIO"]["E21"] = 200
            workbook["APOIO"]["G21"] = 20
            workbook["APOIO"]["H21"] = 10
            workbook.save(output)
            workbook.close()
            recalculate_workbook(output)
            cached = load_workbook(output, data_only=True)
            self.assertEqual(cached["RESUMO GASTOS"]["G8"].value, 344)
            self.assertEqual(cached["APOIO"]["F3"].value, 500)
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
