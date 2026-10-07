import shutil
import tempfile
import unittest
from copy import copy
from pathlib import Path

from openpyxl import load_workbook

from scdp_automation.xlsx_output import build_candidate
from scdp_automation.xlsx_validation import validate_workbook
from tests.support.recalculation import recalculate_workbook
from tests.support.workbooks import (
    final_template_fixture,
    make_trip,
    write_existing_workbook,
)


@unittest.skipUnless(shutil.which("libreoffice"), "LibreOffice unavailable")
class RecalculationTests(unittest.TestCase):
    def test_visible_results_recalculate_while_original_formulas_and_styles_survive(
        self,
    ):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            template, current, candidate = (
                root / name
                for name in ("template.xlsx", "current.xlsx", "candidate.xlsx")
            )
            final_template_fixture(template)
            trip = make_trip()
            trip.total_da_viagem_r = 156
            write_existing_workbook(
                current, [trip], {trip.numero_da_solicitacao: "AGRONOMIA"}
            )
            build_candidate([trip], current, candidate, template_path=template)
            workbook = load_workbook(candidate)
            font = copy(workbook["RESUMO GASTOS"]["B2"].font)
            formula = workbook["RESUMO GASTOS"]["F7"].value
            workbook.close()
            recalculate_workbook(candidate)
            cached = load_workbook(candidate, data_only=True)
            self.assertEqual(
                [
                    cached["RESUMO GASTOS"][c].value
                    for c in ("F7", "G7", "J7", "K7", "M7", "F20")
                ],
                [156, 244, 200, 400, 644, 156],
            )
            self.assertEqual(cached["APOIO"]["F2"].value, 1000)
            cached.close()
            workbook = load_workbook(candidate)
            validate_workbook(workbook)
            self.assertEqual(workbook["RESUMO GASTOS"]["F7"].value, formula)
            self.assertEqual(copy(workbook["RESUMO GASTOS"]["B2"].font), font)
            workbook["APOIO"]["D2"] = 500
            workbook.save(candidate)
            workbook.close()
            recalculate_workbook(candidate)
            cached = load_workbook(candidate, data_only=True)
            self.assertEqual(cached["RESUMO GASTOS"]["G7"].value, 344)
            self.assertEqual(cached["APOIO"]["F2"].value, 1100)
            cached.close()
