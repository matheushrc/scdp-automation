import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path
from unittest.mock import AsyncMock, patch
from zoneinfo import ZoneInfo

from openpyxl import load_workbook

from scdp_automation import extrator, xlsx_output
from scdp_automation.xlsx_models import summarize_trips
from tests.test_xlsx_output import make_trip
from tests.xlsx_fixtures import reference_fixture


class DateTests(unittest.IsolatedAsyncioTestCase):
    def test_dates_and_future_selection(self):
        trip = make_trip()
        trip.trechos = [trip.trechos[0]]
        trip.trechos[0].inicio = "06/10/2026"
        trip.trechos[0].termino = "08/10/2026"
        trip.data_da_ultima_verificacao = date(2026, 10, 4)
        trip.descricao_do_motivo_da_viagem = "Anterior"
        summary = summarize_trips([trip])[0]
        self.assertEqual(summary.start_date, date(2026, 10, 6))
        self.assertEqual(summary.end_date, date(2026, 10, 8))
        self.assertEqual(summary.verified_date, date(2026, 10, 4))
        self.assertTrue(extrator.needs_verification(trip, date(2026, 10, 5)))
        trip.trechos[0].inicio = "05/10/2026"
        self.assertFalse(extrator.needs_verification(trip, date(2026, 10, 5)))

    async def test_success_records_date_failure_keeps_old_date(self):
        trip = make_trip()
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "checkpoint.json"
            with patch.object(
                extrator, "_consult_with_safe_failure", AsyncMock(return_value="Atual")
            ):
                await extrator.collect_pending_descriptions(
                    AsyncMock(), [trip], [trip], output
                )
            self.assertEqual(
                trip.data_da_ultima_verificacao,
                datetime.now(ZoneInfo("America/Sao_Paulo")).date(),
            )
            old_date = date(2026, 1, 1)
            trip.data_da_ultima_verificacao = old_date
            with (
                patch.object(
                    extrator,
                    "_consult_with_safe_failure",
                    AsyncMock(side_effect=RuntimeError("falha")),
                ),
                self.assertRaises(RuntimeError),
            ):
                await extrator.collect_pending_descriptions(
                    AsyncMock(), [trip], [trip], output
                )
            self.assertEqual(trip.data_da_ultima_verificacao, old_date)


class WorkbookAdjustmentsTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.reference = self.root / "reference.xlsx"
        reference_fixture(self.reference)

    def test_decisions_and_dates_follow_complete_pcdp(self):
        current, candidate = self.root / "current.xlsx", self.root / "candidate.xlsx"
        xlsx_output.import_reference_workbook(self.reference, current)
        workbook = load_workbook(current)
        base = workbook["BASE VIAGENS"]
        self.assertEqual(base["Q1"].value, "Descontar do curso?")
        base["Q2"] = "Não"
        base["Q3"] = "Sim"
        workbook.save(current)
        workbook.close()
        trips = [make_trip("999002/26-1C"), make_trip("999001/26")]
        xlsx_output.build_candidate(trips, current, candidate)
        workbook = load_workbook(candidate)
        self.addCleanup(workbook.close)
        base = workbook["BASE VIAGENS"]
        self.assertEqual([base.cell(r, 17).value for r in (2, 3)], ["Sim", "Não"])
        self.assertEqual(base["L2"].value.date(), date(2026, 3, 1))
        self.assertEqual(base["L2"].number_format, "dd/mm/yyyy")
        self.assertTrue(
            any(
                v.formula1 == '"Sim,Não"' and "Q1048576" in str(v.sqref)
                for v in base.data_validations.dataValidation
            )
        )
        self.assertTrue(
            any(str(c.sqref) == "A2:Q1048576" for c in base.conditional_formatting)
        )

    def test_support_order_history_group_and_manual_summary_survive(self):
        current, candidate = self.root / "current.xlsx", self.root / "candidate.xlsx"
        xlsx_output.import_reference_workbook(self.reference, current)
        workbook = load_workbook(current)
        support = workbook["APOIO"]
        self.assertEqual(
            [support.cell(1, c).value for c in (4, 5, 6)],
            [
                "Diárias e passagens distribuído (R$)",
                "Transportes distribuído (R$)",
                "Recurso total (R$)",
            ],
        )
        self.assertEqual(support["E3"].value, 600)
        groups = {
            support.cell(r, 1).value: support.cell(r, 9).value
            for r in range(2, support.max_row + 1)
        }
        self.assertIn("PPGH/PPGDH", groups)
        self.assertNotIn("PPGH", groups)
        self.assertNotIn("PPGDH", groups)
        self.assertNotEqual(groups["HISTÓRIA"], groups["PPGH/PPGDH"])
        support.append(
            [
                "NOVO",
                "Categoria nova",
                "SEG 1 GRADUAÇÃO",
                0,
                0,
                "=D40+E40",
                0,
                0,
                groups["AGRONOMIA"],
            ]
        )
        workbook["RESUMO GASTOS"]["F8"] = "=SUM(1,2)"
        workbook.save(current)
        workbook.close()
        xlsx_output.build_candidate(
            [make_trip("999001/26"), make_trip("999002/26-1C")], current, candidate
        )
        workbook = load_workbook(candidate)
        self.addCleanup(workbook.close)
        self.assertEqual(workbook["RESUMO GASTOS"]["F8"].value, "=SUM(1,2)")
        self.assertEqual(workbook["APOIO"]["E3"].value, 600)
        self.assertEqual(
            workbook["APOIO"].cell(workbook["APOIO"].max_row, 1).value, "NOVO"
        )


class HistoricalDecisionsTests(unittest.TestCase):
    def test_decisions_require_financial_reconciliation_and_zero_is_ambiguous(self):
        from dataclasses import replace

        from openpyxl import Workbook

        from scdp_automation.xlsx_reference import infer_decisions

        workbook = Workbook()
        support = workbook.active
        support.title = "APOIO"
        support["B41"] = "AGRONOMIA"
        support["C41"] = 30
        first = summarize_trips([make_trip("111111/26")])[0]
        trips = (
            replace(first, status="Cancelada", trip_total=10),
            replace(first, pcdp="222222/26", status="Cancelada", trip_total=0),
            replace(first, pcdp="333333/26", status="Concluída", trip_total=20),
        )
        codes = {t.pcdp: "AGRONOMIA" for t in trips}
        self.assertEqual(
            infer_decisions(trips, codes, workbook),
            {"111111/26": "Sim", "333333/26": "Sim"},
        )
        support["C41"] = 20
        self.assertEqual(infer_decisions(trips, codes, workbook)["111111/26"], "Não")
        support["C41"] = 25
        self.assertEqual(infer_decisions(trips, codes, workbook), {})
        workbook.close()


class MigrationTests(unittest.TestCase):
    reference: Path
    root: Path
    setUp = WorkbookAdjustmentsTests.setUp

    def test_migration_recovers_historical_decision_and_preserves_manual_budget(self):
        from openpyxl.workbook.defined_name import DefinedName

        workbook = load_workbook(self.reference)
        workbook["BD D&P"]["I5"] = "Cancelada"
        workbook["BD D&P"]["I11"] = "Cancelada"
        workbook["APOIO"]["B41"] = "AGRONOMIA"
        workbook["APOIO"]["C41"] = 156
        workbook.save(self.reference)
        workbook.close()
        current, candidate = self.root / "current.xlsx", self.root / "candidate.xlsx"
        xlsx_output.import_reference_workbook(self.reference, current)
        workbook = load_workbook(current)
        support = workbook["APOIO"]
        for r in range(2, support.max_row + 1):
            daily, transport = support.cell(r, 4).value, support.cell(r, 5).value
            support.cell(r, 5).value = (
                daily + transport
                if isinstance(daily, (int, float))
                and isinstance(transport, (int, float))
                else None
            )
            support.cell(
                r, 6
            ).value = f'=IF(AND(ISNUMBER(D{r}),ISNUMBER(E{r})),E{r}-D{r},"")'
        support["E1"] = "Recurso total (R$)"
        support["F1"] = "Transportes distribuído (R$)"
        support["D3"] = 555
        support["E3"] = 999
        base = workbook["BASE VIAGENS"]
        for row in range(1, base.max_row + 1):
            saved = {col: base.cell(row, col).value for col in range(12, 18)}
            for old, new in (
                (15, 12),
                (16, 13),
                (17, 14),
                (12, 15),
                (13, 16),
                (14, 17),
            ):
                base.cell(row, new).value = saved[old]
        base.delete_cols(14, 4)
        workbook.defined_names.add(DefinedName("SCDPLayoutVersion", attr_text='"5"'))
        workbook.save(current)
        workbook.close()
        xlsx_output.build_candidate(
            [make_trip("999001/26"), make_trip("999002/26-1C")],
            current,
            candidate,
            reference_path=self.reference,
        )
        workbook = load_workbook(candidate)
        self.addCleanup(workbook.close)
        self.assertEqual(workbook["BASE VIAGENS"]["Q2"].value, "Sim")
        self.assertEqual(workbook["APOIO"]["D3"].value, 555)
        self.assertEqual(workbook["APOIO"]["E3"].value, 444)


class ColumnOrderTests(unittest.TestCase):
    def test_classification_columns_are_last_in_generated_file(self):
        self.assertEqual(
            xlsx_output.BASE_HEADERS[-3:],
            ("Segmento", "Código de débito", "Descontar do curso?"),
        )
        self.assertEqual(
            xlsx_output.BASE_HEADERS[11:14],
            (
                "Data de início da viagem",
                "Data de término da viagem",
                "Data da última verificação",
            ),
        )

    def test_existing_v6_reorders_dates_and_unifies_history_without_losing_values(self):
        from copy import copy

        from openpyxl.workbook.defined_name import DefinedName

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            reference, current, candidate = (
                root / "reference.xlsx",
                root / "current.xlsx",
                root / "candidate.xlsx",
            )
            reference_fixture(reference)
            xlsx_output.import_reference_workbook(reference, current)
            workbook = load_workbook(current)
            support = workbook["APOIO"]
            history_row = next(
                r
                for r in range(2, support.max_row + 1)
                if support.cell(r, 1).value == "PPGH/PPGDH"
            )
            support.insert_rows(history_row + 1)
            for col in range(1, 14):
                support.cell(history_row + 1, col).value = support.cell(
                    history_row, col
                ).value
                support.cell(history_row + 1, col)._style = copy(
                    support.cell(history_row, col)._style
                )
            support.cell(history_row, 1).value = "PPGH"
            support.cell(history_row + 1, 1).value = "PPGDH"
            support.cell(history_row + 1, 4).value = 100
            support.cell(history_row + 1, 5).value = 200
            manual_row = history_row + 2
            support.cell(manual_row, 6).value = "=1+2"
            support.cell(manual_row, 13).value = "=3+4"
            workbook["RESUMO GASTOS"]["F8"] = f"=APOIO!D{manual_row}"
            workbook["RESUMO GASTOS"]["M8"] = "=COUNT('BASE VIAGENS'!O2:O3)"
            base = workbook["BASE VIAGENS"]
            base["P2"], base["P3"] = "PPGH", "PPGDH"
            base["Q2"], base["Q3"] = "Não", "Sim"
            base["L2"] = date(2026, 1, 10)
            for row in range(1, base.max_row + 1):
                saved = {col: base.cell(row, col).value for col in range(12, 18)}
                for old, new in (
                    (15, 12),
                    (16, 13),
                    (17, 14),
                    (12, 15),
                    (13, 16),
                    (14, 17),
                ):
                    base.cell(row, new).value = saved[old]
            workbook.defined_names.add(
                DefinedName("SCDPLayoutVersion", attr_text='"6"')
            )
            workbook.save(current)
            workbook.close()
            before = current.read_bytes()
            from scdp_automation.xlsx_reference import read_reference

            xlsx_output.build_candidate(
                read_reference(reference).trips, current, candidate
            )
            self.assertEqual(current.read_bytes(), before)
            workbook = load_workbook(candidate)
            self.addCleanup(workbook.close)
            base = workbook["BASE VIAGENS"]
            self.assertEqual(
                [base.cell(1, c).value for c in (15, 16, 17)],
                ["Segmento", "Código de débito", "Descontar do curso?"],
            )
            self.assertEqual(
                [base.cell(r, 16).value for r in (2, 3)], ["PPGH/PPGDH", "PPGH/PPGDH"]
            )
            self.assertEqual([base.cell(r, 17).value for r in (2, 3)], ["Não", "Sim"])
            self.assertEqual(base["L2"].value.date(), date(2026, 1, 10))
            support = workbook["APOIO"]
            history = [
                r
                for r in range(2, support.max_row + 1)
                if support.cell(r, 1).value in ("PPGH", "PPGDH", "PPGH/PPGDH")
            ]
            self.assertEqual(support.cell(manual_row, 6).value, "=1+2")
            self.assertEqual(support.cell(manual_row, 13).value, "=3+4")
            self.assertEqual(
                workbook["RESUMO GASTOS"]["F8"].value, f"=APOIO!D{manual_row}"
            )
            self.assertEqual(
                workbook["RESUMO GASTOS"]["M8"].value, "=COUNT('BASE VIAGENS'!L2:L3)"
            )
            self.assertEqual(len(history), 1)
            self.assertEqual(support.cell(history[0], 1).value, "PPGH/PPGDH")
            self.assertEqual(
                [support.cell(history[0], c).value for c in (4, 5)], [500, 800]
            )
