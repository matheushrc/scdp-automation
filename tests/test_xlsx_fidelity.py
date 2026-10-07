import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path
from unittest.mock import patch
from zoneinfo import ZoneInfo

from openpyxl import load_workbook

from scdp_automation import xlsx_output
from scdp_automation.xlsx_validation import WorkbookValidationError
from tests.support.fidelity import assert_manual_sheets_preserved
from tests.support.workbooks import final_template_fixture, make_trip


class FidelityTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)
        self.template = self.root / "template.xlsx"
        self.current = self.root / "current.xlsx"
        self.candidate = self.root / "candidate.xlsx"
        final_template_fixture(self.template)
        final_template_fixture(self.current)
        default = patch.object(xlsx_output, "DEFAULT_TEMPLATE", self.template)
        default.start()
        self.addCleanup(default.stop)

    def test_missing_template_preserves_published_files(self):
        original = self.current.read_bytes()
        checkpoint = self.root / "checkpoint.json"
        checkpoint.write_text('{"private": "synthetic"}')
        checkpoint_bytes = checkpoint.read_bytes()
        with (
            patch.object(xlsx_output, "DEFAULT_TEMPLATE", self.root / "missing.xlsx"),
            self.assertRaises(WorkbookValidationError),
        ):
            xlsx_output.build_candidate([], self.current, self.candidate)
        self.assertEqual(self.current.read_bytes(), original)
        self.assertEqual(checkpoint.read_bytes(), checkpoint_bytes)
        self.assertFalse(self.candidate.exists())

    def test_old_template_is_rejected(self):
        book = load_workbook(self.template)
        book.defined_names["SCDPLayoutVersion"].attr_text = '"8"'
        book.save(self.template)
        book.close()
        original = self.current.read_bytes()
        with self.assertRaises(WorkbookValidationError):
            xlsx_output.build_candidate([], self.current, self.candidate)
        self.assertEqual(self.current.read_bytes(), original)
        self.assertFalse(self.candidate.exists())

    def test_legacy_marker_with_current_manual_sheets_is_recoverable(self):
        book = load_workbook(self.current)
        book.defined_names["SCDPLayoutVersion"].attr_text = '"8"'
        book.save(self.current)
        book.close()
        original = self.current.read_bytes()
        xlsx_output.build_candidate([], self.current, self.candidate)
        self.assertEqual(self.current.read_bytes(), original)
        candidate = load_workbook(self.candidate)
        self.addCleanup(candidate.close)
        self.assertEqual(candidate.defined_names["SCDPLayoutVersion"].attr_text, '"10"')

    def test_fresh_base_does_not_import_template_trips(self):
        book = load_workbook(self.template)
        base = book["BASE VIAGENS"]
        for coordinate, value in (
            ("A2", "999001/26"),
            ("Q2", "AGRONOMIA"),
            ("R2", "Sim"),
            ("O2", "Antiga"),
            ("N2", date(2026, 1, 1)),
        ):
            base[coordinate] = value
        book.save(self.template)
        book.close()
        self.current.unlink()
        original = self.template.read_bytes()
        xlsx_output.build_candidate(
            [make_trip("999001/26")], self.current, self.candidate
        )
        actual = load_workbook(self.candidate)
        source = load_workbook(self.template)
        self.addCleanup(actual.close)
        self.addCleanup(source.close)
        self.assertEqual(self.template.read_bytes(), original)
        self.assertEqual(actual["BASE VIAGENS"]["A2"].value, "999001/26")
        for column in ("N", "O", "Q", "R"):
            self.assertIsNone(actual["BASE VIAGENS"][f"{column}2"].value)
        assert_manual_sheets_preserved(self, actual, source)

    def test_refresh_preserves_manual_sheets_and_full_pcdp_choices(self):
        book = load_workbook(self.current)
        base = book["BASE VIAGENS"]
        for row, (pcdp, code, decision) in enumerate(
            (("999001/26", "AGRONOMIA", "Não"), ("999001/26-1C", "PPGE", "Sim")), 2
        ):
            base.cell(row, 1, pcdp)
            base.cell(row, 17, code)
            base.cell(row, 18, decision)
        base.tables["tblBaseViagens"].ref = "A1:R3"
        book["APOIO"]["D2"] = 12345
        book["RESUMO GASTOS"]["F7"] = "=SUM(1,2)"
        book.save(self.current)
        book.close()
        original = self.current.read_bytes()
        trips = [
            make_trip(pcdp) for pcdp in ("999001/26-1C", "999001/26", "999001/26-2C")
        ]
        xlsx_output.build_candidate(trips, self.current, self.candidate)
        source, actual = load_workbook(self.current), load_workbook(self.candidate)
        self.addCleanup(source.close)
        self.addCleanup(actual.close)
        self.assertEqual(self.current.read_bytes(), original)
        assert_manual_sheets_preserved(self, actual, source)
        self.assertEqual(
            [actual["BASE VIAGENS"].cell(r, 17).value for r in (2, 3, 4)],
            ["PPGE", "AGRONOMIA", None],
        )
        self.assertEqual(
            [actual["BASE VIAGENS"].cell(r, 18).value for r in (2, 3, 4)],
            ["Sim", "Não", None],
        )

    def test_empty_description_does_not_restore_old_value(self):
        book = load_workbook(self.current)
        book["BASE VIAGENS"]["A2"] = "999001/26"
        book["BASE VIAGENS"]["O2"] = "Descrição antiga"
        book.save(self.current)
        book.close()
        trip = make_trip("999001/26")
        trip.descricao_do_motivo_da_viagem = ""
        xlsx_output.build_candidate([trip], self.current, self.candidate)
        updated = load_workbook(self.candidate)
        self.addCleanup(updated.close)
        self.assertIn(updated["BASE VIAGENS"]["O2"].value, (None, ""))
        self.assertEqual(
            trip.model_dump(mode="json")["descricao_do_motivo_da_viagem"], ""
        )

    def test_publication_records_the_value_update_date_in_m2(self):
        updated_at = datetime(2026, 10, 7, 1, 30, tzinfo=ZoneInfo("America/Sao_Paulo"))
        with patch("scdp_automation.xlsx_output.datetime") as clock:
            clock.now.return_value = updated_at
            xlsx_output.build_candidate([], self.current, self.candidate)
        workbook = load_workbook(self.candidate)
        self.addCleanup(workbook.close)
        cell = workbook["RESUMO GASTOS"]["M2"]
        self.assertEqual(cell.value.date(), date(2026, 10, 7))
        self.assertEqual(cell.number_format, "dd/mm/yyyy")

    def test_new_expense_requires_rateio(self):
        book = load_workbook(self.current)
        base, support = book["BASE VIAGENS"], book["APOIO"]
        base["A2"], base["Q2"], base["K2"] = "999001/26", "PPGEL +", 0
        row = next(
            r
            for r in range(2, support.max_row + 1)
            if support.cell(r, 1).value == "PPGEL +"
        )
        for column in (4, 5, 7, 8, 9, 10):
            support.cell(row, column).value = None
        book.save(self.current)
        book.close()
        original = self.current.read_bytes()
        trip = make_trip("999001/26")
        trip.total_da_viagem_r = 100
        with self.assertRaisesRegex(WorkbookValidationError, "100%"):
            xlsx_output.build_candidate([trip], self.current, self.candidate)
        self.assertEqual(self.current.read_bytes(), original)
        self.assertFalse(self.candidate.exists())

    def test_unverified_trip_does_not_restore_workbook_verification_date(self):
        book = load_workbook(self.current)
        book["BASE VIAGENS"]["A2"] = "999001/26"
        for column in ("L", "M", "N"):
            book["BASE VIAGENS"][f"{column}2"] = date(2020, 1, 1)
        book.save(self.current)
        book.close()
        trip = make_trip("999001/26")
        xlsx_output.build_candidate([trip], self.current, self.candidate)
        updated = load_workbook(self.candidate)
        self.addCleanup(updated.close)
        self.assertIsNone(updated["BASE VIAGENS"]["N2"].value)

    def test_stale_base_height_cannot_exclude_second_classified_trip(self):
        from tests.support.workbooks import write_existing_workbook

        trips = [make_trip("111111/26"), make_trip("222222/26")]
        trips[0].total_da_viagem_r = 100
        trips[1].total_da_viagem_r = 200
        write_existing_workbook(
            self.current, trips, {"111111/26": "AGRONOMIA", "222222/26": "PPGE"}
        )
        book = load_workbook(self.current)
        book["RESUMO GASTOS"]["R2"] = "=1"
        book.save(self.current)
        book.close()
        original = self.current.read_bytes()
        template_bytes = self.template.read_bytes()
        with self.assertRaises(WorkbookValidationError):
            xlsx_output.publish_workbook(
                trips, self.current, template_path=self.template
            )
        self.assertEqual(self.current.read_bytes(), original)
        self.assertEqual(self.template.read_bytes(), template_bytes)
        self.assertEqual(list(self.root.glob("*.candidate.xlsx")), [])
        self.assertEqual(list(self.root.glob("*.backup.*.xlsx")), [])

    def test_invalid_range_heights_reject_template_and_output_without_changes(self):
        for path in (self.template, self.current):
            for coordinate in ("R1", "R2", "R3"):
                for value in (
                    None,
                    "=1",
                    "=MAX(1,IFERROR(LOOKUP(2,1/('ERRADA'!$A:$A<>\"\"),ROW('ERRADA'!$A:$A))-1,1))",
                ):
                    with self.subTest(
                        path=path.name, coordinate=coordinate, value=value
                    ):
                        self.candidate.unlink(missing_ok=True)
                        final_template_fixture(self.template)
                        final_template_fixture(self.current)
                        book = load_workbook(path)
                        book["RESUMO GASTOS"][coordinate] = value
                        book.save(path)
                        book.close()
                        template_bytes = self.template.read_bytes()
                        current_bytes = self.current.read_bytes()
                        with self.assertRaises(WorkbookValidationError):
                            xlsx_output.build_candidate(
                                [],
                                self.current,
                                self.candidate,
                                template_path=self.template,
                            )
                        self.assertEqual(self.template.read_bytes(), template_bytes)
                        self.assertEqual(self.current.read_bytes(), current_bytes)
                        self.assertFalse(self.candidate.exists())
