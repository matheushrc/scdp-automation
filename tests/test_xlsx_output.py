import unittest
from concurrent.futures import ThreadPoolExecutor
from dataclasses import FrozenInstanceError
from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from scdp_automation import xlsx_output
from scdp_automation.config import current_year
from scdp_automation.relatorio import Viagem
from scdp_automation.xlsx_output import (
    DEBIT_CATEGORIES,
    DebitCategory,
    TripSummary,
    WorkbookValidationError,
    build_candidate,
    create_workbook_template,
    publish_workbook,
    summarize_trips,
)
from tests.support.workbooks import make_trip, read_base_rows, write_existing_workbook

_reference_patch = patch.object(
    xlsx_output,
    "DEFAULT_REFERENCE",
    Path(__file__).parent / "unavailable-reference.xlsx",
)


def setUpModule() -> None:
    _reference_patch.start()


def tearDownModule() -> None:
    _reference_patch.stop()


class TripSummaryTests(unittest.TestCase):
    def test_summarize_trip_uses_trip_subtotal_once(self) -> None:
        trip = make_trip()

        summary = summarize_trips([trip])[0]

        self.assertIsInstance(summary, TripSummary)
        self.assertEqual(
            summary,
            TripSummary(
                pcdp="123456/26-2B",
                proposed="Pessoa Exemplo",
                status="Autorizada",
                daily_count=8.5,
                daily_amount=1234.56,
                ticket_amount=789.01,
                additional_amount=33.05,
                discount_amount=4.01,
                restitution_amount=5.02,
                reimbursement_amount=6.03,
                trip_total=2063.66,
                start_date=date(2026, 3, 1),
                end_date=date(2026, 3, 5),
            ),
        )

    def test_summarize_trips_rejects_duplicate_full_pcdp(self) -> None:
        with self.assertRaisesRegex(ValueError, "PCDP duplicada na listagem") as error:
            summarize_trips([make_trip(), make_trip()])
        self.assertNotIn("123456/26-2B", str(error.exception))

        summaries = summarize_trips([make_trip("123456/26"), make_trip("123456/26-2B")])

        self.assertEqual(
            [summary.pcdp for summary in summaries], ["123456/26", "123456/26-2B"]
        )
        complementary_summaries = summarize_trips(
            [make_trip("123456/26-1C"), make_trip("123456/26-2C")]
        )
        self.assertEqual(
            [summary.pcdp for summary in complementary_summaries],
            ["123456/26-1C", "123456/26-2C"],
        )


class DebitCategoryTests(unittest.TestCase):
    def test_debit_catalog_covers_active_and_future_categories(self) -> None:
        self.assertTrue(DEBIT_CATEGORIES)
        self.assertTrue(
            all(isinstance(item, DebitCategory) for item in DEBIT_CATEGORIES)
        )
        self.assertTrue(
            all(item.__dataclass_params__.frozen for item in DEBIT_CATEGORIES)
        )
        self.assertEqual(
            {item.segment for item in DEBIT_CATEGORIES},
            {
                "SEG 1 GRADUAÇÃO",
                "SEG 2 MESTRADO",
                "SEG 3 OUTROS",
                "SEG 4 AUX EVENTOS",
                "SEG 5 AFAST PAÍS",
            },
        )

        codes = [item.code for item in DEBIT_CATEGORIES]
        self.assertEqual(len(codes), len(set(codes)))
        self.assertEqual(
            set(codes),
            {
                "ADMINISTRAÇÃO",
                "AGRONOMIA",
                "C COMPUTAÇÃO",
                "C ECONÔMICAS",
                "CIÊNCIAS SOCIAIS",
                "ENFERMAGEM",
                "ENG AMBIENTAL",
                "ENGENHARIA CIVIL",
                "FILOSOFIA",
                "GEOGRAFIA",
                "HISTÓRIA",
                "LETRAS",
                "MATEMÁTICA",
                "MEDICINA",
                "PEDAGOGIA",
                "Lato Oncologia",
                "PPGCB",
                "PPGE",
                "PPGEL",
                "PPGEL +",
                "PPGEnf",
                "PPGFil",
                "PPGGeo",
                "PPGH/PPGDH",
                "PROFIAP",
                "PROFMAT",
                "DIREÇÃO",
                "DIREÇÃO - AGAS",
                "DIREÇÃO - Banca Libras",
                "DIREÇÃO - CAAEX",
                "DIREÇÃO - Empr Junior",
                "DIREÇÃO - StartUp Summit",
                "DIREÇÃO - StartUp Weekend",
                "DIREÇÃO - Sunset",
                "CAPPG - Res 49",
                "AFAST PAÍS",
            },
        )
        with self.assertRaises(FrozenInstanceError):
            DEBIT_CATEGORIES[0].__setattr__("name", "alterado")

        category_by_code = {item.code: item for item in DEBIT_CATEGORIES}
        self.assertEqual(category_by_code["AGRONOMIA"].name, "Agronomia")
        self.assertEqual(
            category_by_code["Lato Oncologia"].name,
            "Especialização em Enfermagem em Oncologia",
        )
        self.assertEqual(
            category_by_code["PPGEL +"].name, "Estudos Linguísticos – Rateio"
        )
        self.assertEqual(
            category_by_code["DIREÇÃO"].name, "Geral (Direção/Coordenações)"
        )
        for label in (
            "DIREÇÃO - AGAS",
            "DIREÇÃO - Banca Libras",
            "DIREÇÃO - CAAEX",
            "DIREÇÃO - Empr Junior",
            "DIREÇÃO - StartUp Summit",
            "DIREÇÃO - StartUp Weekend",
            "DIREÇÃO - Sunset",
        ):
            self.assertEqual(category_by_code[label].name, label)

        for code in (
            "ADMINISTRAÇÃO",
            "C COMPUTAÇÃO",
            "C ECONÔMICAS",
            "LETRAS",
            "PPGFil",
            "PPGH/PPGDH",
            "PROFMAT",
        ):
            self.assertIn(code, category_by_code)

        ppghd = category_by_code["PPGH/PPGDH"]
        self.assertFalse(ppghd.review_required)
        self.assertEqual(ppghd.name, "Mestrado e Doutorado em História")


class WorkbookTemplateTests(unittest.TestCase):
    def create_template(self, directory: Path):
        path = directory / "gastos.xlsx"
        create_workbook_template(path)
        from openpyxl import load_workbook

        return load_workbook(path, data_only=False)

    def test_create_template_has_expected_sheets_and_columns(self) -> None:
        expected_base_headers = [
            "PCDP",
            "Proposto",
            "Situação",
            "Quantidade de diárias",
            "Diárias (R$)",
            "Passagens (R$)",
            "Adicional (R$)",
            "Descontos (R$)",
            "Restituição (R$)",
            "Reembolso (R$)",
            "Total da viagem (R$)",
            "Data de início da viagem",
            "Data de término da viagem",
            "Data da última verificação",
            "Descrição do pedido",
            "Segmento",
            "Código de débito",
            "Descontar do curso?",
        ]

        with TemporaryDirectory() as directory:
            workbook = self.create_template(Path(directory))

        self.assertEqual(
            workbook.sheetnames, ["BASE VIAGENS", "APOIO", "RESUMO GASTOS"]
        )
        self.assertEqual(
            [cell.value for cell in workbook["BASE VIAGENS"][1]], expected_base_headers
        )
        self.assertEqual(
            [cell.value for cell in workbook["APOIO"][1]],
            [
                "Código de débito",
                "Nome por extenso",
                "Segmento",
                "Alocação inicial (R$)",
            ],
        )
        self.assertEqual(workbook["BASE VIAGENS"].cell(1, 17).value, "Código de débito")
        self.assertTrue(workbook.calculation.fullCalcOnLoad)
        self.assertTrue(workbook.calculation.forceFullCalc)
        self.assertEqual(workbook.calculation.calcMode, "auto")

    def test_template_seeds_active_and_future_debit_categories(self) -> None:
        with TemporaryDirectory() as directory:
            workbook = self.create_template(Path(directory))

        support = workbook["APOIO"]
        records = [
            (
                support.cell(row, 1).value,
                support.cell(row, 2).value,
                support.cell(row, 3).value,
            )
            for row in range(2, support.max_row + 1)
        ]
        expected = [(item.code, item.name, item.segment) for item in DEBIT_CATEGORIES]
        self.assertEqual(records, expected)
        self.assertTrue(
            all(
                support.cell(row, 4).value is None
                for row in range(2, support.max_row + 1)
            )
        )

        ppghd_row = next(
            row
            for row in range(2, support.max_row + 1)
            if support.cell(row, 1).value == "PPGH/PPGDH"
        )
        self.assertEqual(
            support.cell(ppghd_row, 2).value, "Mestrado e Doutorado em História"
        )
        self.assertIsNotNone(support.cell(ppghd_row, 2).comment)
        self.assertIn("compartilhada", support.cell(ppghd_row, 2).comment.text.lower())

    def test_base_code_column_has_dropdown_validation(self) -> None:
        with TemporaryDirectory() as directory:
            workbook = self.create_template(Path(directory))

        base = workbook["BASE VIAGENS"]
        validations = [
            validation
            for validation in base.data_validations.dataValidation
            if validation.type == "list" and validation.formula1 == "=CodigosDebito"
        ]
        self.assertEqual(len(validations), 1)
        self.assertEqual(validations[0].formula1, "=CodigosDebito")
        self.assertEqual(str(validations[0].sqref), "Q2:Q1048576")

        named_range = workbook.defined_names["CodigosDebito"]
        self.assertEqual(
            named_range.attr_text,
            f"'APOIO'!$A$2:$A${len(DEBIT_CATEGORIES) + 1}",
        )

    def test_summary_contains_excel_formulas_for_both_grouping_keys(self) -> None:
        with TemporaryDirectory() as directory:
            workbook = self.create_template(Path(directory))

        summary = workbook["RESUMO GASTOS"]
        headers = [cell.value for cell in summary[1]]
        self.assertIn("Quantidade PCDPs", headers)
        self.assertIn("TOTAL DIÁRIAS (R$)", headers)
        self.assertIn("PASS AÉREA+ROD (R$)", headers)
        self.assertIn("CANCELADAS (R$)", headers)

        self.assertTrue(summary["O2"].value.startswith("=SUMIFS("))
        self.assertIn("'BASE VIAGENS'!$P:$P,$A2", summary["O2"].value)
        self.assertIn("'BASE VIAGENS'!$Q:$Q,$B2", summary["O2"].value)
        self.assertEqual(summary["J2"].value, "=G2+H2-I2")
        self.assertEqual(summary["M2"].value, "=K2+L2")
        self.assertEqual(summary["Q2"].value, "=O2")
        self.assertIn("'BASE VIAGENS'!$C:$C,\"*Cancel*\"", summary["P2"].value)
        self.assertEqual(
            summary["D2"].value, "=IF('APOIO'!$D$2=\"\",\"Pendente\",'APOIO'!$D$2)"
        )
        self.assertIn('IF(D2="Pendente","Pendente",', summary["R2"].value)

        pending_row = summary.max_row
        self.assertEqual(summary.cell(pending_row, 1).value, "SEM CLASSIFICAÇÃO")
        self.assertTrue(summary.cell(pending_row, 5).value.startswith("=COUNTIFS("))
        self.assertTrue(summary.cell(pending_row, 15).value.startswith("=SUMIFS("))
        self.assertEqual(summary.cell(pending_row, 4).value, "Pendente")

    def test_summary_metrics_reconcile_daily_ticket_refund_and_cancelled_totals(
        self,
    ) -> None:
        active_data = make_trip("111111/26").model_dump()
        canceled_data = make_trip("222222/26").model_dump()
        canceled_data["situacao_da_viagem"] = "Cancelada"
        zero_data = make_trip("333333/26").model_dump()
        zero_data["sub_total"] = {
            "quantidade_diarias": 0,
            "diarias_r": 0,
            "passagens_e_taxas_iniciais_r": 0,
            "total_r": 0,
        }
        zero_data.update(
            {
                "total_adicional_r": 0,
                "descontos_r": 0,
                "restituicao_r": 0,
                "reembolso_r": 0,
                "total_da_viagem_r": 0,
            }
        )
        trips = [
            Viagem.model_validate(active_data),
            Viagem.model_validate(canceled_data),
            Viagem.model_validate(zero_data),
        ]
        summaries = summarize_trips(trips)

        daily_total = sum(
            item.daily_amount + item.additional_amount - item.discount_amount
            for item in summaries
        )
        airfare_and_rail = sum(
            item.ticket_amount + item.restitution_amount for item in summaries
        )
        used_once = sum(item.trip_total for item in summaries)
        canceled_breakdown = sum(
            item.trip_total
            for trip, item in zip(trips, summaries, strict=True)
            if "cancel" in trip.situacao_da_viagem.casefold()
        )
        restitution_breakdown = sum(item.restitution_amount for item in summaries)

        self.assertAlmostEqual(daily_total, 2527.20)
        self.assertAlmostEqual(airfare_and_rail, 1588.06)
        self.assertAlmostEqual(used_once, 4127.32)
        self.assertAlmostEqual(used_once, daily_total + airfare_and_rail + 12.06)
        self.assertAlmostEqual(canceled_breakdown, 2063.66)
        self.assertAlmostEqual(restitution_breakdown, 10.04)

        with TemporaryDirectory() as directory:
            path = Path(directory) / "reconciliation.xlsx"
            write_existing_workbook(
                path,
                trips,
                {
                    trips[0].numero_da_solicitacao: "AGRONOMIA",
                    trips[1].numero_da_solicitacao: "AFAST PAÍS",
                    trips[2].numero_da_solicitacao: "AFAST PAÍS",
                },
            )
            from openpyxl import load_workbook

            workbook = load_workbook(path, data_only=False)

        summary = workbook["RESUMO GASTOS"]
        base = workbook["BASE VIAGENS"]
        self.assertEqual(
            [base.cell(row, 17).value for row in range(2, 5)],
            ["AGRONOMIA", "AFAST PAÍS", "AFAST PAÍS"],
        )
        rows_by_code = {
            item.code: index + 2 for index, item in enumerate(DEBIT_CATEGORIES)
        }
        active_row = rows_by_code["AGRONOMIA"]
        future_row = rows_by_code["AFAST PAÍS"]
        self.assertEqual(
            summary.cell(active_row, 10).value,
            f"=G{active_row}+H{active_row}-I{active_row}",
        )
        self.assertEqual(
            summary.cell(active_row, 13).value,
            f"=K{active_row}+L{active_row}",
        )
        self.assertIn(
            "'BASE VIAGENS'!$C:$C,\"*Cancel*\"",
            summary.cell(future_row, 16).value,
        )
        used_formula = summary.cell(active_row, 15).value
        self.assertTrue(used_formula.startswith("=SUMIFS("))
        self.assertIn("'BASE VIAGENS'!$K:$K", used_formula)
        self.assertIn(f"'BASE VIAGENS'!$P:$P,$A{active_row}", used_formula)
        self.assertIn(f"'BASE VIAGENS'!$Q:$Q,$B{active_row}", used_formula)
        self.assertEqual(summary.cell(active_row, 17).value, f"=O{active_row}")
        self.assertNotIn("+P", summary.cell(active_row, 17).value)
        self.assertNotIn("+L", summary.cell(active_row, 17).value)


class WorkbookRefreshTests(unittest.TestCase):
    def test_description_migration_preserves_codes_decisions_and_reference_formulas(
        self,
    ) -> None:
        from openpyxl import load_workbook
        from openpyxl.workbook.defined_name import DefinedName

        from tests.support.workbooks import restore_base_without_description

        trip = make_trip()
        trip.descricao_do_motivo_da_viagem = "Descrição já extraída."
        with TemporaryDirectory() as directory:
            current = Path(directory) / "current.xlsx"
            candidate = Path(directory) / "candidate.xlsx"
            write_existing_workbook(
                current, [trip], {trip.numero_da_solicitacao: "AGRONOMIA"}
            )
            workbook = load_workbook(current)
            workbook["BASE VIAGENS"].cell(2, 18, "Não")
            restore_base_without_description(workbook)
            workbook.defined_names.add(
                DefinedName("ManualCode", attr_text="'BASE VIAGENS'!$P$2")
            )
            workbook.save(current)
            workbook.close()
            before = current.read_bytes()

            build_candidate([trip], current, candidate)

            self.assertEqual(current.read_bytes(), before)
            workbook = load_workbook(candidate)
            base = workbook["BASE VIAGENS"]
            self.assertEqual(base["O2"].value, trip.descricao_do_motivo_da_viagem)
            self.assertEqual(base["Q2"].value, "AGRONOMIA")
            self.assertEqual(base["R2"].value, "Não")
            self.assertEqual(
                workbook.defined_names["ManualCode"].attr_text,
                "'BASE VIAGENS'!$Q$2",
            )
            self.assertEqual(base.tables["tblBaseViagens"].ref, "A1:R2")
            workbook.close()

    def test_refresh_exports_request_description_before_classification(self) -> None:
        from openpyxl import load_workbook

        trip = make_trip()
        trip.descricao_do_motivo_da_viagem = "Participação em evento acadêmico."
        with TemporaryDirectory() as directory:
            current = Path(directory) / "missing.xlsx"
            candidate = Path(directory) / "candidate.xlsx"
            build_candidate([trip], current, candidate)
            workbook = load_workbook(candidate)
            headers = [cell.value for cell in workbook["BASE VIAGENS"][1]]
            self.assertIn("Descrição do pedido", headers)
            self.assertEqual(
                headers[-3:], ["Segmento", "Código de débito", "Descontar do curso?"]
            )
            column = headers.index("Descrição do pedido") + 1
            self.assertEqual(
                workbook["BASE VIAGENS"].cell(2, column).value,
                trip.descricao_do_motivo_da_viagem,
            )
            workbook.close()

    def test_first_refresh_uses_current_reference_without_changing_it(self) -> None:
        from openpyxl import load_workbook

        trip = make_trip()
        with TemporaryDirectory() as directory:
            reference = Path(directory) / "reference.xlsx"
            current = Path(directory) / "missing.xlsx"
            candidate = Path(directory) / "candidate.xlsx"
            template = Path(directory) / "template.xlsx"
            write_existing_workbook(
                reference, [trip], {trip.numero_da_solicitacao: "AGRONOMIA"}
            )
            workbook = load_workbook(reference)
            workbook["BASE VIAGENS"].cell(2, 18, "Sim")
            workbook["APOIO"].cell(2, 4, 4321)
            workbook.save(reference)
            workbook.close()
            original_bytes = reference.read_bytes()

            create_workbook_template(template, reference)
            build_candidate([trip], current, candidate, reference_path=reference)

            self.assertEqual(reference.read_bytes(), original_bytes)
            self.assertFalse(current.exists())
            self.assertEqual(read_base_rows(template), {})
            rows = read_base_rows(candidate)
            self.assertIsNone(rows[trip.numero_da_solicitacao][16])
            for path in (template, candidate):
                workbook = load_workbook(path)
                self.assertIsNone(workbook["BASE VIAGENS"].cell(2, 18).value)
                self.assertEqual(workbook["APOIO"].cell(2, 4).value, 4321)
                workbook.close()

    def test_refresh_preserves_manual_codes_by_full_pcdp_after_reordering(self) -> None:
        first = make_trip("123456/26-1C")
        second = make_trip("123456/26-2C")
        new_trip = make_trip("234567/26-1C")

        with TemporaryDirectory() as directory:
            current = Path(directory) / "current.xlsx"
            candidate = Path(directory) / "candidate.xlsx"
            write_existing_workbook(
                current,
                [first, second],
                {
                    first.numero_da_solicitacao: "AGRONOMIA",
                    second.numero_da_solicitacao: "PPGEL +",
                },
            )
            original_bytes = current.read_bytes()

            build_candidate([second, new_trip, first], current, candidate)

            self.assertEqual(current.read_bytes(), original_bytes)
            rows = read_base_rows(candidate)
            self.assertEqual(
                list(rows), ["123456/26-2C", "234567/26-1C", "123456/26-1C"]
            )
            self.assertEqual(rows["123456/26-2C"][16], "PPGEL +")
            self.assertEqual(rows["123456/26-1C"][16], "AGRONOMIA")

    def test_refresh_leaves_new_pcdp_code_blank(self) -> None:
        existing = make_trip("123456/26")
        added = make_trip("234567/26")

        with TemporaryDirectory() as directory:
            current = Path(directory) / "current.xlsx"
            candidate = Path(directory) / "candidate.xlsx"
            write_existing_workbook(
                current, [existing], {existing.numero_da_solicitacao: "AGRONOMIA"}
            )

            build_candidate([existing, added], current, candidate)

            rows = read_base_rows(candidate)
            self.assertEqual(rows[existing.numero_da_solicitacao][16], "AGRONOMIA")
            self.assertIsNone(rows[added.numero_da_solicitacao][16])

    def test_refresh_preserves_manual_allocations_and_formula_cells(self) -> None:
        first = make_trip("123456/26")
        second = make_trip("234567/26-1C")
        allocations = {"AGRONOMIA": 12345.67, "PPGEL +": 987.65}

        with TemporaryDirectory() as directory:
            current = Path(directory) / "current.xlsx"
            candidate = Path(directory) / "candidate.xlsx"
            write_existing_workbook(
                current,
                [first, second],
                {first.numero_da_solicitacao: "AGRONOMIA"},
                allocations,
            )
            from openpyxl import load_workbook

            original = load_workbook(current, data_only=False)
            original_formulas = {
                cell: original["RESUMO GASTOS"][cell].value
                for cell in ("J2", "M2", "O2", "P2", "Q2", "R2")
            }
            original.close()

            build_candidate([second, first], current, candidate)

            refreshed = load_workbook(candidate, data_only=False)
            support = refreshed["APOIO"]
            actual_allocations = {
                support.cell(row, 1).value: support.cell(row, 4).value
                for row in range(2, support.max_row + 1)
            }
            self.assertEqual(
                {code: actual_allocations[code] for code in allocations}, allocations
            )
            self.assertEqual(
                {
                    cell: refreshed["RESUMO GASTOS"][cell].value
                    for cell in original_formulas
                },
                original_formulas,
            )
            self.assertIn("CodigosDebito", refreshed.defined_names)
            self.assertEqual(
                refreshed["BASE VIAGENS"].data_validations.dataValidation[0].formula1,
                "=CodigosDebito",
            )
            refreshed.close()

    def test_refresh_rejects_unknown_code(self) -> None:
        trip = make_trip()

        with TemporaryDirectory() as directory:
            current = Path(directory) / "current.xlsx"
            candidate = Path(directory) / "candidate.xlsx"
            write_existing_workbook(
                current, [trip], {trip.numero_da_solicitacao: "CODIGO DESCONHECIDO"}
            )
            original_bytes = current.read_bytes()

            with self.assertRaisesRegex(ValueError, "código de débito desconhecido"):
                build_candidate([trip], current, candidate)

            self.assertEqual(current.read_bytes(), original_bytes)
            self.assertFalse(candidate.exists())

    def test_refresh_rejects_duplicate_pcdp(self) -> None:
        trip = make_trip()

        with TemporaryDirectory() as directory:
            current = Path(directory) / "current.xlsx"
            candidate = Path(directory) / "candidate.xlsx"
            write_existing_workbook(current, [trip])
            original_bytes = current.read_bytes()

            with self.assertRaisesRegex(ValueError, "PCDP duplicada"):
                build_candidate([trip, trip], current, candidate)

            self.assertEqual(current.read_bytes(), original_bytes)
            self.assertFalse(candidate.exists())

    def test_refresh_rejects_missing_previous_pcdp_without_changing_source(
        self,
    ) -> None:
        first = make_trip("123456/26")
        second = make_trip("234567/26")

        with TemporaryDirectory() as directory:
            current = Path(directory) / "current.xlsx"
            candidate = Path(directory) / "candidate.xlsx"
            write_existing_workbook(current, [first, second])
            original_bytes = current.read_bytes()

            with self.assertRaises(WorkbookValidationError) as context:
                build_candidate([first], current, candidate)

            self.assertEqual(context.exception.missing_pcdps, ("234567/26",))
            self.assertEqual(current.read_bytes(), original_bytes)
            self.assertFalse(candidate.exists())

    def test_refresh_rejects_corrupt_workbook_without_changing_source(self) -> None:
        with TemporaryDirectory() as directory:
            current = Path(directory) / "corrupt.xlsx"
            candidate = Path(directory) / "candidate.xlsx"
            current.write_bytes(b"not an Excel workbook")
            original_bytes = current.read_bytes()

            with self.assertRaisesRegex(ValueError, "abrir o workbook"):
                build_candidate([make_trip()], current, candidate)

            self.assertEqual(current.read_bytes(), original_bytes)
            self.assertFalse(candidate.exists())

    def test_refresh_rejects_missing_sheet_workbook_without_changing_source(
        self,
    ) -> None:
        from openpyxl import Workbook

        with TemporaryDirectory() as directory:
            current = Path(directory) / "missing-sheet.xlsx"
            candidate = Path(directory) / "candidate.xlsx"
            Workbook().save(current)
            original_bytes = current.read_bytes()

            with self.assertRaisesRegex(WorkbookValidationError, "três worksheets"):
                build_candidate([make_trip()], current, candidate)

            self.assertEqual(current.read_bytes(), original_bytes)
            self.assertFalse(candidate.exists())

    def test_refresh_rejects_nonnumeric_allocation_without_changing_source(
        self,
    ) -> None:
        trip = make_trip()

        with TemporaryDirectory() as directory:
            current = Path(directory) / "current.xlsx"
            candidate = Path(directory) / "candidate.xlsx"
            write_existing_workbook(
                current, [trip], allocations={"AGRONOMIA": "dez mil"}
            )
            original_bytes = current.read_bytes()

            with self.assertRaisesRegex(WorkbookValidationError, "alocação inicial"):
                build_candidate([trip], current, candidate)

            self.assertEqual(current.read_bytes(), original_bytes)
            self.assertFalse(candidate.exists())

    def test_refresh_rejects_static_summary_value_without_changing_source(self) -> None:
        from openpyxl import load_workbook

        trip = make_trip()
        with TemporaryDirectory() as directory:
            current = Path(directory) / "current.xlsx"
            candidate = Path(directory) / "candidate.xlsx"
            create_workbook_template(current)
            workbook = load_workbook(current, data_only=False)
            workbook["RESUMO GASTOS"]["O2"] = 0
            workbook.save(current)
            workbook.close()
            original_bytes = current.read_bytes()

            with self.assertRaisesRegex(WorkbookValidationError, "fórmula"):
                build_candidate([trip], current, candidate)

            self.assertEqual(current.read_bytes(), original_bytes)
            self.assertFalse(candidate.exists())

    def test_refresh_rejects_changed_summary_formula_without_changing_source(
        self,
    ) -> None:
        from openpyxl import load_workbook

        trip = make_trip()
        with TemporaryDirectory() as directory:
            current = Path(directory) / "current.xlsx"
            candidate = Path(directory) / "candidate.xlsx"
            create_workbook_template(current)
            workbook = load_workbook(current, data_only=False)
            workbook["RESUMO GASTOS"]["G2"] = "=0"
            workbook.save(current)
            workbook.close()
            original_bytes = current.read_bytes()

            with self.assertRaisesRegex(WorkbookValidationError, "fórmula"):
                build_candidate([trip], current, candidate)

            self.assertEqual(current.read_bytes(), original_bytes)
            self.assertFalse(candidate.exists())


class WorkbookPublicationTests(unittest.TestCase):
    def test_default_workbook_path_is_checkout_relative(self) -> None:
        self.assertEqual(
            xlsx_output.DEFAULT_WORKBOOK,
            Path(xlsx_output.__file__).resolve().parents[1]
            / "output"
            / f"gastos_scdp_{current_year()}.xlsx",
        )

    def test_first_publish_creates_workbook_without_backup(self) -> None:
        with TemporaryDirectory() as directory:
            path = Path(directory) / "gastos.xlsx"

            backup = publish_workbook([make_trip()], path)

            self.assertIsNone(backup)
            self.assertTrue(path.exists())
            self.assertEqual(list(Path(directory).glob("*.backup.*.xlsx")), [])

    def test_refresh_creates_timestamped_backup_before_replace(self) -> None:
        trip = make_trip()

        with TemporaryDirectory() as directory:
            path = Path(directory) / "gastos.xlsx"
            self.assertIsNone(publish_workbook([trip], path))
            from openpyxl import load_workbook

            original = load_workbook(path)
            original["APOIO"]["D2"] = 5000.0
            original.save(path)
            original_bytes = path.read_bytes()

            backup = publish_workbook([trip], path)

            self.assertIsNotNone(backup)
            assert backup is not None
            self.assertTrue(backup.exists())
            self.assertIn(".backup.", backup.name)
            self.assertEqual(backup.read_bytes(), original_bytes)
            refreshed = load_workbook(path, data_only=False)
            self.assertEqual(refreshed["APOIO"]["D2"].value, 5000.0)

    def test_invalid_candidate_keeps_published_workbook_unchanged(self) -> None:
        trip = make_trip()

        with TemporaryDirectory() as directory:
            path = Path(directory) / "gastos.xlsx"
            publish_workbook([trip], path)
            original_bytes = path.read_bytes()

            with (
                patch(
                    "scdp_automation.xlsx_output._validate_candidate",
                    side_effect=ValueError("candidato inválido"),
                ),
                self.assertRaisesRegex(ValueError, "candidato inválido"),
            ):
                publish_workbook([trip], path)

            self.assertEqual(path.read_bytes(), original_bytes)
            self.assertEqual(list(Path(directory).glob("*.backup.*.xlsx")), [])
            self.assertEqual(list(Path(directory).glob("*.candidate.xlsx")), [])

    def test_backup_failure_keeps_published_workbook_unchanged(self) -> None:
        trip = make_trip()

        with TemporaryDirectory() as directory:
            path = Path(directory) / "gastos.xlsx"
            publish_workbook([trip], path)
            original_bytes = path.read_bytes()

            with (
                patch(
                    "scdp_automation.xlsx_output.shutil.copy2",
                    side_effect=OSError("falha de backup"),
                ),
                self.assertRaisesRegex(OSError, "workbook publicado permanece intacto"),
            ):
                publish_workbook([trip], path)

            self.assertEqual(path.read_bytes(), original_bytes)
            self.assertEqual(list(Path(directory).glob("*.backup.*.xlsx")), [])
            self.assertEqual(list(Path(directory).glob("*.candidate.xlsx")), [])

    def test_replace_failure_retains_validated_candidate(self) -> None:
        trip = make_trip()

        with TemporaryDirectory() as directory:
            path = Path(directory) / "gastos.xlsx"
            publish_workbook([trip], path)
            original_bytes = path.read_bytes()

            with (
                patch(
                    "scdp_automation.xlsx_output.os.replace",
                    side_effect=OSError("falha de promoção"),
                ),
                self.assertRaisesRegex(RuntimeError, "candidato validado"),
            ):
                publish_workbook([trip], path)

            self.assertEqual(path.read_bytes(), original_bytes)
            candidates = list(Path(directory).glob("*.candidate.xlsx"))
            self.assertEqual(len(candidates), 1)
            self.assertTrue(candidates[0].is_file())
            self.assertEqual(len(list(Path(directory).glob("*.backup.*.xlsx"))), 1)

    def test_locked_workbook_reports_close_and_retry(self) -> None:
        from filelock import FileLock

        with TemporaryDirectory() as directory:
            path = Path(directory) / "gastos.xlsx"
            from scdp_automation.xlsx_output import _workbook_lock_path

            lock = FileLock(_workbook_lock_path(path), timeout=0.1)
            with (
                lock,
                patch("scdp_automation.xlsx_output.FILELOCK_TIMEOUT", 0.01),
                self.assertRaisesRegex(TimeoutError, "feche-o e tente novamente"),
            ):
                publish_workbook([make_trip()], path)

    def test_concurrent_publishers_are_serialized(self) -> None:
        trip = make_trip()

        with TemporaryDirectory() as directory:
            path = Path(directory) / "gastos.xlsx"
            with ThreadPoolExecutor(max_workers=2) as executor:
                results = list(
                    executor.map(lambda _: publish_workbook([trip], path), range(2))
                )

            self.assertCountEqual([result is None for result in results], [True, False])
            self.assertTrue(path.exists())
            self.assertEqual(len(list(Path(directory).glob("*.backup.*.xlsx"))), 1)


if __name__ == "__main__":
    unittest.main()
