import unittest
from collections.abc import Mapping
from concurrent.futures import ThreadPoolExecutor
from dataclasses import FrozenInstanceError
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from scdp_automation import xlsx_output
from scdp_automation.relatorio import Viagem
from scdp_automation.xlsx_output import (
    BASE_HEADERS,
    DEBIT_CATEGORIES,
    DebitCategory,
    TripSummary,
    WorkbookValidationError,
    build_candidate,
    create_workbook_template,
    publish_workbook,
    summarize_trips,
)


def make_trip(pcdp: str = "123456/26-2B") -> Viagem:
    """Create a validated trip whose itinerary differs from its totals."""
    return Viagem.model_validate(
        {
            "numero_da_solicitacao": pcdp,
            "nome_do_proposto": "Pessoa Exemplo",
            "orgao_solicitante": "ORG-TESTE",
            "orgao_superior": "ORG-SUPERIOR-TESTE",
            "tipo_da_viagem": "NACIONAL",
            "situacao_da_viagem": "Autorizada",
            "motivo_viagem": "Nacional - A Serviço",
            "trechos": [
                {
                    "inicio": "01/03/2026",
                    "termino": "03/03/2026",
                    "origem": "Cidade Alfa (AA)",
                    "destino": "Cidade Beta (BB)",
                    "meio_de_transporte": "Aéreo",
                    "quantidade_diarias": 2.0,
                    "diarias_r": 700.25,
                    "passagens_e_taxas_iniciais_r": 800.50,
                    "total_r": 1500.75,
                },
                {
                    "inicio": "03/03/2026",
                    "termino": "05/03/2026",
                    "origem": "Cidade Beta (BB)",
                    "destino": "Cidade Alfa (AA)",
                    "meio_de_transporte": "Aéreo",
                    "quantidade_diarias": 3.0,
                    "diarias_r": 900.75,
                    "passagens_e_taxas_iniciais_r": 1000.25,
                    "total_r": 1901.00,
                },
            ],
            "custo_com_bilhetes_remarcados_nao_utilizados_cancelados_r": {
                "passagens_e_taxas_iniciais_r": 2.0,
                "total_r": 2.0,
            },
            "sub_total": {
                "quantidade_diarias": 8.5,
                "diarias_r": 1234.56,
                "passagens_e_taxas_iniciais_r": 789.01,
                "total_r": 2023.57,
            },
            "total_adicional_r": 33.05,
            "descontos_r": 4.01,
            "restituicao_r": 5.02,
            "reembolso_r": 6.03,
            "total_da_viagem_r": 2063.66,
        }
    )


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
            ),
        )

    def test_summarize_trips_rejects_duplicate_full_pcdp(self) -> None:
        with self.assertRaisesRegex(ValueError, "123456/26-2B"):
            summarize_trips([make_trip(), make_trip()])

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
                "PPGH",
                "PPGDH",
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
        self.assertEqual(category_by_code["Lato Oncologia"].name, "LS Enf em Oncologia")
        self.assertEqual(category_by_code["PPGEL +"].name, "PPGEL +")
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
            "PPGDH",
            "PROFMAT",
        ):
            self.assertIn(code, category_by_code)

        ppghd = category_by_code["PPGDH"]
        self.assertTrue(ppghd.review_required)
        self.assertEqual(ppghd.name, "PPGDH")


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
            "Segmento",
            "Código de débito",
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
        self.assertEqual(workbook["BASE VIAGENS"].cell(1, 13).value, "Código de débito")
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
            if support.cell(row, 1).value == "PPGDH"
        )
        self.assertEqual(support.cell(ppghd_row, 2).value, "PPGDH")
        self.assertIsNotNone(support.cell(ppghd_row, 2).comment)
        self.assertIn("conflit", support.cell(ppghd_row, 2).comment.text.lower())

    def test_base_code_column_has_dropdown_validation(self) -> None:
        with TemporaryDirectory() as directory:
            workbook = self.create_template(Path(directory))

        base = workbook["BASE VIAGENS"]
        validations = [
            validation
            for validation in base.data_validations.dataValidation
            if validation.type == "list"
        ]
        self.assertEqual(len(validations), 1)
        self.assertEqual(validations[0].formula1, "=CodigosDebito")
        self.assertEqual(str(validations[0].sqref), "M2:M1048576")

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
        self.assertIn("'BASE VIAGENS'!$L:$L,$A2", summary["O2"].value)
        self.assertIn("'BASE VIAGENS'!$M:$M,$B2", summary["O2"].value)
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


def write_existing_workbook(
    path: Path,
    trips: list[Viagem],
    codes: dict[str, str] | None = None,
    allocations: Mapping[str, object] | None = None,
) -> None:
    from openpyxl import load_workbook
    from openpyxl.formula.translate import Translator

    create_workbook_template(path)
    workbook = load_workbook(path, data_only=False)
    base = workbook["BASE VIAGENS"]
    codes = codes or {}
    allocations = allocations or {}

    for row, summary in enumerate(summarize_trips(trips), start=2):
        if row > 2:
            for column in range(1, len(BASE_HEADERS) + 1):
                base.cell(row, column)._style = base.cell(2, column)._style
        values = (
            summary.pcdp,
            summary.proposed,
            summary.status,
            summary.daily_count,
            summary.daily_amount,
            summary.ticket_amount,
            summary.additional_amount,
            summary.discount_amount,
            summary.restitution_amount,
            summary.reimbursement_amount,
            summary.trip_total,
        )
        for column, value in enumerate(values, start=1):
            base.cell(row, column, value)
        base.cell(
            row,
            12,
            Translator(base["L2"].value, origin="L2").translate_formula(f"L{row}"),
        )
        base.cell(row, 13, codes.get(summary.pcdp))

    base.tables["tblBaseViagens"].ref = f"A1:M{max(2, len(trips) + 1)}"
    support = workbook["APOIO"]
    for row in range(2, support.max_row + 1):
        code = support.cell(row, 1).value
        if code in allocations:
            support.cell(row, 4, allocations[code])
    workbook.save(path)


def read_base_rows(path: Path) -> dict[str, tuple[object, ...]]:
    from openpyxl import load_workbook

    workbook = load_workbook(path, data_only=False)
    base = workbook["BASE VIAGENS"]
    rows = {
        base.cell(row, 1).value: tuple(
            base.cell(row, column).value for column in range(1, 14)
        )
        for row in range(2, base.max_row + 1)
        if base.cell(row, 1).value
    }
    workbook.close()
    return rows


class WorkbookRefreshTests(unittest.TestCase):
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
            self.assertEqual(rows["123456/26-2C"][12], "PPGEL +")
            self.assertEqual(rows["123456/26-1C"][12], "AGRONOMIA")

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
            self.assertEqual(rows[existing.numero_da_solicitacao][12], "AGRONOMIA")
            self.assertIsNone(rows[added.numero_da_solicitacao][12])

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


class WorkbookPublicationTests(unittest.TestCase):
    def test_default_workbook_path_is_checkout_relative(self) -> None:
        self.assertEqual(
            xlsx_output.DEFAULT_WORKBOOK,
            Path(xlsx_output.__file__).resolve().parents[1]
            / "output"
            / "gastos_scdp_2026.xlsx",
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
            lock = FileLock(f"{path}.lock", timeout=0.1)
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
