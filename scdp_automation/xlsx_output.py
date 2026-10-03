"""Domain mapping and debit categories for the annual spending workbook."""

from __future__ import annotations

import math
import os
import shutil
from collections.abc import Sequence
from copy import copy
from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path
from uuid import uuid4
from zipfile import BadZipFile

from filelock import FileLock, Timeout
from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.worksheet import Worksheet

from scdp_automation.relatorio import Viagem


@dataclass(frozen=True, slots=True)
class DebitCategory:
    """A stable debit key with its display name and budget segment."""

    code: str
    name: str
    segment: str
    review_required: bool = False


@dataclass(frozen=True, slots=True)
class TripSummary:
    """One row of aggregated values from a complete SCDP request."""

    pcdp: str
    proposed: str
    status: str
    daily_count: float
    daily_amount: float
    ticket_amount: float
    additional_amount: float
    discount_amount: float
    restitution_amount: float
    reimbursement_amount: float
    trip_total: float


DEBIT_CATEGORIES: tuple[DebitCategory, ...] = (
    DebitCategory("ADMINISTRAÇÃO", "Administração", "SEG 1 GRADUAÇÃO"),
    DebitCategory("AGRONOMIA", "Agronomia", "SEG 1 GRADUAÇÃO"),
    DebitCategory("C COMPUTAÇÃO", "Ciências da Computação", "SEG 1 GRADUAÇÃO"),
    DebitCategory("C ECONÔMICAS", "Ciências Econômicas", "SEG 1 GRADUAÇÃO"),
    DebitCategory("CIÊNCIAS SOCIAIS", "Ciências Sociais", "SEG 1 GRADUAÇÃO"),
    DebitCategory("ENFERMAGEM", "Enfermagem", "SEG 1 GRADUAÇÃO"),
    DebitCategory("ENG AMBIENTAL", "Engenharia Ambiental", "SEG 1 GRADUAÇÃO"),
    DebitCategory("ENGENHARIA CIVIL", "Engenharia Civil", "SEG 1 GRADUAÇÃO"),
    DebitCategory("FILOSOFIA", "Filosofia", "SEG 1 GRADUAÇÃO"),
    DebitCategory("GEOGRAFIA", "Geografia", "SEG 1 GRADUAÇÃO"),
    DebitCategory("HISTÓRIA", "História", "SEG 1 GRADUAÇÃO"),
    DebitCategory("LETRAS", "Letras", "SEG 1 GRADUAÇÃO"),
    DebitCategory("MATEMÁTICA", "Matemática", "SEG 1 GRADUAÇÃO"),
    DebitCategory("MEDICINA", "Medicina", "SEG 1 GRADUAÇÃO"),
    DebitCategory("PEDAGOGIA", "Pedagogia", "SEG 1 GRADUAÇÃO"),
    DebitCategory("Lato Oncologia", "LS Enf em Oncologia", "SEG 2 MESTRADO"),
    DebitCategory("PPGCB", "PPG Ciências Biomédicas", "SEG 2 MESTRADO"),
    DebitCategory("PPGE", "PPG Educação", "SEG 2 MESTRADO"),
    DebitCategory("PPGEL", "PPG Estudos Linguísticos", "SEG 2 MESTRADO"),
    DebitCategory("PPGEL +", "PPGEL +", "SEG 2 MESTRADO"),
    DebitCategory("PPGEnf", "PPG Enfermagem", "SEG 2 MESTRADO"),
    DebitCategory("PPGFil", "PPG Filosofia", "SEG 2 MESTRADO"),
    DebitCategory("PPGGeo", "PPG Geografia", "SEG 2 MESTRADO"),
    DebitCategory("PPGH", "PPG História", "SEG 2 MESTRADO"),
    DebitCategory("PPGDH", "PPGDH", "SEG 2 MESTRADO", review_required=True),
    DebitCategory("PROFIAP", "PROFIAP", "SEG 2 MESTRADO"),
    DebitCategory("PROFMAT", "PROFMAT", "SEG 2 MESTRADO"),
    DebitCategory("DIREÇÃO", "Geral (Direção/Coordenações)", "SEG 3 OUTROS"),
    DebitCategory("DIREÇÃO - AGAS", "DIREÇÃO - AGAS", "SEG 3 OUTROS"),
    DebitCategory("DIREÇÃO - Banca Libras", "DIREÇÃO - Banca Libras", "SEG 3 OUTROS"),
    DebitCategory("DIREÇÃO - CAAEX", "DIREÇÃO - CAAEX", "SEG 3 OUTROS"),
    DebitCategory("DIREÇÃO - Empr Junior", "DIREÇÃO - Empr Junior", "SEG 3 OUTROS"),
    DebitCategory(
        "DIREÇÃO - StartUp Summit", "DIREÇÃO - StartUp Summit", "SEG 3 OUTROS"
    ),
    DebitCategory(
        "DIREÇÃO - StartUp Weekend", "DIREÇÃO - StartUp Weekend", "SEG 3 OUTROS"
    ),
    DebitCategory("DIREÇÃO - Sunset", "DIREÇÃO - Sunset", "SEG 3 OUTROS"),
    DebitCategory("CAPPG - Res 49", "CAPPG - Res 49", "SEG 4 AUX EVENTOS"),
    DebitCategory("AFAST PAÍS", "Afastamento no país", "SEG 5 AFAST PAÍS"),
)


def summarize_trips(trips: Sequence[Viagem]) -> list[TripSummary]:
    """Map one summary row per unique full PCDP, using trip-level subtotals."""
    summaries: list[TripSummary] = []
    seen_pcdps: set[str] = set()

    for trip in trips:
        pcdp = trip.numero_da_solicitacao
        if pcdp in seen_pcdps:
            raise ValueError(f"PCDP duplicada na listagem: {pcdp}")
        seen_pcdps.add(pcdp)

        summaries.append(
            TripSummary(
                pcdp=pcdp,
                proposed=trip.nome_do_proposto,
                status=trip.situacao_da_viagem,
                daily_count=trip.sub_total.quantidade_diarias,
                daily_amount=trip.sub_total.diarias_r,
                ticket_amount=trip.sub_total.passagens_e_taxas_iniciais_r,
                additional_amount=trip.total_adicional_r,
                discount_amount=trip.descontos_r,
                restitution_amount=trip.restituicao_r,
                reimbursement_amount=trip.reembolso_r,
                trip_total=trip.total_da_viagem_r,
            )
        )

    return summaries


DEFAULT_WORKBOOK = (
    Path(__file__).resolve().parents[1] / "output" / "gastos_scdp_2026.xlsx"
)

BASE_HEADERS = (
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
)

SUPPORT_HEADERS = (
    "Código de débito",
    "Nome por extenso",
    "Segmento",
    "Alocação inicial (R$)",
)

SUMMARY_HEADERS = (
    "Segmento",
    "Código de débito",
    "Nome por extenso",
    "Alocação inicial (R$)",
    "Quantidade PCDPs",
    "Quantidade de diárias",
    "Diárias (R$)",
    "Adicional (R$)",
    "Descontos (R$)",
    "TOTAL DIÁRIAS (R$)",
    "Passagens (R$)",
    "Restituições (R$)",
    "PASS AÉREA+ROD (R$)",
    "Reembolsos (R$)",
    "TOTAL DA VIAGEM (R$)",
    "CANCELADAS (R$)",
    "TOTAL UTILIZADO (R$)",
    "SALDO DISPONÍVEL (R$)",
)

CODE_LIST_NAME = "CodigosDebito"
BASE_TABLE_NAME = "tblBaseViagens"
SUPPORT_TABLE_NAME = "tblApoioDebito"
SUMMARY_TABLE_NAME = "tblResumoGastos"
_BASE_LAST_ROW = 1_048_576
FILELOCK_TIMEOUT = 30


class WorkbookValidationError(ValueError):
    """The published workbook or refresh candidate is inconsistent."""

    def __init__(self, message: str, *, missing_pcdps: tuple[str, ...] = ()) -> None:
        super().__init__(message)
        self.missing_pcdps = missing_pcdps


class WorkbookLockedError(TimeoutError):
    """The output lock could not be acquired before its timeout."""


class WorkbookBackupError(OSError):
    """The existing workbook could not be backed up before replacement."""


class WorkbookPublishError(RuntimeError):
    """A validated workbook could not be promoted to the output path."""


def _add_table(worksheet: Worksheet, display_name: str, ref: str) -> None:
    table = Table(displayName=display_name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    worksheet.add_table(table)


def _style_header(worksheet: Worksheet) -> None:
    for cell in worksheet[1]:
        cell.fill = PatternFill(fill_type="solid", fgColor="1F4E78")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    worksheet.row_dimensions[1].height = 30
    worksheet.freeze_panes = "A2"
    worksheet.sheet_view.showGridLines = False


def _segment_formula(row: int, support_last_row: int) -> str:
    return (
        f'=IF($M{row}="","",IFERROR(INDEX(\'APOIO\'!$C$2:$C${support_last_row},'
        f"MATCH($M{row},'APOIO'!$A$2:$A${support_last_row},0)),\"\"))"
    )


def _sumifs_formula(sum_column: str, row: int, *, canceled_only: bool = False) -> str:
    formula = (
        f"=SUMIFS('BASE VIAGENS'!${sum_column}:${sum_column},"
        f"'BASE VIAGENS'!$L:$L,$A{row},"
        f"'BASE VIAGENS'!$M:$M,$B{row}"
    )
    if canceled_only:
        formula += ",'BASE VIAGENS'!$C:$C,\"*Cancel*\""
    return formula + ")"


def _unclassified_sumifs_formula(
    sum_column: str, *, canceled_only: bool = False
) -> str:
    formula = (
        f"=SUMIFS('BASE VIAGENS'!${sum_column}:${sum_column},"
        "'BASE VIAGENS'!$M:$M,\"\",'BASE VIAGENS'!$A:$A,\"<>\""
    )
    if canceled_only:
        formula += ",'BASE VIAGENS'!$C:$C,\"*Cancel*\""
    return formula + ")"


def _countifs_formula(row: int) -> str:
    return (
        f"=COUNTIFS('BASE VIAGENS'!$L:$L,$A{row},"
        f"'BASE VIAGENS'!$M:$M,$B{row},'BASE VIAGENS'!$A:$A,\"<>\")"
    )


def build_summary_formulas(workbook: Workbook) -> None:
    """Build formula-only totals grouped by segment and debit code."""
    summary = workbook["RESUMO GASTOS"]
    support = workbook["APOIO"]

    if SUMMARY_TABLE_NAME in summary.tables:
        del summary.tables[SUMMARY_TABLE_NAME]
    if summary.max_row > 1:
        summary.delete_rows(2, summary.max_row - 1)
    for column, header in enumerate(SUMMARY_HEADERS, start=1):
        summary.cell(1, column, header)

    row = 2
    for support_row in range(2, support.max_row + 1):
        code = support.cell(support_row, 1).value
        if not code:
            continue

        summary.cell(row, 1, f"='APOIO'!$C${support_row}")
        summary.cell(row, 2, f"='APOIO'!$A${support_row}")
        summary.cell(row, 3, f"='APOIO'!$B${support_row}")
        summary.cell(
            row,
            4,
            f"=IF('APOIO'!$D${support_row}=\"\",\"Pendente\",'APOIO'!$D${support_row})",
        )
        summary.cell(row, 5, _countifs_formula(row))
        summary.cell(row, 6, _sumifs_formula("D", row))
        summary.cell(row, 7, _sumifs_formula("E", row))
        summary.cell(row, 8, _sumifs_formula("G", row))
        summary.cell(row, 9, _sumifs_formula("H", row))
        summary.cell(row, 10, f"=G{row}+H{row}-I{row}")
        summary.cell(row, 11, _sumifs_formula("F", row))
        summary.cell(row, 12, _sumifs_formula("I", row))
        summary.cell(row, 13, f"=K{row}+L{row}")
        summary.cell(row, 14, _sumifs_formula("J", row))
        summary.cell(row, 15, _sumifs_formula("K", row))
        summary.cell(row, 16, _sumifs_formula("K", row, canceled_only=True))
        summary.cell(row, 17, f"=O{row}")
        summary.cell(row, 18, f'=IF(D{row}="Pendente","Pendente",D{row}-Q{row})')
        row += 1

    summary.cell(row, 1, "SEM CLASSIFICAÇÃO")
    summary.cell(row, 2, None)
    summary.cell(row, 3, "PCDPs sem código de débito")
    summary.cell(row, 4, "Pendente")
    summary.cell(
        row,
        5,
        "=COUNTIFS('BASE VIAGENS'!$M:$M,\"\",'BASE VIAGENS'!$A:$A,\"<>\")",
    )
    summary.cell(row, 6, _unclassified_sumifs_formula("D"))
    summary.cell(row, 7, _unclassified_sumifs_formula("E"))
    summary.cell(row, 8, _unclassified_sumifs_formula("G"))
    summary.cell(row, 9, _unclassified_sumifs_formula("H"))
    summary.cell(row, 10, f"=G{row}+H{row}-I{row}")
    summary.cell(row, 11, _unclassified_sumifs_formula("F"))
    summary.cell(row, 12, _unclassified_sumifs_formula("I"))
    summary.cell(row, 13, f"=K{row}+L{row}")
    summary.cell(row, 14, _unclassified_sumifs_formula("J"))
    summary.cell(row, 15, _unclassified_sumifs_formula("K"))
    summary.cell(row, 16, _unclassified_sumifs_formula("K", canceled_only=True))
    summary.cell(row, 17, f"=O{row}")
    summary.cell(row, 18, "Pendente")

    _style_header(summary)
    for column in (4, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18):
        for cells in summary.iter_cols(
            min_col=column, max_col=column, min_row=2, max_row=row
        ):
            for cell in cells:
                cell.number_format = '"R$" #,##0.00;[Red]-"R$" #,##0.00'
    for column, width in enumerate(
        (23, 30, 38, 22, 18, 22, 18, 18, 18, 22, 18, 18, 24, 18, 24, 22, 24, 24),
        start=1,
    ):
        summary.column_dimensions[summary.cell(1, column).column_letter].width = width

    _add_table(summary, SUMMARY_TABLE_NAME, f"A1:R{row}")


def create_workbook_template(path: Path) -> None:
    """Create the formula-driven three-sheet annual spending workbook."""
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    base = workbook.active
    base.title = "BASE VIAGENS"
    support = workbook.create_sheet("APOIO")
    summary = workbook.create_sheet("RESUMO GASTOS")

    for column, header in enumerate(BASE_HEADERS, start=1):
        base.cell(1, column, header)
    base.append([None] * len(BASE_HEADERS))

    for column, header in enumerate(SUPPORT_HEADERS, start=1):
        support.cell(1, column, header)
    for category in DEBIT_CATEGORIES:
        support.append([category.code, category.name, category.segment, None])
    ppghd_row = next(
        row
        for row in range(2, support.max_row + 1)
        if support.cell(row, 1).value == "PPGDH"
    )
    support.cell(ppghd_row, 2).comment = Comment(
        "A planilha de referência apresenta descrições conflitantes para PPGDH. "
        "Confirme o nome e a alocação antes de preencher.",
        "SCDP",
    )

    workbook.defined_names.add(
        DefinedName(
            CODE_LIST_NAME,
            attr_text=f"'APOIO'!$A$2:$A${support.max_row}",
        )
    )
    code_validation = DataValidation(
        type="list",
        formula1=f"={CODE_LIST_NAME}",
        allow_blank=True,
        showErrorMessage=True,
        errorStyle="stop",
        errorTitle="Código de débito inválido",
        error="Selecione um código existente na worksheet APOIO.",
        showInputMessage=True,
        promptTitle="Classificação da PCDP",
        prompt="Escolha o curso, programa ou setor que receberá o débito.",
    )
    code_validation.add(f"M2:M{_BASE_LAST_ROW}")
    base.add_data_validation(code_validation)

    base.cell(2, 12, _segment_formula(2, support.max_row))
    for column, header in enumerate(SUMMARY_HEADERS, start=1):
        summary.cell(1, column, header)

    _style_header(base)
    _style_header(support)
    for column, width in enumerate(
        (18, 34, 24, 22, 18, 18, 18, 18, 18, 18, 24, 24, 34), start=1
    ):
        base.column_dimensions[base.cell(1, column).column_letter].width = width
    for column, width in enumerate((34, 40, 26, 24), start=1):
        support.column_dimensions[support.cell(1, column).column_letter].width = width
    for row in range(2, support.max_row + 1):
        support.cell(row, 4).number_format = '"R$" #,##0.00;[Red]-"R$" #,##0.00'

    _add_table(base, BASE_TABLE_NAME, "A1:M2")
    _add_table(support, SUPPORT_TABLE_NAME, f"A1:D{support.max_row}")
    build_summary_formulas(workbook)

    workbook.calculation.calcMode = "auto"
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.save(path)


def _check_workbook_structure(workbook: Workbook) -> None:
    if workbook.sheetnames != ["BASE VIAGENS", "APOIO", "RESUMO GASTOS"]:
        raise WorkbookValidationError(
            "O workbook não contém as três worksheets esperadas."
        )

    base = workbook["BASE VIAGENS"]
    support = workbook["APOIO"]
    summary = workbook["RESUMO GASTOS"]
    if tuple(cell.value for cell in base[1]) != BASE_HEADERS:
        raise WorkbookValidationError("As colunas de BASE VIAGENS foram alteradas.")
    if tuple(cell.value for cell in support[1]) != SUPPORT_HEADERS:
        raise WorkbookValidationError("As colunas de APOIO foram alteradas.")
    if tuple(cell.value for cell in summary[1]) != SUMMARY_HEADERS:
        raise WorkbookValidationError("As colunas de RESUMO GASTOS foram alteradas.")

    expected_categories = {category.code: category for category in DEBIT_CATEGORIES}
    actual_codes: set[str] = set()
    for row in range(2, support.max_row + 1):
        code = support.cell(row, 1).value
        if code is None:
            continue
        if not isinstance(code, str) or code in actual_codes:
            raise WorkbookValidationError(
                "A tabela APOIO contém códigos inválidos ou duplicados."
            )
        actual_codes.add(code)
        expected = expected_categories.get(code)
        if expected is None or (
            support.cell(row, 2).value != expected.name
            or support.cell(row, 3).value != expected.segment
        ):
            raise WorkbookValidationError(
                "O catálogo de débitos em APOIO foi alterado."
            )
    if actual_codes != set(expected_categories):
        raise WorkbookValidationError("O catálogo de débitos em APOIO está incompleto.")
    for row in range(2, support.max_row + 1):
        allocation = support.cell(row, 4).value
        if allocation is not None and (
            isinstance(allocation, bool)
            or not isinstance(allocation, (int, float))
            or not math.isfinite(allocation)
        ):
            raise WorkbookValidationError(
                "A alocação inicial em APOIO precisa ser numérica."
            )

    expected_range = f"'APOIO'!$A$2:$A${support.max_row}"
    defined_name = workbook.defined_names.get(CODE_LIST_NAME)
    if defined_name is None or defined_name.attr_text != expected_range:
        raise WorkbookValidationError(
            "A lista nomeada de códigos de débito está ausente."
        )
    validations = [
        validation
        for validation in base.data_validations.dataValidation
        if validation.type == "list" and validation.formula1 == f"={CODE_LIST_NAME}"
    ]
    if len(validations) != 1 or str(validations[0].sqref) != f"M2:M{_BASE_LAST_ROW}":
        raise WorkbookValidationError(
            "A validação de códigos em BASE VIAGENS está ausente."
        )

    required_tables = {
        (base, BASE_TABLE_NAME),
        (support, SUPPORT_TABLE_NAME),
        (summary, SUMMARY_TABLE_NAME),
    }
    if any(
        table_name not in worksheet.tables for worksheet, table_name in required_tables
    ):
        raise WorkbookValidationError("Uma tabela necessária foi removida do workbook.")
    if base.tables[BASE_TABLE_NAME].ref.split(":")[0] != "A1":
        raise WorkbookValidationError("A tabela BASE VIAGENS está malformada.")
    if support.tables[SUPPORT_TABLE_NAME].ref != f"A1:D{support.max_row}":
        raise WorkbookValidationError("A tabela APOIO está malformada.")
    if summary.tables[SUMMARY_TABLE_NAME].ref != f"A1:R{summary.max_row}":
        raise WorkbookValidationError("A tabela RESUMO GASTOS está malformada.")

    for row in range(2, len(DEBIT_CATEGORIES) + 2):
        for column in range(1, len(SUMMARY_HEADERS) + 1):
            value = summary.cell(row, column).value
            if not isinstance(value, str) or not value.startswith("="):
                raise WorkbookValidationError(
                    "Há uma fórmula ausente ou um valor estático na worksheet RESUMO GASTOS."
                )
        if not summary.cell(row, 15).value.startswith("=SUMIFS("):
            raise WorkbookValidationError("O total da viagem do resumo não usa SUMIFS.")
        if summary.cell(row, 17).value != f"=O{row}":
            raise WorkbookValidationError(
                "O total utilizado do resumo não possui fórmula."
            )
        if summary.cell(row, 10).value != f"=G{row}+H{row}-I{row}":
            raise WorkbookValidationError(
                "O total de diárias do resumo está incorreto."
            )
        if summary.cell(row, 13).value != f"=K{row}+L{row}":
            raise WorkbookValidationError(
                "O total de passagens e restituições está incorreto."
            )
        if "*Cancel*" not in summary.cell(row, 16).value:
            raise WorkbookValidationError(
                "O resumo não calcula as solicitações canceladas."
            )
    pending_row = len(DEBIT_CATEGORIES) + 2
    if summary.cell(pending_row, 1).value != "SEM CLASSIFICAÇÃO":
        raise WorkbookValidationError(
            "A linha de PCDPs sem classificação foi removida."
        )
    if not summary.cell(pending_row, 5).value.startswith("=COUNTIFS("):
        raise WorkbookValidationError("A contagem de PCDPs sem código está ausente.")
    if not summary.cell(pending_row, 15).value.startswith("=SUMIFS("):
        raise WorkbookValidationError("O total de PCDPs sem código está ausente.")
    if summary.cell(pending_row, 17).value != f"=O{pending_row}":
        raise WorkbookValidationError(
            "O total utilizado das PCDPs sem código está incorreto."
        )
    if (
        workbook.calculation.calcMode != "auto"
        or not workbook.calculation.fullCalcOnLoad
        or not workbook.calculation.forceFullCalc
    ):
        raise WorkbookValidationError(
            "O workbook não solicita recálculo completo no Excel."
        )


def _load_workbook_for_refresh(path: Path) -> Workbook:
    try:
        workbook = load_workbook(path, data_only=False)
    except (OSError, InvalidFileException, BadZipFile, KeyError, ValueError) as error:
        raise WorkbookValidationError(
            f"Não foi possível abrir o workbook existente em {path}."
        ) from error
    try:
        _check_workbook_structure(workbook)
    except Exception:
        workbook.close()
        raise
    return workbook


def _write_base_rows(
    workbook: Workbook,
    summaries: Sequence[TripSummary],
    manual_codes: dict[str, str],
) -> None:
    base = workbook["BASE VIAGENS"]
    support_last_row = workbook["APOIO"].max_row
    old_last_row = max(base.max_row, 2)
    target_last_row = max(2, len(summaries) + 1)
    style_source = [copy(cell._style) for cell in base[2]]
    height_source = base.row_dimensions[2].height

    if old_last_row > target_last_row:
        base.delete_rows(target_last_row + 1, old_last_row - target_last_row)

    for index, row in enumerate(range(2, target_last_row + 1)):
        if row > old_last_row:
            for column, style in enumerate(style_source, start=1):
                base.cell(row, column)._style = copy(style)
            if height_source is not None:
                base.row_dimensions[row].height = height_source

        if index >= len(summaries):
            values: tuple[object, ...] = (None,) * 11
            pcdp = ""
        else:
            summary = summaries[index]
            pcdp = summary.pcdp
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
        base.cell(row, 12, _segment_formula(row, support_last_row))
        base.cell(row, 13, manual_codes.get(pcdp))

    base.tables[BASE_TABLE_NAME].ref = f"A1:M{target_last_row}"


def build_candidate(
    trips: Sequence[Viagem], current_path: Path, candidate_path: Path
) -> None:
    """Write a candidate snapshot while leaving the published workbook untouched."""
    if current_path.resolve() == candidate_path.resolve():
        raise WorkbookValidationError(
            "O candidato deve ter caminho separado do workbook publicado."
        )
    if candidate_path.exists():
        raise WorkbookValidationError(
            f"O caminho de candidato já existe: {candidate_path}"
        )

    summaries = summarize_trips(trips)
    incoming_pcdps = {summary.pcdp for summary in summaries}
    manual_codes: dict[str, str] = {}
    workbook: Workbook

    if current_path.exists():
        workbook = _load_workbook_for_refresh(current_path)
        base = workbook["BASE VIAGENS"]
        support = workbook["APOIO"]
        support_codes = {
            support.cell(row, 1).value
            for row in range(2, support.max_row + 1)
            if support.cell(row, 1).value
        }
        existing_pcdps: set[str] = set()
        for row in range(2, base.max_row + 1):
            pcdp = base.cell(row, 1).value
            if not pcdp:
                continue
            if not isinstance(pcdp, str) or pcdp in existing_pcdps:
                workbook.close()
                raise WorkbookValidationError(
                    "A base publicada contém PCDPs inválidas ou duplicadas."
                )
            existing_pcdps.add(pcdp)
            code = base.cell(row, 13).value
            if code in (None, ""):
                continue
            if not isinstance(code, str) or code not in support_codes:
                workbook.close()
                raise WorkbookValidationError(
                    "A base publicada contém código de débito desconhecido."
                )
            manual_codes[pcdp] = code

        missing_pcdps = tuple(sorted(existing_pcdps - incoming_pcdps))
        if missing_pcdps:
            workbook.close()
            raise WorkbookValidationError(
                f"A listagem completa está sem {len(missing_pcdps)} PCDP(s) já publicadas; "
                "reconcilie o relatório antes de atualizar.",
                missing_pcdps=missing_pcdps,
            )
    else:
        candidate_path.parent.mkdir(parents=True, exist_ok=True)
        create_workbook_template(candidate_path)
        workbook = _load_workbook_for_refresh(candidate_path)

    try:
        candidate_path.parent.mkdir(parents=True, exist_ok=True)
        _write_base_rows(workbook, summaries, manual_codes)
        workbook.save(candidate_path)
    except Exception:
        workbook.close()
        candidate_path.unlink(missing_ok=True)
        raise
    workbook.close()


def _validate_candidate(candidate_path: Path, trips: Sequence[Viagem]) -> None:
    summaries = summarize_trips(trips)
    workbook = _load_workbook_for_refresh(candidate_path)
    try:
        _check_workbook_structure(workbook)
        base = workbook["BASE VIAGENS"]
        support = workbook["APOIO"]
        support_codes = {
            support.cell(row, 1).value
            for row in range(2, support.max_row + 1)
            if support.cell(row, 1).value
        }

        actual_pcdps: list[str] = []
        for row in range(2, base.max_row + 1):
            pcdp = base.cell(row, 1).value
            if not pcdp:
                continue
            if not isinstance(pcdp, str) or pcdp in actual_pcdps:
                raise WorkbookValidationError(
                    "O candidato contém PCDPs inválidas ou duplicadas."
                )
            actual_pcdps.append(pcdp)
            code = base.cell(row, 13).value
            if code not in (None, "") and code not in support_codes:
                raise WorkbookValidationError(
                    "O candidato contém código de débito desconhecido."
                )
            segment_formula = base.cell(row, 12).value
            if not isinstance(segment_formula, str) or not segment_formula.startswith(
                "=IF("
            ):
                raise WorkbookValidationError(
                    "A fórmula Segmento está ausente na base."
                )

        if actual_pcdps != [summary.pcdp for summary in summaries]:
            raise WorkbookValidationError(
                "As PCDPs do candidato não correspondem à listagem completa."
            )

        for row, summary in enumerate(summaries, start=2):
            actual = tuple(base.cell(row, column).value for column in range(1, 12))
            expected = (
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
            if actual != expected:
                raise WorkbookValidationError(
                    "Os campos agregados do candidato estão inconsistentes."
                )
    finally:
        workbook.close()


def _candidate_path_for(workbook_path: Path) -> Path:
    return workbook_path.with_name(
        f"{workbook_path.stem}.{uuid4().hex}.candidate{workbook_path.suffix}"
    )


def _backup_path_for(workbook_path: Path) -> Path:
    timestamp = datetime.now(UTC).strftime("%Y%m%dT%H%M%S%fZ")
    backup = workbook_path.with_name(
        f"{workbook_path.stem}.backup.{timestamp}{workbook_path.suffix}"
    )
    if backup.exists():
        backup = backup.with_name(f"{backup.stem}.{uuid4().hex[:8]}{backup.suffix}")
    return backup


def publish_workbook(
    trips: Sequence[Viagem], workbook_path: Path = DEFAULT_WORKBOOK
) -> Path | None:
    """Serialize refreshes, validate a neighboring candidate, then replace atomically."""
    workbook_path.parent.mkdir(parents=True, exist_ok=True)
    lock = FileLock(f"{workbook_path}.lock", timeout=FILELOCK_TIMEOUT)
    try:
        with lock:
            candidate = _candidate_path_for(workbook_path)
            candidate_retained = False
            try:
                build_candidate(trips, workbook_path, candidate)
                _validate_candidate(candidate, trips)
                backup: Path | None = None
                if workbook_path.exists():
                    backup = _backup_path_for(workbook_path)
                    try:
                        shutil.copy2(workbook_path, backup)
                    except OSError as error:
                        backup.unlink(missing_ok=True)
                        raise WorkbookBackupError(
                            f"Não foi possível criar o backup em {backup}; "
                            f"o workbook publicado permanece intacto: {error}"
                        ) from error
                try:
                    os.replace(candidate, workbook_path)
                except OSError as error:
                    candidate_retained = True
                    raise WorkbookPublishError(
                        "Não foi possível substituir o workbook; feche o arquivo no Excel e "
                        f"tente novamente. O candidato validado permanece em {candidate}."
                    ) from error
                return backup
            except Exception:
                if not candidate_retained:
                    candidate.unlink(missing_ok=True)
                raise
    except Timeout as error:
        raise WorkbookLockedError(
            f"O workbook está em uso por outra atualização; feche-o e tente novamente: {workbook_path}"
        ) from error
