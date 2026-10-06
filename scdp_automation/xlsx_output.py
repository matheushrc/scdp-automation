"""Domain mapping and debit categories for the annual spending workbook."""

from __future__ import annotations

import math
import os
import shutil
from collections.abc import Sequence
from copy import copy
from datetime import UTC, datetime
from pathlib import Path
from typing import cast
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

from scdp_automation.config import current_year
from scdp_automation.relatorio import Viagem
from scdp_automation.xlsx_models import (
    DEBIT_CATEGORIES,
    DebitCategory,
    TripSummary,
    summarize_trips,
)

__all__ = ["DEBIT_CATEGORIES", "DebitCategory", "TripSummary", "summarize_trips"]

DEFAULT_WORKBOOK = (
    Path(__file__).resolve().parents[1]
    / "output"
    / f"gastos_scdp_{current_year()}.xlsx"
)

_CHECKOUT = Path(__file__).resolve().parents[1]
_REFERENCE_ROOT = (
    _CHECKOUT.parent.parent if _CHECKOUT.parent.name == ".worktrees" else _CHECKOUT
)
DEFAULT_REFERENCE = _REFERENCE_ROOT / "input" / "gastos_scdp_template.xlsx"


LEGACY_BASE_HEADERS = (
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
    "Segmento",
    "Código de débito",
    "Descontar do curso?",
)

BASE_HEADERS = (
    *LEGACY_BASE_HEADERS[:14],
    "Descrição do pedido",
    *LEGACY_BASE_HEADERS[14:],
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
        f'=IF($P{row}="","",IFERROR(INDEX(\'APOIO\'!$C$2:$C${support_last_row},'
        f"MATCH($P{row},'APOIO'!$A$2:$A${support_last_row},0)),\"\"))"
    )


def _sumifs_formula(sum_column: str, row: int, *, canceled_only: bool = False) -> str:
    formula = (
        f"=SUMIFS('BASE VIAGENS'!${sum_column}:${sum_column},"
        f"'BASE VIAGENS'!$O:$O,$A{row},"
        f"'BASE VIAGENS'!$P:$P,$B{row}"
    )
    if canceled_only:
        formula += ",'BASE VIAGENS'!$C:$C,\"*Cancel*\""
    return formula + ")"


def _unclassified_sumifs_formula(
    sum_column: str, *, canceled_only: bool = False
) -> str:
    formula = (
        f"=SUMIFS('BASE VIAGENS'!${sum_column}:${sum_column},"
        "'BASE VIAGENS'!$P:$P,\"\",'BASE VIAGENS'!$A:$A,\"<>\""
    )
    if canceled_only:
        formula += ",'BASE VIAGENS'!$C:$C,\"*Cancel*\""
    return formula + ")"


def _countifs_formula(row: int) -> str:
    return (
        f"=COUNTIFS('BASE VIAGENS'!$O:$O,$A{row},"
        f"'BASE VIAGENS'!$P:$P,$B{row},'BASE VIAGENS'!$A:$A,\"<>\")"
    )


def _category_summary_formulas(row: int, support_row: int) -> tuple[str, ...]:
    return (
        f"='APOIO'!$C${support_row}",
        f"='APOIO'!$A${support_row}",
        f"='APOIO'!$B${support_row}",
        f"=IF('APOIO'!$D${support_row}=\"\",\"Pendente\",'APOIO'!$D${support_row})",
        _countifs_formula(row),
        _sumifs_formula("D", row),
        _sumifs_formula("E", row),
        _sumifs_formula("G", row),
        _sumifs_formula("H", row),
        f"=G{row}+H{row}-I{row}",
        _sumifs_formula("F", row),
        _sumifs_formula("I", row),
        f"=K{row}+L{row}",
        _sumifs_formula("J", row),
        _sumifs_formula("K", row),
        _sumifs_formula("K", row, canceled_only=True),
        f"=O{row}",
        f'=IF(D{row}="Pendente","Pendente",D{row}-Q{row})',
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
        "=COUNTIFS('BASE VIAGENS'!$P:$P,\"\",'BASE VIAGENS'!$A:$A,\"<>\")",
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


def _is_current_reference(path: Path) -> bool:
    if not path.exists():
        return False
    workbook = load_workbook(path, read_only=True)
    try:
        return "BASE VIAGENS" in workbook.sheetnames
    finally:
        workbook.close()


def create_workbook_template(path: Path, reference_path: Path | None = None) -> None:
    """Create the formula-driven three-sheet annual spending workbook."""
    source_path = reference_path or DEFAULT_REFERENCE
    if path.resolve() == source_path.resolve():
        raise WorkbookValidationError(
            "O destino não pode ser a planilha de referência."
        )
    path.parent.mkdir(parents=True, exist_ok=True)
    if _is_current_reference(source_path):
        workbook = _load_workbook_for_refresh(source_path)
        try:
            _write_base_rows(workbook, [], {})
            workbook.save(path)
        finally:
            workbook.close()
        return
    workbook = Workbook()
    base = workbook.active
    base.title = "BASE VIAGENS"
    support = workbook.create_sheet("APOIO")
    summary = workbook.create_sheet("RESUMO GASTOS")

    for column, header in enumerate(LEGACY_BASE_HEADERS, start=1):
        base.cell(1, column, header)
    base.append([None] * len(LEGACY_BASE_HEADERS))

    for column, header in enumerate(SUPPORT_HEADERS, start=1):
        support.cell(1, column, header)
    for category in DEBIT_CATEGORIES:
        support.append([category.code, category.name, category.segment, None])
    ppghd_row = next(
        row
        for row in range(2, support.max_row + 1)
        if support.cell(row, 1).value == "PPGH/PPGDH"
    )
    support.cell(ppghd_row, 2).comment = Comment(
        "Verba compartilhada de Mestrado e Doutorado em História.",
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
    code_validation.add(f"P2:P{_BASE_LAST_ROW}")
    base.add_data_validation(code_validation)

    install_base_controls(base)
    base.cell(2, 15, _segment_formula(2, support.max_row))
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

    _add_table(base, BASE_TABLE_NAME, "A1:Q2")
    _add_table(support, SUPPORT_TABLE_NAME, f"A1:D{support.max_row}")
    if reference_path is not None or source_path.exists():
        from scdp_automation.xlsx_layout import install_layout
        from scdp_automation.xlsx_reference import read_reference

        install_layout(workbook, source_path, read_reference(source_path))
    else:
        build_summary_formulas(workbook)

    from scdp_automation.xlsx_description import migrate_description
    from scdp_automation.xlsx_presentation import format_input_sheets

    migrate_description(workbook)
    format_input_sheets(workbook)

    workbook.calculation.calcMode = "auto"
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.save(path)


def install_base_controls(base: Worksheet) -> None:
    """Cover future manual rows with choice lists and missing-decision highlighting."""

    headers = (
        BASE_HEADERS
        if base.cell(1, 15).value == "Descrição do pedido"
        else LEGACY_BASE_HEADERS
    )
    decision_column = "R" if headers == BASE_HEADERS else "Q"
    for column, header in enumerate(headers, 1):
        base.cell(1, column, header)
    validations = cast(list[DataValidation], base.data_validations.dataValidation)
    if not any(v.formula1 == '"Sim,Não"' for v in validations):
        validation = DataValidation(
            type="list",
            formula1='"Sim,Não"',
            allow_blank=True,
            showDropDown=False,
            showErrorMessage=True,
            errorStyle="stop",
            errorTitle="Decisão inválida",
            error="Selecione Sim ou Não.",
        )
        validation.add(f"{decision_column}2:{decision_column}{_BASE_LAST_ROW}")
        base.add_data_validation(validation)
    install_decision_highlighting(base)
    for column in (14, 15, 16, 17):
        base.column_dimensions[base.cell(1, column).column_letter].width = 26


def install_decision_highlighting(base: Worksheet) -> None:
    """Distinguish undecided cancellations from missing travel classifications."""
    from openpyxl.formatting.rule import FormulaRule

    modern = base.cell(1, 15).value == "Descrição do pedido"
    decision = "R" if modern else "Q"
    segment = "P" if modern else "O"
    last = "R" if modern else "Q"
    red_formula = f'AND($C2="Cancelada",${decision}2="",$A2<>"")'
    yellow_formula = f'AND($A2<>"",${segment}2="")'
    managed_formulas = {
        red_formula,
        yellow_formula,
        'AND($C2="Cancelada",$Q2="",$A2<>"")',
        'AND($A2<>"",$O2="")',
        'AND($Q2="",$A2<>"")',
        'AND($A2<>"",OR($P2="",$Q2=""))',
        'AND($R2="",$A2<>"")',
        'AND($A2<>"",OR($Q2="",$R2=""))',
    }
    for area in list(base.conditional_formatting):
        rules = base.conditional_formatting[area]
        rules[:] = [
            rule
            for rule in rules
            if not (
                rule.formula
                and len(rule.formula) == 1
                and rule.formula[0] in managed_formulas
            )
        ]
        if not rules:
            del base.conditional_formatting[str(area.sqref)]
        for rule in rules:
            if rule.priority < 3:
                rule.priority += 2
    for formula, color, priority in (
        (red_formula, "FFC7CE", 1),
        (yellow_formula, "FFF2CC", 2),
    ):
        rule = FormulaRule(
            formula=[formula], fill=PatternFill("solid", fgColor=color), stopIfTrue=True
        )
        rule.priority = priority
        base.conditional_formatting.add(f"A2:{last}{_BASE_LAST_ROW}", rule)


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
    from scdp_automation.xlsx_layout import LAYOUT_NAME, validate_layout

    if LAYOUT_NAME in workbook.defined_names:
        validate_layout(workbook)
        _check_common_structure(workbook)
        return
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
    if len(validations) != 1 or str(validations[0].sqref) != f"Q2:Q{_BASE_LAST_ROW}":
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

    summary_row = 2
    for support_row in range(2, support.max_row + 1):
        if not support.cell(support_row, 1).value:
            continue
        actual_formulas = tuple(
            summary.cell(summary_row, column).value
            for column in range(1, len(SUMMARY_HEADERS) + 1)
        )
        from scdp_automation.xlsx_description import shift_description_references

        if actual_formulas != tuple(
            shift_description_references(value, "RESUMO GASTOS")
            if isinstance(value, str) and value.startswith("=")
            else value
            for value in _category_summary_formulas(summary_row, support_row)
        ):
            raise WorkbookValidationError(
                "Uma fórmula da worksheet RESUMO GASTOS está ausente ou foi alterada."
            )
        summary_row += 1

    pending_row = len(DEBIT_CATEGORIES) + 2
    expected_pending_formulas = (
        "SEM CLASSIFICAÇÃO",
        None,
        "PCDPs sem código de débito",
        "Pendente",
        "=COUNTIFS('BASE VIAGENS'!$P:$P,\"\",'BASE VIAGENS'!$A:$A,\"<>\")",
        _unclassified_sumifs_formula("D"),
        _unclassified_sumifs_formula("E"),
        _unclassified_sumifs_formula("G"),
        _unclassified_sumifs_formula("H"),
        f"=G{pending_row}+H{pending_row}-I{pending_row}",
        _unclassified_sumifs_formula("F"),
        _unclassified_sumifs_formula("I"),
        f"=K{pending_row}+L{pending_row}",
        _unclassified_sumifs_formula("J"),
        _unclassified_sumifs_formula("K"),
        _unclassified_sumifs_formula("K", canceled_only=True),
        f"=O{pending_row}",
        "Pendente",
    )
    expected_pending_formulas = tuple(
        shift_description_references(value, "RESUMO GASTOS")
        if isinstance(value, str) and value.startswith("=")
        else value
        for value in expected_pending_formulas
    )
    actual_pending_formulas = tuple(
        summary.cell(pending_row, column).value
        for column in range(1, len(SUMMARY_HEADERS) + 1)
    )
    if actual_pending_formulas != expected_pending_formulas:
        raise WorkbookValidationError(
            "A linha de PCDPs sem classificação contém fórmula ausente ou alterada."
        )
    if (
        workbook.calculation.calcMode != "auto"
        or not workbook.calculation.fullCalcOnLoad
        or not workbook.calculation.forceFullCalc
    ):
        raise WorkbookValidationError(
            "O workbook não solicita recálculo completo no Excel."
        )


def _check_common_structure(workbook: Workbook) -> None:
    from scdp_automation.xlsx_adjustments import dynamic_range
    from scdp_automation.xlsx_layout import (
        LAYOUT_NAME,
        same_formula,
    )

    base = workbook["BASE VIAGENS"]
    support = workbook["APOIO"]
    if workbook.defined_names[LAYOUT_NAME].attr_text != '"9"':
        raise WorkbookValidationError("Versão do layout inválida.")
    name = workbook.defined_names.get(CODE_LIST_NAME)
    if name is None or not same_formula(
        "=" + name.attr_text, "=" + dynamic_range("APOIO", "A")
    ):
        raise WorkbookValidationError("Lista de códigos de débito ausente.")
    validations = base.data_validations.dataValidation
    if not any(
        v.type == "list"
        and v.formula1 in (f"={CODE_LIST_NAME}", CODE_LIST_NAME)
        and str(v.sqref) == f"Q2:Q{_BASE_LAST_ROW}"
        for v in validations
    ):
        raise WorkbookValidationError("Validação de códigos de débito ausente.")
    if BASE_TABLE_NAME not in base.tables or SUPPORT_TABLE_NAME not in support.tables:
        raise WorkbookValidationError("Tabela necessária ausente.")

    calculation = workbook.calculation
    if (
        calculation is None
        or calculation.calcMode != "auto"
        or not calculation.fullCalcOnLoad
        or not calculation.forceFullCalc
    ):
        raise WorkbookValidationError("Recálculo completo ausente.")


def _summaries(trips: Sequence[Viagem] | Sequence[TripSummary]) -> list[TripSummary]:
    result = []
    seen: set[str] = set()
    for trip in trips:
        summary = summarize_trips([trip])[0] if isinstance(trip, Viagem) else trip
        if summary.pcdp in seen:
            raise WorkbookValidationError("PCDP duplicada na listagem.")
        seen.add(summary.pcdp)
        result.append(summary)
    return result


def _load_workbook_for_refresh(
    path: Path, reference_path: Path | None = None
) -> Workbook:
    try:
        workbook = load_workbook(path, data_only=False)
    except (OSError, InvalidFileException, BadZipFile, KeyError, ValueError) as error:
        raise WorkbookValidationError(
            f"Não foi possível abrir o workbook existente em {path}."
        ) from error
    try:
        from scdp_automation.xlsx_adjustments import (
            install_names,
            merge_history,
            migrate_adjustments,
            reorder_base_columns,
        )
        from scdp_automation.xlsx_layout import (
            LAYOUT_NAME,
            legacy_validate_layout,
            update_display_names,
        )

        if (
            LAYOUT_NAME in workbook.defined_names
            and workbook.defined_names[LAYOUT_NAME].attr_text == '"3"'
        ):
            update_display_names(workbook)
            legacy_validate_layout(workbook)
            workbook.defined_names[LAYOUT_NAME].attr_text = '"5"'
        if (
            LAYOUT_NAME in workbook.defined_names
            and workbook.defined_names[LAYOUT_NAME].attr_text == '"5"'
        ):
            from scdp_automation.xlsx_reference import read_reference

            source_path = reference_path or DEFAULT_REFERENCE
            data = read_reference(source_path) if source_path.exists() else None
            migrate_adjustments(workbook, data)
        if LAYOUT_NAME in workbook.defined_names and workbook.defined_names[
            LAYOUT_NAME
        ].attr_text in ('"6"', '"7"'):
            reorder_base_columns(workbook)
            merge_history(workbook)
            install_names(workbook)
            workbook.defined_names[LAYOUT_NAME].attr_text = '"8"'
        if (
            LAYOUT_NAME in workbook.defined_names
            and workbook.defined_names[LAYOUT_NAME].attr_text == '"8"'
        ):
            from scdp_automation.xlsx_code_summary import migrate_code_summary

            migrate_code_summary(workbook)
        if (
            LAYOUT_NAME not in workbook.defined_names
            and "BASE VIAGENS" in workbook.sheetnames
            and tuple(c.value for c in workbook["BASE VIAGENS"][1])
            == (*BASE_HEADERS[:11], "Segmento", "Código de débito")
        ):
            reorder_base_columns(workbook)
        if LAYOUT_NAME in workbook.defined_names and workbook.calculation is not None:
            if workbook.calculation.calcMode is None:
                workbook.calculation.calcMode = "auto"
            workbook.calculation.fullCalcOnLoad = True
            workbook.calculation.forceFullCalc = True
            base = workbook["BASE VIAGENS"]
            for validation in base.data_validations.dataValidation:
                matches_code_list = (
                    validation.type == "list"
                    and validation.formula1
                    in (
                        CODE_LIST_NAME,
                        f"={CODE_LIST_NAME}",
                    )
                )
                covers_base = any(
                    area.min_col == area.max_col == 16
                    and area.min_row == 2
                    and area.max_row >= base.max_row
                    for area in validation.sqref.ranges
                )
                if matches_code_list and covers_base:
                    validation.formula1 = f"={CODE_LIST_NAME}"
                    validation.sqref = f"P2:P{_BASE_LAST_ROW}"
        if "BASE VIAGENS" in workbook.sheetnames:
            from scdp_automation.xlsx_description import migrate_description

            migrate_description(workbook)
            install_decision_highlighting(workbook["BASE VIAGENS"])
        _check_workbook_structure(workbook)
    except Exception:
        workbook.close()
        raise
    return workbook


def _write_base_rows(
    workbook: Workbook,
    summaries: Sequence[TripSummary],
    manual_codes: dict[str, str],
    manual_decisions: dict[str, str | None] | None = None,
) -> None:
    base = workbook["BASE VIAGENS"]
    support_last_row = workbook["APOIO"].max_row
    from scdp_automation.xlsx_layout import (
        LAYOUT_NAME,
        range_segment_formula,
        support_limit,
    )

    ranged = LAYOUT_NAME in workbook.defined_names
    previous_dates = {
        base.cell(r, 1).value: tuple(base.cell(r, c).value for c in (12, 13, 14))
        for r in range(2, base.max_row + 1)
    }
    previous_descriptions = {
        base.cell(r, 1).value: base.cell(r, 15).value
        for r in range(2, base.max_row + 1)
    }
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
            base.cell(row, column).value = value
        base.cell(
            row,
            16,
            range_segment_formula(row, support_limit(workbook)).replace("$P", "$Q")
            if ranged
            else _segment_formula(row, support_last_row).replace("$P", "$Q"),
        )
        base.cell(row, 15).value = (
            summary.description or previous_descriptions.get(pcdp)
            if index < len(summaries)
            else None
        )
        base.cell(row, 17).value = manual_codes.get(pcdp)
        base.cell(row, 18).value = (manual_decisions or {}).get(pcdp)
        dates = (
            (summary.start_date, summary.end_date, summary.verified_date)
            if index < len(summaries)
            else (None, None, None)
        )
        old_dates = previous_dates.get(pcdp, (None, None, None))
        dates = tuple(new or old for new, old in zip(dates, old_dates, strict=True))
        for column, value in enumerate(dates, 12):
            base.cell(row, column).value = value
            base.cell(row, column).number_format = "dd/mm/yyyy"
        base.cell(row, 4).number_format = "0.0"
        for column in range(5, 12):
            base.cell(row, column).number_format = '"R$" #,##0.00;[Red]-"R$" #,##0.00'

    base.tables[BASE_TABLE_NAME].ref = f"A1:R{target_last_row}"
    base.tables[BASE_TABLE_NAME].tableColumns = []
    from scdp_automation.xlsx_presentation import format_base_sheet

    format_base_sheet(base)


def build_candidate(
    trips: Sequence[Viagem] | Sequence[TripSummary],
    current_path: Path,
    candidate_path: Path,
    *,
    reference_path: Path | None = None,
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

    source_path = reference_path or DEFAULT_REFERENCE
    summaries = _summaries(trips)
    incoming_pcdps = {summary.pcdp for summary in summaries}
    manual_codes: dict[str, str] = {}
    manual_decisions: dict[str, str | None] = {}
    workbook: Workbook

    if current_path.exists():
        workbook = _load_workbook_for_refresh(current_path, reference_path)
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
            decision = base.cell(row, 18).value
            if decision not in (None, "", "Sim", "Não"):
                workbook.close()
                raise WorkbookValidationError("Decisão de desconto inválida na base.")
            manual_decisions[pcdp] = decision or None
            code = base.cell(row, 17).value
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
        create_workbook_template(candidate_path, reference_path)
        workbook = _load_workbook_for_refresh(candidate_path)

    try:
        from scdp_automation.xlsx_layout import LAYOUT_NAME, install_layout
        from scdp_automation.xlsx_reference import read_reference

        source_path = reference_path or DEFAULT_REFERENCE
        if (
            LAYOUT_NAME not in workbook.defined_names
            and source_path.exists()
            and not _is_current_reference(source_path)
        ):
            old_allocations = {
                workbook["APOIO"].cell(r, 1).value: workbook["APOIO"].cell(r, 4).value
                for r in range(2, workbook["APOIO"].max_row + 1)
            }
            data = read_reference(source_path)
            install_layout(workbook, source_path, data)
            for row in range(2, workbook["APOIO"].max_row + 1):
                value = old_allocations.get(workbook["APOIO"].cell(row, 1).value)
                if value is not None:
                    workbook["APOIO"].cell(row, 4, value)
            manual_codes = data.codes | manual_codes
            manual_decisions = data.decisions | manual_decisions
        elif (
            not current_path.exists()
            and source_path.exists()
            and not _is_current_reference(source_path)
        ):
            data = read_reference(source_path)
            manual_codes = data.codes
            manual_decisions = dict(data.decisions)
        candidate_path.parent.mkdir(parents=True, exist_ok=True)
        _write_base_rows(workbook, summaries, manual_codes, manual_decisions)
        workbook.save(candidate_path)
    except Exception:
        workbook.close()
        candidate_path.unlink(missing_ok=True)
        raise
    workbook.close()


def _validate_candidate(
    candidate_path: Path, trips: Sequence[Viagem] | Sequence[TripSummary]
) -> None:
    summaries = _summaries(trips)
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
            code = base.cell(row, 17).value
            if code not in (None, "") and code not in support_codes:
                raise WorkbookValidationError(
                    "O candidato contém código de débito desconhecido."
                )
            segment_formula = base.cell(row, 16).value
            from scdp_automation.xlsx_layout import (
                LAYOUT_NAME,
                range_segment_formula,
                same_formula,
                support_limit,
            )

            expected_segment = (
                range_segment_formula(row, support_limit(workbook)).replace("$P", "$Q")
                if LAYOUT_NAME in workbook.defined_names
                else _segment_formula(row, support.max_row).replace("$P", "$Q")
            )
            if not same_formula(segment_formula, expected_segment):
                raise WorkbookValidationError(
                    "A fórmula Segmento está ausente ou foi alterada na base."
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
    trips: Sequence[Viagem] | Sequence[TripSummary],
    workbook_path: Path = DEFAULT_WORKBOOK,
    *,
    reference_path: Path | None = None,
    recalculate: bool = False,
) -> Path | None:
    """Serialize refreshes, validate a neighboring candidate, then replace atomically."""
    workbook_path.parent.mkdir(parents=True, exist_ok=True)
    lock = FileLock(f"{workbook_path}.lock", timeout=FILELOCK_TIMEOUT)
    try:
        with lock:
            candidate = _candidate_path_for(workbook_path)
            candidate_retained = False
            try:
                if reference_path is None:
                    build_candidate(trips, workbook_path, candidate)
                else:
                    build_candidate(
                        trips, workbook_path, candidate, reference_path=reference_path
                    )
                _validate_candidate(candidate, trips)
                if recalculate:
                    from scdp_automation.xlsx_recalculate import recalculate_workbook

                    recalculate_workbook(candidate)
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


def import_reference_workbook(
    reference: Path, output: Path, *, recalculate: bool = False
) -> Path | None:
    """Bootstrap the simplified workbook from the original, without extraction."""
    from scdp_automation.xlsx_reference import read_reference

    if reference.resolve() == output.resolve():
        raise WorkbookValidationError(
            "O destino não pode ser a planilha de referência."
        )
    data = read_reference(reference)
    return publish_workbook(
        data.trips, output, reference_path=reference, recalculate=recalculate
    )
