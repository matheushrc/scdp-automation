"""Domain mapping and debit categories for the annual spending workbook."""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
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
