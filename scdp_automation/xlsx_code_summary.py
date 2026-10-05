"""Link summary rows to debit codes without a manually maintained group."""

from __future__ import annotations

import re
from copy import copy

from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils.cell import range_boundaries
from openpyxl.workbook.workbook import Workbook

from scdp_automation.xlsx_adjustments import (
    charged_for_codes,
    group_formulas,
    install_names,
    pending_for_codes,
)

SUMMARY_ALIASES = {
    "OUTROS": "DIREÇÃO",
    "RESOLUÇÃO 049/2022-CONSUNI/CPPGEC": "CAPPG - Res 49",
}


def summary_code(label: str) -> str:
    fallback = '""'
    for name, code in SUMMARY_ALIASES.items():
        fallback = f'IF({label}="{name}","{code}",{fallback})'
    return (
        f"IFERROR(INDEX(ApoioA,MATCH({label},ApoioB,0)),"
        f"IF(COUNTIF(ApoioA,{label})>0,{label},{fallback}))"
    )


def selection_for_row(row: int) -> str:
    code = summary_code(f"$B{row}")
    return (
        f'--(ApoioA<>"")*--(((ApoioA=({code}))'
        f'+(({code})="DIREÇÃO")*(LEFT(ApoioA,9)="DIREÇÃO -"))>0)'
    )


def code_group_formulas(row: int) -> dict[str, str]:
    code = summary_code(f"$B{row}")
    selected = selection_for_row(row)
    split = (
        '--(ApoioA="PPGEL +")*('
        f'(({code})="PPGE")*ApoioJ+(({code})="PPGEL")*ApoioK'
        f'+(({code})="PPGH/PPGDH")*ApoioL)'
    )
    result = {}
    for output, source in (("E", "D"), ("I", "E"), ("J", "G")):
        amounts = f"Apoio{source}"
        count = f"SUMPRODUCT(({selected})+({split}),--ISNUMBER({amounts}))"
        summed = f"SUMPRODUCT(({selected})+({split}),{amounts})"
        result[f"{output}{row}"] = f'=IF({count}=0,"Pendente",{summed})'
    result[f"C{row}"] = (
        f'=IF(AND(ISNUMBER(E{row}),ISNUMBER(I{row})),E{row}+I{row},"Pendente")'
    )
    pending = f"SUMPRODUCT(({selected})+({split}),{pending_for_codes('ApoioA')})"
    charged = f"SUMPRODUCT(({selected})+({split}),({charged_for_codes('ApoioA')}))"
    result[f"F{row}"] = f'=IF({pending}>0,"Pendente",{charged})'
    for output, first, second in (("G", "E", "F"), ("K", "I", "J")):
        result[f"{output}{row}"] = (
            f"=IF(AND(ISNUMBER({first}{row}),ISNUMBER({second}{row})),"
            f'{first}{row}-{second}{row},"Pendente")'
        )
    result[f"M{row}"] = (
        f'=IF(AND(ISNUMBER(G{row}),ISNUMBER(K{row})),G{row}+K{row},"Pendente")'
    )
    return result


def total_for_rows(column: str, *, principal: bool = False) -> str:
    amounts = f"Resumo{column}"
    selected = '--(ApoioA<>"")*--(ApoioA<>"PPGEL +")*--(LEFT(ApoioA,9)<>"DIREÇÃO -")'
    if principal:
        selected += '*--(ApoioC<>"SEG 5 AFAST PAÍS")'
    code_only = "--(ApoioA<>ApoioB)"
    summed = (
        f'SUMPRODUCT({selected},SUMIFS({amounts},ResumoB,ApoioB,ResumoE,"<>"))'
        f'+SUMPRODUCT({selected},{code_only},SUMIFS({amounts},ResumoB,ApoioA,ResumoE,"<>"))'
    )
    pending = (
        f'SUMPRODUCT({selected},COUNTIFS(ResumoB,ApoioB,ResumoE,"<>",{amounts},"Pendente"))'
        f'+SUMPRODUCT({selected},{code_only},COUNTIFS(ResumoB,ApoioA,ResumoE,"<>",{amounts},"Pendente"))'
    )
    for label in SUMMARY_ALIASES:
        absent = f'COUNTIF(ApoioA,"{label}")+COUNTIF(ApoioB,"{label}")=0'
        summed += f'+IF({absent},SUMIF(ResumoB,"{label}",{amounts}),0)'
        pending += f'+IF({absent},COUNTIFS(ResumoB,"{label}",{amounts},"Pendente"),0)'
    return f'IF({pending}>0,"Pendente",{summed})'


def translate_support_references(formula: str, sheet: str) -> str:
    tokens = Tokenizer(formula).items
    for token in tokens:
        if token.type != "OPERAND" or token.subtype != "RANGE":
            continue
        prefix, separator, address = token.value.rpartition("!")
        if separator and prefix.strip("'") != "APOIO":
            continue
        if not separator:
            if sheet != "APOIO":
                continue
            address = token.value
        if not re.fullmatch(
            r"\$?[A-Z]{1,3}(?:\$?\d+)?(?::\$?[A-Z]{1,3}(?:\$?\d+)?)?", address
        ):
            continue

        def move(match: re.Match[str]) -> str:
            column = match[0]
            return {"J": "I", "K": "J", "L": "K", "M": "L"}.get(column, column)

        token.value = prefix + separator + re.sub(r"[A-Z]+", move, address)
    return "=" + "".join(token.value for token in tokens)


def references_old_group(formula: str, sheet: str) -> bool:
    for token in Tokenizer(formula).items:
        if token.type != "OPERAND" or token.subtype != "RANGE":
            continue
        if token.value.lower() == "apoioi":
            return True
        prefix, separator, address = token.value.rpartition("!")
        if separator and prefix.strip("'").upper() != "APOIO":
            continue
        if not separator:
            if sheet != "APOIO":
                continue
            address = token.value
        if re.fullmatch(
            r"\$?[A-Z]{1,3}(?:\$?\d+)?(?::\$?[A-Z]{1,3}(?:\$?\d+)?)?", address
        ):
            first, _, last, _ = range_boundaries(address)
            if first is not None and last is not None and first <= 9 <= last:
                return True
    return False


def migrate_code_summary(workbook: Workbook) -> None:
    """Remove the old group column once, preserving manual rows and formulas."""
    from scdp_automation.xlsx_layout import LAYOUT_NAME
    from scdp_automation.xlsx_presentation import format_input_sheets

    support, summary = workbook["APOIO"], workbook["RESUMO GASTOS"]
    if support["I1"].value != "Grupo no resumo":
        return
    old_groups = {
        support.cell(r, 9).value: support.cell(r, 1).value
        for r in range(2, support.max_row + 1)
        if support.cell(r, 9).value and support.cell(r, 1).value
    }
    financial_rows = [
        r
        for r in range(2, summary.max_row + 1)
        if any(
            isinstance(summary.cell(r, c).value, str)
            and re.search(rf"ApoioI(?:=|,)\$?B\$?{r}(?!\d)", summary.cell(r, c).value)
            for c in (5, 6, 9, 10)
        )
    ]
    from scdp_automation.xlsx_layout import same_formula

    generated_cells = {
        cell.coordinate
        for r in financial_rows
        for cell in summary[r]
        if cell.data_type == "f"
        and "ApoioI" in cell.value
        and any(
            same_formula(
                cell.value,
                group_formulas(r, weight_column=weight).get(cell.coordinate, "=0"),
            )
            for weight in (None, "J", "K", "L")
        )
    }
    custom_group_cells = {
        cell.coordinate: cell.value
        for r in financial_rows
        for cell in summary[r]
        if cell.data_type == "f"
        and "ApoioI" in cell.value
        and cell.coordinate not in generated_cells
    }
    old_diagnostic = (
        '=IF(SUMPRODUCT(--(ApoioA<>""),--(ApoioA<>"PPGEL +"),'
        '--ISNA(MATCH(ApoioI,ResumoP,0)))=0,0,"Pendente")'
    )
    for sheet in workbook:
        for cells in sheet:
            for cell in cells:
                if cell.data_type != "f" or not references_old_group(
                    cell.value, sheet.title
                ):
                    continue
                supported = sheet.title == summary.title and (
                    cell.coordinate in generated_cells
                    or cell.coordinate in custom_group_cells
                    or (
                        summary.cell(cell.row, 2).value
                        == "Categorias sem grupo reconhecido"
                        and same_formula(cell.value, old_diagnostic)
                    )
                )
                if not supported:
                    from scdp_automation.xlsx_output import WorkbookValidationError

                    raise WorkbookValidationError(
                        f"A fórmula manual {sheet.title}!{cell.coordinate} usa o grupo antigo e precisa de adaptação."
                    )
    for name, definition in workbook.defined_names.items():
        if name != "ApoioI" and references_old_group("=" + definition.attr_text, ""):
            from scdp_automation.xlsx_output import WorkbookValidationError

            raise WorkbookValidationError(
                f"O intervalo {name} usa o grupo antigo e precisa de adaptação."
            )
    # Retain an existing custom display label by moving it to the category name.
    names = {support.cell(r, 2).value for r in range(2, support.max_row + 1)}
    codes = {support.cell(r, 1).value for r in range(2, support.max_row + 1)}
    for row in financial_rows:
        label = summary.cell(row, 2).value
        if label not in names and label not in codes and label not in SUMMARY_ALIASES:
            code = old_groups.get(label)
            for r in range(2, support.max_row + 1):
                if code and support.cell(r, 1).value == code:
                    support.cell(r, 2).value = label
    for sheet in workbook:
        for row in sheet:
            for cell in row:
                if cell.data_type == "f":
                    cell.value = translate_support_references(cell.value, sheet.title)
    for name in workbook.defined_names.values():
        name.attr_text = translate_support_references("=" + name.attr_text, "")[1:]
    dimensions = {col: copy(support.column_dimensions[col]) for col in "JKLM"}
    support.delete_cols(9)
    for old, new in zip("JKLM", "IJKL", strict=True):
        dimension = dimensions[old]
        dimension.index = new
        dimension.min = dimension.max = ord(new) - 64
        support.column_dimensions[new] = dimension
    support.column_dimensions.pop("M", None)
    table = support.tables["tblApoioDebito"]
    table.ref = f"A1:L{support.max_row}"
    table.tableColumns = []
    if table.autoFilter is not None:
        table.autoFilter.ref = table.ref
    workbook.defined_names[LAYOUT_NAME].attr_text = '"9"'
    install_names(workbook)
    for row in financial_rows:
        for coordinate, formula in code_group_formulas(row).items():
            if coordinate in generated_cells:
                summary[coordinate] = formula
    for coordinate, formula in custom_group_cells.items():
        row = summary[coordinate].row
        selection = selection_for_row(row)
        formula = re.sub(
            rf"SUMIF\(ApoioI,\$?B\$?{row},([^()]+)\)",
            lambda match, selection=selection: f"SUMPRODUCT({selection},{match[1]})",
            formula,
            flags=re.IGNORECASE,
        )
        formula = re.sub(
            rf"--\(ApoioI=\$?B\$?{row}\)",
            selection_for_row(row),
            formula,
            flags=re.IGNORECASE,
        )
        if "ApoioI" in formula:
            from scdp_automation.xlsx_output import WorkbookValidationError

            raise WorkbookValidationError(
                f"A fórmula manual {coordinate} usa o grupo antigo e precisa de adaptação."
            )
        summary[coordinate] = translate_support_references(formula, summary.title)
    by_label = {
        summary.cell(r, 2).value: r
        for r in range(2, summary.max_row + 1)
        if isinstance(summary.cell(r, 2).value, str)
    }
    total_row = by_label["TOTAL CONSOLIDADO"]
    unknown_row = by_label["Categorias sem grupo reconhecido"]
    summary.cell(unknown_row, 2).value = "Categorias sem linha no resumo"
    unclassified_row = by_label["SEM CLASSIFICAÇÃO"]
    principal_row = by_label["TOTAL (Diárias, Passagens, Transportes)"]
    # Recognition uses visible names/codes, so copied
    # formulas work in newly inserted rows without hidden metadata.
    visible = "ResumoB"
    unknown = (
        f'SUMPRODUCT(--(ApoioA<>""),--(ApoioA<>"PPGEL +"),'
        '--(LEFT(ApoioA,9)<>"DIREÇÃO -"),'
        f"--ISNA(MATCH(ApoioA,{visible},0)),--ISNA(MATCH(ApoioB,{visible},0)),"
        '--(ApoioA<>"DIREÇÃO"),--(ApoioA<>"CAPPG - Res 49"))'
    )
    for column in ("C", "E", "F", "G", "I", "J", "K", "M"):
        summary[f"{column}{principal_row}"] = "=" + total_for_rows(
            column, principal=True
        )
        if column != "F":
            summary[f"{column}{unknown_row}"] = f'=IF({unknown}=0,0,"Pendente")'
            subtotal = total_for_rows(column)
            summary[f"{column}{total_row}"] = (
                f"=IF(AND(ISNUMBER({subtotal}),ISNUMBER({column}{unknown_row}),"
                f"ISNUMBER({column}{unclassified_row})),"
                f'{subtotal}+{column}{unknown_row}+{column}{unclassified_row},"Pendente")'
            )
    subtotal = total_for_rows("F")
    summary[f"F{unknown_row}"] = (
        f"=IF(AND(ISNUMBER(F{total_row}),ISNUMBER({subtotal}),"
        f'ISNUMBER(F{unclassified_row})),F{total_row}-({subtotal})-F{unclassified_row},"Pendente")'
    )
    # Old metadata stays available to custom formulas but no longer controls totals.
    format_input_sheets(workbook)
