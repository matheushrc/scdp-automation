"""Small public workbook fixtures with the same source coordinates."""

from collections.abc import Mapping
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from scdp_automation.relatorio import Viagem
from scdp_automation.xlsx_models import summarize_trips
from scdp_automation.xlsx_validation import BASE_HEADERS


def final_template_fixture(path: Path) -> None:
    """Build the final editable layout independently of production writers."""
    from datetime import date

    from openpyxl.workbook.defined_name import DefinedName
    from openpyxl.workbook.properties import CalcProperties
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.worksheet.table import Table

    workbook = Workbook()
    base = workbook.active
    base.title = "BASE VIAGENS"
    support = workbook.create_sheet("APOIO")
    summary = workbook.create_sheet("RESUMO GASTOS")
    base.append(BASE_HEADERS)
    base.append([None] * 18)
    base["P2"] = '=IF($Q2="","",IFERROR(VLOOKUP($Q2,ApoioCatalogo,3,FALSE),""))'
    support.append(
        (
            "Código de débito",
            "Nome por extenso",
            "Segmento",
            "Diárias e passagens distribuído (R$)",
            "Transportes distribuído (R$)",
            "Recurso total (R$)",
            "Transportes agendado (R$)",
            "Transportes pago (R$)",
            "Rateio para PPGE (%)",
            "Rateio para PPGEL (%)",
            "Rateio para PPGH (%)",
            "Total utilizado por categoria (R$)",
        )
    )
    categories = (
        ("AGRONOMIA", "Agronomia", "SEG 1 GRADUAÇÃO"),
        ("PPGE", "Mestrado em Educação", "SEG 2 MESTRADO"),
        ("PPGEL", "Mestrado em Estudos Linguísticos", "SEG 2 MESTRADO"),
        ("HISTÓRIA", "História", "SEG 1 GRADUAÇÃO"),
        ("PPGH/PPGDH", "Mestrado e Doutorado em História", "SEG 2 MESTRADO"),
        ("PPGEL +", "Rateio entre programas", "SEG 2 MESTRADO"),
        ("DIREÇÃO", "Direção", "SEG 3 OUTROS"),
        ("DIREÇÃO - AGAS", "AGAS", "SEG 3 OUTROS"),
        ("AFASTAMENTO", "Afastamento", "SEG 3 OUTROS"),
    )
    for row, category in enumerate(categories, 2):
        support.append((*category, 400, 600, None, 200, 150, None, None, None, None))
        support.cell(
            row, 6, f'=IF(AND(ISNUMBER(D{row}),ISNUMBER(E{row})),D{row}+E{row},"")'
        )
        support.cell(
            row,
            12,
            f'=IF(COUNTIFS(ViagensM,A{row},ViagensC,"Cancelada",ViagensN,"")>0,"Pendente",SUMIFS(ViagensK,ViagensM,A{row},ViagensC,"<>Cancelada")+SUMIFS(ViagensK,ViagensM,A{row},ViagensC,"Cancelada",ViagensN,"Sim"))',
        )
        if category[0] == "PPGEL +":
            for column, weight in zip((9, 10, 11), (0.5, 0.25, 0.25), strict=True):
                support.cell(row, column, weight)
    summary["B2"] = "RESUMO DE GASTOS"
    summary.merge_cells("B2:I2")
    summary["M2"] = date(2026, 1, 1)
    summary["M2"].number_format = "dd/mm/yyyy"
    # Nine representative final rows exercise financial rules without rebuilding
    # the removed legacy catalogue or presentation generator.
    for row, code in (
        (7, "AGRONOMIA"),
        (8, "PPGE"),
        (9, "PPGEL"),
        (10, "HISTÓRIA"),
        (11, "PPGH/PPGDH"),
        (12, "DIREÇÃO"),
        (13, "AFASTAMENTO"),
    ):
        summary.cell(row, 2, code)
        for output, amounts in (("E", "D"), ("I", "E"), ("J", "G")):
            formula = f"SUMIF(ApoioA,$B{row},Apoio{amounts})"
            if code == "DIREÇÃO":
                formula += f'+SUMIF(ApoioA,"DIREÇÃO -*",Apoio{amounts})'
            weight = {"PPGE": "J", "PPGEL": "K", "PPGH/PPGDH": "L"}.get(code)
            if weight:
                formula += (
                    f'+SUMPRODUCT(--(ApoioA="PPGEL +"),Apoio{weight},Apoio{amounts})'
                )
            summary[f"{output}{row}"] = "=" + formula
        criterion = '"DIREÇÃO*"' if code == "DIREÇÃO" else f"$B{row}"
        charged = f'SUMIFS(ViagensK,ViagensM,{criterion},ViagensC,"<>Cancelada")+SUMIFS(ViagensK,ViagensM,{criterion},ViagensC,"Cancelada",ViagensN,"Sim")'
        pending = f'COUNTIFS(ViagensM,{criterion},ViagensC,"Cancelada",ViagensN,"")'
        if weight:
            split_charged = 'SUMIFS(ViagensK,ViagensM,"PPGEL +",ViagensC,"<>Cancelada")+SUMIFS(ViagensK,ViagensM,"PPGEL +",ViagensC,"Cancelada",ViagensN,"Sim")'
            charged += f'+SUMIF(ApoioA,"PPGEL +",Apoio{weight})*({split_charged})'
            pending += f'+SUMIF(ApoioA,"PPGEL +",Apoio{weight})*COUNTIFS(ViagensM,"PPGEL +",ViagensC,"Cancelada",ViagensN,"")'
        summary[f"F{row}"] = f'=IF({pending}>0,"Pendente",{charged})'
        summary[f"C{row}"] = f"=E{row}+I{row}"
        for output, first, second in (("G", "E", "F"), ("K", "I", "J")):
            summary[f"{output}{row}"] = (
                f'=IF(AND(ISNUMBER({first}{row}),ISNUMBER({second}{row})),{first}{row}-{second}{row},"Pendente")'
            )
        summary[f"M{row}"] = (
            f'=IF(AND(ISNUMBER(G{row}),ISNUMBER(K{row})),G{row}+K{row},"Pendente")'
        )
    summary["B20"] = "TOTAL"
    summary["C20"] = "=SUM(ApoioD)+SUM(ApoioE)"
    summary["F20"] = (
        '=IF(COUNTIFS(ViagensC,"Cancelada",ViagensN,"",ViagensA,"<>")>0,"Pendente",SUMIFS(ViagensK,ViagensC,"<>Cancelada")+SUMIFS(ViagensK,ViagensC,"Cancelada",ViagensN,"Sim"))'
    )
    summary["J20"] = "=SUM(ApoioG)"
    for sheet, prefix, columns, height_row in (
        ("APOIO", "Apoio", "ABCDEFGHJKLM", 1),
        ("BASE VIAGENS", "Viagens", "ABCDEFGHIJKLMNOPQ", 2),
        ("RESUMO GASTOS", "Resumo", "BCDEFGHIJKLMNOP", 3),
    ):
        key = "B" if sheet == "RESUMO GASTOS" else "A"
        summary.cell(
            height_row,
            18,
            f"""=MAX(1,IFERROR(LOOKUP(2,1/('{sheet}'!${key}:${key}<>""),ROW('{sheet}'!${key}:${key}))-1,1))""",
        )
        for column in columns:
            actual_column = (
                {"L": "P", "M": "Q", "N": "R", "O": "L", "P": "M", "Q": "N"}.get(
                    column, column
                )
                if sheet == "BASE VIAGENS"
                else {"J": "I", "K": "J", "L": "K", "M": "L"}.get(column, column)
                if sheet == "APOIO"
                else column
            )
            workbook.defined_names.add(
                DefinedName(
                    prefix + column,
                    attr_text=f"OFFSET('{sheet}'!${actual_column}$1,1,0,MAX(1,'RESUMO GASTOS'!$R${height_row}),1)",
                )
            )
    for name, width in (("CodigosDebito", 1), ("ApoioCatalogo", 3)):
        workbook.defined_names.add(
            DefinedName(
                name,
                attr_text=f"OFFSET('APOIO'!$A$1,1,0,MAX(1,'RESUMO GASTOS'!$R$1),{width})",
            )
        )
    workbook.defined_names.add(DefinedName("SCDPLayoutVersion", attr_text='"9"'))
    for sheet, name, area in (
        (base, "tblBaseViagens", "A1:R2"),
        (support, "tblApoioDebito", "A1:L10"),
    ):
        sheet.add_table(Table(displayName=name, ref=area))
    for formula, area in (
        ("CodigosDebito", "Q2:Q1048576"),
        ('"Sim,Não"', "R2:R1048576"),
    ):
        validation = DataValidation(type="list", formula1=formula, allow_blank=True)
        validation.add(area)
        base.add_data_validation(validation)
    validation = DataValidation(
        type="decimal", operator="between", formula1=0, formula2=1, allow_blank=True
    )
    validation.add("I2:K10")
    support.add_data_validation(validation)
    for sheet in workbook:
        sheet.column_dimensions["B"].width = 44
        sheet.row_dimensions[2].height = 28
        for row in sheet:
            for cell in row:
                if cell.__class__.__name__ != "MergedCell":
                    cell.font = Font(name="Arial", size=10, color="123456")
                    cell.fill = PatternFill("solid", fgColor="DDEEFF")
                    cell.alignment = Alignment(horizontal="center")
    workbook.calculation = CalcProperties(
        calcMode="auto", fullCalcOnLoad=True, forceFullCalc=True
    )
    workbook.save(path)
    workbook.close()


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


def write_existing_workbook(
    path: Path,
    trips: list[Viagem],
    codes: dict[str, str] | None = None,
    allocations: Mapping[str, object] | None = None,
) -> None:
    from openpyxl import load_workbook
    from openpyxl.formula.translate import Translator

    final_template_fixture(path)
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
            16,
            Translator(base["P2"].value, origin="P2").translate_formula(f"P{row}"),
        )
        base.cell(row, 17, codes.get(summary.pcdp))

    base.tables["tblBaseViagens"].ref = f"A1:R{max(2, len(trips) + 1)}"
    support = workbook["APOIO"]
    for row in range(2, support.max_row + 1):
        code = support.cell(row, 1).value
        if code in allocations:
            support.cell(row, 4, allocations[code])
    workbook.save(path)
    workbook.close()


def read_base_rows(path: Path) -> dict[str, tuple[object, ...]]:
    from openpyxl import load_workbook

    workbook = load_workbook(path, data_only=False)
    base = workbook["BASE VIAGENS"]
    rows = {
        base.cell(row, 1).value: tuple(
            base.cell(row, column).value for column in range(1, 19)
        )
        for row in range(2, base.max_row + 1)
        if base.cell(row, 1).value
    }
    workbook.close()
    return rows
