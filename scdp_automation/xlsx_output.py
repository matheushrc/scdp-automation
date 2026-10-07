"""Refresh BASE from complete trips, preserving the final template's manual sheets."""

from __future__ import annotations

import os
from collections.abc import Sequence
from copy import copy
from datetime import datetime
from pathlib import Path
from typing import cast
from uuid import uuid4
from zipfile import BadZipFile
from zoneinfo import ZoneInfo

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.worksheet import Worksheet

from scdp_automation.config import REPO_ROOT
from scdp_automation.relatorio import Viagem
from scdp_automation.xlsx_models import TripSummary, summarize_trips
from scdp_automation.xlsx_validation import (
    WorkbookValidationError,
    same_formula,
    segment_formula,
    validate_workbook,
)

DEFAULT_TEMPLATE = REPO_ROOT / "input" / "gastos_scdp_template.xlsx"
BASE_TABLE_NAME = "tblBaseViagens"


class WorkbookPublishError(RuntimeError):
    """A validated workbook could not be promoted to the output path."""


def install_decision_highlighting(base: Worksheet) -> None:
    """Distinguish undecided cancellations from missing travel classifications."""
    from openpyxl.formatting.rule import FormulaRule

    decision, segment, last = "R", "P", "R"
    red_formula = f'AND($C2="Cancelada",${decision}2="",$A2<>"")'
    yellow_formula = f'AND($A2<>"",${segment}2="")'
    managed_formulas = {red_formula, yellow_formula}
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
        base.conditional_formatting.add(f"A2:{last}1048576", rule)


def _load_workbook_for_refresh(
    path: Path, *, role: str = "workbook", template_path: Path | None = None
) -> Workbook:
    context = f"{role} em {path}"
    if role == "Template obrigatório":
        recovery = "Restaure um template final válido, versão 9, nesse caminho."
    elif role == "Workbook publicado":
        source = template_path or DEFAULT_TEMPLATE
        recovery = (
            f"O template de origem é {source}. Para começar novamente, primeiro "
            f"preserve/mova a pasta de saída {path.parent} inteira, incluindo JSON "
            "e históricos, para fora da saída ativa; depois execute novamente. "
            "git pull/reset não altera a saída gerada localmente."
        )
    else:
        recovery = ""
    try:
        workbook = load_workbook(path, data_only=False)
    except PermissionError as error:
        raise WorkbookValidationError(
            f"Não foi possível acessar {context}; feche o arquivo no Excel e "
            "verifique as permissões de acesso."
        ) from error
    except (OSError, InvalidFileException, BadZipFile, KeyError, ValueError) as error:
        raise WorkbookValidationError(
            f"Não foi possível abrir {context}. {recovery}"
        ) from error
    try:
        validate_workbook(workbook)
    except WorkbookValidationError as error:
        workbook.close()
        raise WorkbookValidationError(
            f"{context}: {error} {recovery}", missing_pcdps=error.missing_pcdps
        ) from error
    except Exception:
        workbook.close()
        raise
    return workbook


def validate_workbook_sources(
    current_path: Path, *, template_path: Path | None = None
) -> None:
    """Preflight required source workbooks without creating or modifying output."""
    source = template_path or DEFAULT_TEMPLATE
    if source.resolve() == current_path.resolve():
        raise WorkbookValidationError(
            "Template e workbook publicado devem ter caminhos separados."
        )
    template = _load_workbook_for_refresh(source, role="Template obrigatório")
    template.close()
    if current_path.exists():
        current = _load_workbook_for_refresh(
            current_path, role="Workbook publicado", template_path=source
        )
        current.close()


def _extend_base_choices(base: Worksheet) -> None:
    """Candidate controls intentionally cover future BASE rows too."""
    for formula, column in (("CodigosDebito", "Q"), ('"Sim,Não"', "R")):
        for validation in cast(
            list[DataValidation], base.data_validations.dataValidation
        ):
            if validation.type == "list" and validation.formula1 in (
                formula,
                "=" + formula,
            ):
                validation.add(f"{column}2:{column}1048576")
                break


def _write_base_rows(
    workbook: Workbook,
    summaries: Sequence[TripSummary],
    manual_codes: dict[str, str],
    manual_decisions: dict[str, str | None] | None = None,
) -> None:
    base = workbook["BASE VIAGENS"]
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
        base.cell(row, 16).value = segment_formula(row)
        base.cell(row, 15).value = (
            summary.description if index < len(summaries) else None
        )
        base.cell(row, 17).value = manual_codes.get(pcdp)
        base.cell(row, 18).value = (manual_decisions or {}).get(pcdp)
        dates = (
            (summary.start_date, summary.end_date, summary.verified_date)
            if index < len(summaries)
            else (None, None, None)
        )
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


def _validate_candidate(
    candidate_path: Path, summaries: Sequence[TripSummary] | None = None
) -> None:
    workbook = _load_workbook_for_refresh(candidate_path)
    try:
        base = workbook["BASE VIAGENS"]
        actual_rows = [
            row for row in range(2, base.max_row + 1) if base.cell(row, 1).value
        ]
        for row in actual_rows:
            if not same_formula(base.cell(row, 16).value, segment_formula(row)):
                raise WorkbookValidationError(
                    "A fórmula Segmento está ausente ou foi alterada na base."
                )
        if summaries is not None:
            if [base.cell(row, 1).value for row in actual_rows] != [
                summary.pcdp for summary in summaries
            ]:
                raise WorkbookValidationError(
                    "As PCDPs do candidato não correspondem à listagem completa."
                )
            for row, summary in zip(actual_rows, summaries, strict=True):
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
                if tuple(base.cell(row, col).value for col in range(1, 12)) != expected:
                    raise WorkbookValidationError(
                        "Os campos agregados do candidato estão inconsistentes."
                    )
    finally:
        workbook.close()


def build_candidate(
    trips: Sequence[Viagem],
    current_path: Path,
    candidate_path: Path,
    *,
    template_path: Path | None = None,
) -> None:
    """Validate the required final template, refresh a separate candidate, then validate it."""
    source_path = template_path or DEFAULT_TEMPLATE
    resolved = (source_path.resolve(), current_path.resolve(), candidate_path.resolve())
    if len(set(resolved)) != 3:
        raise WorkbookValidationError(
            "Template, candidato e workbook publicado devem ter caminhos separados."
        )
    if candidate_path.exists():
        raise WorkbookValidationError(
            f"O caminho de candidato já existe: {candidate_path}"
        )
    template = _load_workbook_for_refresh(source_path, role="Template obrigatório")
    workbook = template
    try:
        summaries = summarize_trips(trips)
        manual_codes: dict[str, str] = {}
        manual_decisions: dict[str, str | None] = {}
        if current_path.exists():
            workbook = _load_workbook_for_refresh(
                current_path, role="Workbook publicado", template_path=source_path
            )
            base = workbook["BASE VIAGENS"]
            existing_pcdps = {
                base.cell(row, 1).value
                for row in range(2, base.max_row + 1)
                if base.cell(row, 1).value
            }
            missing = tuple(
                sorted(existing_pcdps - {summary.pcdp for summary in summaries})
            )
            if missing:
                raise WorkbookValidationError(
                    f"O workbook publicado em {current_path}: a listagem completa está sem {len(missing)} PCDP(s) já publicadas; reconcilie o relatório antes de atualizar.",
                    missing_pcdps=missing,
                )
            for row in range(2, base.max_row + 1):
                pcdp = base.cell(row, 1).value
                if pcdp:
                    if base.cell(row, 17).value:
                        manual_codes[pcdp] = base.cell(row, 17).value
                    manual_decisions[pcdp] = base.cell(row, 18).value or None
        else:
            # Template travel data is never an extraction input, even for matching keys.
            base = workbook["BASE VIAGENS"]
            for row in base.iter_rows(min_row=2):
                for cell in row:
                    cell.value = None
        _write_base_rows(workbook, summaries, manual_codes, manual_decisions)
        _extend_base_choices(workbook["BASE VIAGENS"])
        install_decision_highlighting(workbook["BASE VIAGENS"])
        updated = workbook["RESUMO GASTOS"]["M2"]
        updated.value = datetime.now(ZoneInfo("America/Sao_Paulo")).date()
        updated.number_format = "dd/mm/yyyy"
        workbook.calculation.calcMode = "auto"
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        validate_workbook(workbook)
        candidate_path.parent.mkdir(parents=True, exist_ok=True)
        workbook.save(candidate_path)
        _validate_candidate(candidate_path, summaries)
    except Exception:
        candidate_path.unlink(missing_ok=True)
        raise
    finally:
        workbook.close()
        if workbook is not template:
            template.close()


def _candidate_path_for(workbook_path: Path) -> Path:
    return workbook_path.with_name(
        f"{workbook_path.stem}.{uuid4().hex}.candidate{workbook_path.suffix}"
    )


def publish_workbook(
    trips: Sequence[Viagem],
    workbook_path: Path,
    *,
    template_path: Path | None = None,
) -> None:
    """Validate and atomically publish; OutputHistory owns session serialization."""
    workbook_path.parent.mkdir(parents=True, exist_ok=True)
    candidate = _candidate_path_for(workbook_path)
    candidate_retained = False
    try:
        if template_path is None:
            build_candidate(trips, workbook_path, candidate)
        else:
            build_candidate(
                trips, workbook_path, candidate, template_path=template_path
            )
        _validate_candidate(candidate)
        try:
            os.replace(candidate, workbook_path)
        except OSError as error:
            recovery_directory = workbook_path.parent / "backup"
            recovery = recovery_directory / f"recovery.{uuid4().hex}.xlsx"
            try:
                recovery_directory.mkdir(parents=True, exist_ok=True)
                candidate.rename(recovery)
            except OSError as recovery_error:
                candidate_retained = True
                surviving_path = candidate if candidate.exists() else recovery
                raise WorkbookPublishError(
                    f"Não foi possível substituir o workbook em {workbook_path} nem mover a candidata "
                    f"para recuperação: {recovery_error}. O candidato validado "
                    f"permanece em {surviving_path}; feche o arquivo no Excel e "
                    "tente novamente."
                ) from error
            raise WorkbookPublishError(
                f"Não foi possível substituir o workbook em {workbook_path}; feche o arquivo no Excel e "
                f"tente novamente. O candidato validado permanece em {recovery}."
            ) from error
    finally:
        if not candidate_retained:
            candidate.unlink(missing_ok=True)
