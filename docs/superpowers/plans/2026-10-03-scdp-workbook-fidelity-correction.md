# SCDP workbook fidelity correction implementation plan

> **For agentic workers:** Use superpowers:executing-plans. Execute in the existing xlsx-report-output worktree.

**Goal:** Rebuild all three worksheets from the original inputs while retaining the original RESUMO GASTOS presentation.

**Architecture:** Separate domain records, reference reading and visual layout from publication. Read the reference without saving it; deduplicate equivalent PCDPs and import classifications, budgets and transport inputs. Copy the summary presentation and replace its formulas with references to the simplified input. Preserve manual inputs on subsequent refresh.

**Tech Stack:** Python, openpyxl, filelock, unittest, LibreOffice for disposable recalculation and PDF rendering.

**Spec:** docs/superpowers/specs/2026-10-03-scdp-workbook-fidelity-correction-design.md

## Global constraints

- Preserve the original reference and the currently opened generated workbook.
- Use exactly BASE VIAGENS, APOIO and RESUMO GASTOS.
- Keep debit code as the final base column; preserve it by full PCDP.
- Keep summary formulas, all original B2:M58 visual styles and transport columns.
- Use original data for validation, distinguishing stale pivots from input totals.
- Never commit travel data or generated workbooks.

## Review focus

- Equivalent duplicate requests with one blank classification must retain the populated classification.
- Conflicting populated classifications or financial values must reject import.
- Grouped debit categories must not duplicate either expense or budget.
- Manual budget, scheduled and paid transport inputs must survive refresh.
- New or unbudgeted categories must be visible and included in reconciliation.

### Task 1: Import original records and inputs

**Files:** scdp_automation/xlsx_models.py, scdp_automation/xlsx_reference.py, tests/test_xlsx_reference.py.

**Interfaces:** read_reference(path: Path) -> ReferenceData, with TripSummary records, classifications, budgets and transport inputs.

- [x] Write synthetic reference tests for totals, classification recovery, identical duplicates, conflicts and budget source selection.
- [x] Run tests and observe missing import behavior.
- [x] Implement the reader with numeric validation and no writes to the reference.
- [x] Run focused tests.

### Task 2: Preserve original visual layout and recalculating formulas

**Files:** scdp_automation/xlsx_layout.py, scdp_automation/xlsx_output.py, tests/test_xlsx_output.py.

**Interfaces:** install_layout(workbook: Workbook, reference: Path) -> None; validate_layout(workbook: Workbook) -> None; import_reference_workbook(reference: Path, output: Path) -> Path | None.

- [x] Add failing tests for summary styles/merges/dimensions, grouped formulas, transport inputs and actual reference import.
- [x] Copy B1:M58 presentation with public style properties; replace formulas and append additional categories and pending totals.
- [x] Integrate template creation, validation, safe publication and compatibility migration.
- [x] Verify refresh preserves manual values and rejects corrupt formulas.

### Task 3: Produce and reconcile review artifacts

**Files:** scdp_automation/xlsx_output.py, README.md, local ignored output artifacts.

- [x] Use the existing workbook generator; do not add a separate command.
- [x] Generate output/gastos_scdp_2026.xlsx from the original data.
- [x] Recalculate a disposable copy with LibreOffice; compare evaluated formulas with independently aggregated source inputs.
- [x] Compare original and generated visual properties and render PDF previews; report source discrepancies locally.
- [x] Run unittest discover, Ruff lint/format and ty. Review the full diff.

## Execution rulings

The user approved implementation of the correction. Execute inline without another approval gate. The user subsequently requested a commit in this worktree. Preserve the existing deletion of the previous spec in the worktree. The currently opened old workbook is not the destination for this review artifact.

## Validation results

Original data: 124 occurrences consolidated into 111 classified PCDPs. Original summary styles, dimensions, theme, merged cells, budgets and transport inputs reconciled. Original PPGEL + rateio preserved as editable thirds. The saved original summary differs from its own trip records by R$ 3.019,93 across four categories; the local validation report documents this discrepancy. Regression coverage rejects missing rateio for both expenses and budget inputs and prevents classification inheritance on inserted rows. Output and validation report remain ignored local artifacts.

User correction: all changes and outputs stay in this worktree. Replace the existing generator and default workbook; remove the separate reconstruction CLI. Main artifacts created during this correction were removed.
