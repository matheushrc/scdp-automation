# SCDP official workbook update design

## Status

Proposed design for review. This worktree contains the design only; no writer has
been implemented and the official workbook has not been modified.

## Goal

Refresh the repository-local official workbook
`input/.Diárias-Pass-Transp 2026 - Consulta Saldos.xlsx` whenever a complete
annual SCDP report listing has been collected. Write the report data into the
`BD D&P` worksheet, columns `D:S`, while preserving the workbook's existing
layout, formulas, summaries, and pivot tables.

The workbook is the recurring official output. A normal successful extraction
must update it automatically; operators should not need to promote a candidate
manually after each run. The update must remain recoverable if writing or
validation fails.

## Current behavior and workbook structure

`scdp_automation.extrator` currently collects and validates the annual report,
then writes an atomic JSON checkpoint at
`output/viagens_scdp_2026.json`. It saves that checkpoint after each reason
description so an interrupted enrichment pass can resume. The JSON remains the
recovery checkpoint; the workbook becomes the refreshed report destination.

The target worksheet has report headers in `D3:S3`:

| Column | Report field |
| --- | --- |
| D | Número da Solicitação |
| E | Nome do Proposto |
| F | Órgão Solicitante |
| G | Órgão Superior |
| H | Tipo da Viagem |
| I | Situação da Viagem |
| J | Motivo Viagem |
| K | Início |
| L | Término |
| M | Origem |
| N | Destino |
| O | Meio de Transporte |
| P | Quantidade Diárias |
| Q | Diárias (R$) |
| R | Passagens e Taxas Iniciais (R$) |
| S | Total (R$) |

One request can contain multiple itinerary segments and subtotal/footer rows.
The writer must render the validated `Viagem` model back into the same row
structure, including segment rows and the cost/subtotal/total rows consumed by
the worksheet. It must preserve multi-segment trips and complementary PCDP
suffixes.

`BD D&P` also contains a formula-derived summary in `Z:AN`; those cells are not
an input range and must never be overwritten by the extractor. The workbook has
five sheets, a hidden support sheet, extensive formulas and merged cells, and
two pivot caches. Preserve those features and all content outside the target
report block. The workbook is ignored under `input/` and must not be committed.

## Proposed update flow

1. Collect every page of the annual report and validate the complete set of
   `Viagem` records. If pagination, parsing, or validation fails, do not change
   the workbook.
2. Save the complete listing to the existing atomic JSON checkpoint, retaining
   any previously collected descriptions by PCDP. Write this checkpoint before
   attempting the workbook update so a workbook failure cannot discard
   extraction progress.
3. Render the records into the existing `BD D&P` report block in `D:S`, using
   the headers above and the current workbook's row, style, and merge layout.
   Treat the block as a full snapshot: replace stale report rows in that block
   rather than appending duplicate rows. Determine the fixed block boundaries
   and capacity during the compatibility gate. Do not insert or delete
   worksheet rows; if the rendered report does not fit the existing block or
   its merge layout cannot represent it, abort before promotion and revise this
   design. Do not change `A:C`, `T:X`, `Z:AN`, or any other worksheet.
4. Build a candidate workbook beside the official file. Validate the candidate
   before promotion: it opens as an XLSX package, contains the expected report
   rows, retains the workbook structure, and has no duplicate PCDP keys in the
   rendered report. Remove an invalid/incomplete candidate. If final promotion
   fails after validation, retain the candidate for retry and report its path.
5. Create a timestamped backup of the current official workbook. If backup
   creation fails, abort. After validation and backup succeed, atomically
   replace the official path with the candidate. Keep backups; do not prune them
   automatically in the first release.
6. Log the updated workbook path and backup path. Log description-enrichment
   progress separately from workbook publication, because descriptions are
   stored in the JSON checkpoint and are not columns in `D:S`.

Publish after the complete report listing is collected, before optional reason
description enrichment. This means `--limite` does not block an XLSX refresh,
and an enrichment failure does not leave the report snapshot stale. A failure
before the listing is complete must leave the official workbook unchanged.

The default workbook path must resolve from the current checkout's repository
root rather than the shell's current directory. If the file is missing, fail
with a clear message and do not create an empty workbook. The ignored official
workbook is not present in this feature worktree, so tests and compatibility
experiments must inject a temporary destination or use a local copy; they must
never write the official file.

## Preservation and failure behavior

- Write only the report-owned cells in `BD D&P!D:S`. Keep the formula-derived
  summary, formulas, pivot tables, caches, defined names, hidden sheets, and
  unrelated workbook content intact.
- Do not use a workbook-writing library until a round-trip experiment against a
  copy of this workbook proves those preservation requirements. If the writer
  cannot retain them, reject that approach rather than silently degrading the
  official file.
- Keep the JSON checkpoint independent of workbook replacement. A workbook
  write failure must not discard extraction progress.
- Serialize concurrent workbook updates with a per-file lock. If the workbook
  is open or the operating system refuses replacement, leave the official file
  intact and report that the operator should close the workbook and retry.
- On any failure before atomic replacement, leave the official workbook intact
  and retain enough information to diagnose the failure. Delete incomplete
  candidates; if a validated candidate cannot be promoted, retain it for retry.
  Never log names, descriptions, or full travel records.
- A successful full-snapshot refresh clears obsolete report rows only inside
  the designated `D:S` report block. Columns `A:C` are outside the writer's
  scope and must be preserved along with all calculated summary columns.
- Preserve the workbook's existing pivot-cache refresh behavior and formula
  calculation settings. Confirm refresh and calculation behavior by opening a
  candidate in desktop Excel before declaring the writer compatible.

## Compatibility gate and validation

Before selecting the production writer, use a local copy of the official
workbook to prove that a write to `D:S` can preserve the workbook. Compare the
candidate with the source copy and verify sheet names/visibility, merged ranges
outside the target block, formulas, tables, defined names,
pivot-table/cache parts and relationships, and all cell contents outside
`D:S`. Confirm the written report
matches the parsed records for a trip with multiple segments, subtotal/footer
rows, zero costs, a cancelled status, and a complementary PCDP.

Unit tests should cover field mapping, row rendering, full-snapshot idempotence,
stale-row clearing within `D:S`, duplicate-key rejection, candidate validation,
backup creation, replacement failure, locked-file failure, and preservation of
the existing JSON checkpoint. Tests use synthetic `Viagem` objects and
temporary workbook copies. No test runs a live SCDP extraction or writes to the
official workbook.

Run the repository checks from the worktree:

- `uv run python -m unittest discover -v`
- `uv run ruff check .`
- `uv run ruff format --check .`
- `uv run ty check`

On Windows, manually open a candidate in desktop Excel and verify the report,
formulas, and pivot summaries refresh correctly. Linux unit tests cannot prove
Excel's Windows file-lock and recalculation behavior.

## Scope and review assumptions

- The report block in `D:S` is generated from the current annual SCDP listing;
  each successful collection replaces that block as a snapshot.
- Columns `A:C` are outside the writer's scope and remain untouched.
- JSON stays as the resumable checkpoint; removing or replacing it is out of
  scope.
- Browser startup, authentication, scraping concurrency, and gov.br behavior
  do not change.

The row-block boundaries, available capacity, and merge rules still need to be
confirmed against the workbook copy during the compatibility gate. If rendering
requires changing worksheet rows or merged regions outside `D:S`, stop and
revise this design before implementation.
