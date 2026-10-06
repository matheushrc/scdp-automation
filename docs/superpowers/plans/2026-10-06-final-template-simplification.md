# Final Template Simplification Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Simplificar a automação para usar exclusivamente o template final, preservando dados manuais e corrigindo rateio, caminhos e retenção entre anos.

**Architecture:** O template final fornece a estrutura; a saída existente fornece os preenchimentos manuais. A atualização escreve BASE/M2 e publica uma candidata validada, enquanto `OutputHistory` controla a execução e os pares históricos. Remover migrações, criação alternativa e utilitários sem consumidores; manter Chrome, login, parsing e paginação.

**Tech Stack:** Python 3.14, uv, Playwright, openpyxl, Pydantic, filelock, loguru, unittest, Ruff, ty; LibreOffice apenas na verificação de fórmulas.

**Spec:** `docs/superpowers/specs/2026-10-06-final-template-simplification-design.md`

Método previsto: execução nativa pelo agente na worktree, usando `superpowers:executing-plans`; sem delegação automática.

Status: plano proposto para revisão. A execução em worktree foi solicitada para uma etapa posterior; não iniciar implementação junto com a escrita deste documento.

## Global Constraints

- Python `>=3.14.5, <3.15`; manter uv, Playwright, openpyxl, Pydantic, filelock e loguru.
- Manter Windows e Linux, Chrome visível, cópia sem extensões e caminhos longos Windows.
- Preservar `--login`, `--abrir-navegador`, `--selecionar-perfil-chrome` e `--limite`.
- Preservar autenticação gov.br e resolução de CAPTCHA pelo operador.
- Preservar paginação, prevenção de páginas repetidas, parsing, decisões manuais, checkpoints e a proteção contra PCDPs ausentes.
- Não reescrever clonagem de perfil, protocolos de login ou seletores do SCDP nesta mudança.
- Nunca usar input privado como fixture de teste ou commitar `.env`, perfis, input, output, logs e scripts de diagnóstico.
- Helpers compartilhados ficam em `tests/support/`; testes importam produção diretamente.
- Usar `uv add` para adicionar bibliotecas e `git mv` para mover arquivos rastreados. Não adicionar bibliotecas nesta mudança.
- Template obrigatório: `input/gastos_scdp_template.xlsx`, layout `SCDPLayoutVersion="9"`, somente leitura.
- Cinco execuções contando a atual: quatro pares anteriores em `output/backup/`, com retenção conjunta entre anos.
- Apenas BASE, M2 e flags de recálculo podem mudar automaticamente; preservar APOIO/resumo existentes.

## Review Focus

- Descrição consultada e vazia: não recuperar a descrição antiga da planilha; teste na tarefa 2.
- Rateio sem orçamento mas com gasto recém-coletado: conferir Q, exigir percentuais que somem 1; testes nas tarefas 1 e 2.
- Virada do ano com publicação bloqueada pelo Excel: manter o par anterior recuperável e não podar históricos; teste na tarefa 3.
- Arquivos órfãos/manuais em output: não apagar nem completar pares artificialmente; teste na tarefa 3.
- Duas execuções no mesmo output: impedir mistura do JSON, snapshots e operações da página; teste de lock entre processos na tarefa 3.

## Estrutura final e interfaces

| Arquivo | Responsabilidade |
| --- | --- |
| `scdp_automation/xlsx_models.py` | `TripSummary`, `summarize_trips`; nenhuma lista fixa de categorias |
| `scdp_automation/xlsx_validation.py` | `BASE_HEADERS`, `SUPPORT_HEADERS`, `WorkbookValidationError`, `validate_workbook(workbook: Workbook) -> None`; validação sem mutação |
| `scdp_automation/xlsx_output.py` | `build_candidate(trips: Sequence[Viagem], current_path: Path, candidate_path: Path, *, template_path: Path | None = None) -> None`; `publish_workbook(trips: Sequence[Viagem], workbook_path: Path, *, template_path: Path | None = None) -> None` após tarefa 3 |
| `scdp_automation/output_history.py` | `OutputHistory(checkpoint: Path, workbook: Path)` como context manager; `archive_previous() -> Path | None`, `prune() -> None` |
| `scdp_automation/config.py` | `REPO_ROOT`, `current_date() -> date`, `current_year() -> int`; configuração existente |
| `scdp_automation/navegador_chrome.py` | conexão CDP; `select_scdp_page(browser: Browser) -> Page` |
| `tests/support/workbooks.py` | `final_template_fixture(path: Path) -> None`, builders de viagens e leitura de saídas sintéticas |
| `tests/support/recalculation.py` | `recalculate_workbook(path: Path) -> None`, somente testes |

A validação e o publisher podem manter temporariamente exports para seus consumidores existentes durante uma tarefa. Remover os exports antigos na tarefa 4 depois de atualizar todos os imports. Não deixar adaptadores de compatibilidade na entrega final.

## Preparação da execução em worktree

- [ ] Ler spec, plano e AGENTS.md; preservar alterações/documentos locais anteriores. Confirmar que o HEAD inclui `c13dd75` e estes documentos.
- [ ] Aplicar `superpowers:using-git-worktrees`: detectar isolamento existente e ferramentas nativas primeiro. Sem ferramenta nativa, verificar `git check-ignore .worktrees/` e executar `git worktree add .worktrees/final-template-simplification -b refactor/final-template-simplification`. A worktree é preferência explícita; não substituir por implementação no checkout principal.
- [ ] Usar `uv sync` na worktree. Rodar baseline: `uv run python -m unittest discover -v`, `uv run ruff check .`, `uv run ruff format --check .`, `uv run ty check`. Baseline da revisão: 160 testes, 159 passando e um ignorado no Linux.
- [ ] Criar fixtures somente em diretórios temporários. Não copiar nem vincular `.env`, `.scdp-browser`, output ou input privados para a worktree. Um smoke operacional posterior exige configuração local explícita; não é requisito para os testes unitários.

## Tarefa 1: fixture final independente e regressão de rateio

**Arquivos:** modificar `tests/support/workbooks.py`, `tests/test_spending_adjustments.py`, `tests/test_xlsx_fidelity.py`, `scdp_automation/xlsx_adjustments.py`.

**Interfaces:** consome `Viagem`, `BASE_HEADERS` e openpyxl; produz `final_template_fixture(path: Path) -> None` para todas as tarefas seguintes. A validação existente permanece acessível até a tarefa 2.

- [ ] Adicionar fixture sintética diretamente com as três worksheets, 18 colunas na BASE, 12 em APOIO, versão 9, tabelas, validações, intervalos nomeados e M2. Incluir categorias suficientes para Agronomia, PPGE, PPGEL, História, PPGEL +, Direção/subcategoria e Afastamento; fórmulas representativas dos cálculos atuais. Construir com openpyxl no helper, sem chamar importador, migrador ou publisher de produção e sem ler input privado.
- [ ] Escrever `test_rateio_required_for_classified_expense_without_budget`: BASE Q2=`PPGEL +`, K2=100; distribuição e percentuais vazios. Esperar erro contendo `100%`. Cobrir percentuais `(0.5, 0.25, 0.25)` válidos, `(0.5, 0.25, 0.20)` inválidos, NaN, booleano e ausência de gasto/orçamento com percentuais vazios.

Asserção central do caso de regressão, usando a fixture sintética:

```python
base["Q2"] = "PPGEL +"
base["K2"] = 100
for column in (4, 5, 7, 8, 9, 10, 11):
    support.cell(rateio_row, column).value = None
with self.assertRaisesRegex(WorkbookValidationError, "100%"):
    validate_editable_layout(workbook)
```

- [ ] Rodar `uv run python -m unittest tests.test_spending_adjustments -v`; confirmar falha do caso Q2=PPGEL + no código atual.
- [ ] Corrigir a checagem de coluna usando o cabeçalho final de código de débito, sem alterar a regra de despesas/decisões. Adaptar os testes de fidelidade do fluxo final para partir da nova fixture; comparar fórmulas, valores, estilos, mesclagens, validações, dimensões e nomes de APOIO/resumo, exceto M2.
- [ ] Rodar os dois módulos e verificar que os testes existentes de migração ainda passam enquanto seus consumidores permanecem.
- [ ] Commit específico: `fix rateio validation for final workbook columns`.

## Tarefa 2: atualização exclusiva do template e retirada de migrações

**Arquivos:** `git mv scdp_automation/xlsx_adjustments.py scdp_automation/xlsx_validation.py`; modificar `xlsx_output.py`, `xlsx_models.py`, `extrator.py`, `tests/support/workbooks.py` e todos os testes XLSX; remover `xlsx_reference.py`, `xlsx_code_summary.py`, `xlsx_description.py`, `xlsx_layout.py` e `tests/test_xlsx_reference.py`; `git mv scdp_automation/xlsx_recalculate.py tests/support/recalculation.py`.

**Interfaces:** consome `final_template_fixture`, `Viagem` e `summarize_trips`; produz `validate_workbook`, `BASE_HEADERS`, `SUPPORT_HEADERS`, `WorkbookValidationError` no módulo de validação, e `build_candidate` com o parâmetro `template_path`. O publisher conserva o backup por publicação temporariamente até a tarefa 3, mas já deixa de oferecer recálculo/importação antiga.

- [ ] Escrever `test_missing_template_preserves_published_files`, `test_old_template_is_rejected`, `test_old_output_is_rejected`, `test_fresh_base_does_not_import_template_trips`, `test_refresh_preserves_manual_sheets_and_full_pcdp_choices` e `test_empty_description_does_not_restore_old_value`. Nos erros, conferir bytes originais e ausência de promoção. Na atualização, inverter PCDPs e incluir uma complementação para provar preservação por chave.

Asserções de template obrigatório e descrição vazia:

```python
with self.assertRaises(WorkbookValidationError):
    build_candidate(trips, current, candidate, template_path=missing_template)
self.assertEqual(current.read_bytes(), original_bytes)
trip.descricao_do_motivo_da_viagem = ""
build_candidate([trip], current, candidate, template_path=template)
updated = load_workbook(candidate)
self.addCleanup(updated.close)
self.assertIn(updated["BASE VIAGENS"]["O2"].value, (None, ""))
```

O XLSX pode representar string vazia como célula vazia; o JSON deve continuar contendo `""`, nunca `None` nem a descrição anterior.

- [ ] Escrever `test_new_expense_requires_rateio`: saída sem despesas PPGEL +, nova viagem classificada nessa categoria com total 100 e percentuais vazios; a candidata deve ser recusada após a escrita da BASE. Fixar M2 em `07/10/2026` por clock mockado em São Paulo e verificar formato `dd/mm/yyyy`.
- [ ] Rodar `uv run python -m unittest tests.test_xlsx_fidelity tests.test_spending_adjustments -v`; os testes dos novos contratos devem falhar antes da implementação.
- [ ] Implementar validação exclusivamente final, sem reparar workbook: worksheets, cabeçalhos, versão, tabelas, nomes essenciais, códigos únicos, valores finitos, decisões e rateio. Transferir somente comparação de fórmulas/contratos realmente necessários dos módulos removidos; eliminar geração e migração de apoio/resumo. Usar nomes finais de colunas em vez de mapas de posições antigas.
- [ ] Simplificar `build_candidate`: template obrigatório e validado; copiar template somente para a primeira saída, senão carregar saída existente; atualizar BASE e M2; conservar entradas manuais por PCDP; manter guarda contra PCDPs ausentes, dados agregados, estilo/controles de BASE e publicação atômica. Aceitar apenas `Sequence[Viagem]`; agregar a lista uma única vez. Não exigir uma fórmula fixa para cada célula manual de APOIO/resumo.
- [ ] Remover catálogo fixo, `DebitCategory`, importação `BD D&P`, `canonical_code`, exports e parâmetros legados. Adaptar fixtures/testes ao formato final antes de retirar os helpers `reference_fixture` e `restore_legacy_*`. O helper `write_existing_workbook` passa a usar `final_template_fixture`, não a criação de produção.
- [ ] Mover recálculo com `git mv`, atualizar imports dos testes e remover `recalculate` do publisher. Manter verificação financeira em `test_xlsx_code_summary`, `test_spending_adjustments` e `test_xlsx_recalculate`; substituir casos de migração por casos de preservação/expansão do formato final. Não manter o velho gerador completo como fixture disfarçada.
- [ ] Rodar toda a suíte e checks. Verificar ausência de importações dos módulos removidos com `rg`. Validar resultados recalculados em cópias sintéticas: cancelada Sim/Não/pendente, restituição, rateio, Direção/subcategoria e categoria adicionada, preservando as expectativas financeiras úteis dos testes anteriores.
- [ ] Commit específico: `simplify workbook updates to the final template`.

## Tarefa 3: único histórico por execução e retenção entre anos

**Arquivos:** modificar `output_history.py`, `extrator.py`, `xlsx_output.py`, `tests/test_output_history.py`, `tests/test_extrator.py`, `tests/test_xlsx_output.py`.

**Interfaces:** `OutputHistory` mantém os caminhos atuais e passa a implementar `__enter__(self) -> OutputHistory` e `__exit__(self, exc_type, exc_value, traceback) -> None`. O context manager detém um FileLock por diretório absoluto de saída, fora de output, timeout de 30 segundos. `archive_previous()` continua idempotente por objeto; `prune()` finaliza a retenção após publicação. `publish_workbook` passa a retornar `None`, sem `create_backup` e sem criar backups próprios.

- [ ] Escrever `test_history_is_global_across_years`: seis execuções em 2026 e seis em 2027, raízes temporárias; ao final exatamente dois arquivos atuais de 2027 e oito arquivos históricos, formando os quatro pares anteriores. Verificar o conteúdo dos snapshots, não apenas nomes/quantidade.

Asserções do cenário completo de virada do ano, sem arquivos manuais no cenário:

```python
self.assertEqual(
    {p.name for p in output.iterdir()},
    {"gastos_scdp_2027.xlsx", "viagens_scdp_2027.json", "backup"},
)
self.assertEqual(len(list((output / "backup").iterdir())), 8)
self.assertEqual(archived_json_values, expected_previous_four_values)
self.assertEqual(archived_workbook_values, expected_previous_four_values)
```

- [ ] Escrever `test_year_rollover_failure_preserves_previous_pair`, `test_orphan_and_manual_files_are_preserved`, `test_two_publications_archive_once` e `test_copy_failure_keeps_originals`. Em falhas, não remover o par antigo nem podar o histórico. Não sintetizar pares para arquivo órfão.
- [ ] Escrever `test_output_session_serializes_processes` com dois processos, sinalização explícita de aquisição e diretório temporário: a segunda sessão não entra enquanto a primeira detém o lock. Sem sleeps usados como única evidência. Cobrir liberação por exceção. A integração do extrator deve mostrar que o lock cobre carregar JSON, navegador, checkpoints e publicações.
- [ ] Rodar `uv run python -m unittest tests.test_output_history tests.test_extrator -v`; confirmar falhas dos casos de ano/lock no baseline.
- [ ] Implementar arquivamento e seleção global de pares gerenciados usando o identificador compartilhado para ordenar execuções. Na troca de ano, copiar o par anterior para backup, usar template limpo para o novo ano e remover os arquivos antigos da raiz apenas após o novo par ser publicado. `prune()` mantém quatro pares anteriores, sem tocar arquivos desconhecidos/órfãos.
- [ ] Envolver a execução do extrator em `with OutputHistory(...) as history`; remover backup e lock exclusivos do publisher. Atualizar `save_checkpoint_and_publish(trips: list[Viagem], checkpoint_path: Path, workbook_path: Path, *, history: OutputHistory) -> Path | None` para uma única estratégia obrigatória. O retorno é o snapshot anterior fornecido por `history`, usado no log.
- [ ] Em falha de `os.replace`, guardar a candidata validada em `output/backup/recovery.<id>.xlsx` e informar o caminho. Testar original intacto, ausência de candidata na raiz e remoção de recovery gerenciada apenas depois de uma publicação posterior válida. Atualizar testes antes centrados em backups por publicação para validar o coordenador.
- [ ] Rodar os três módulos e checks; commit `keep output history bounded across extraction years`.

## Tarefa 4: remover utilitários mortos e duplicações do fluxo

**Arquivos:** modificar `extrator.py`, `relatorio.py`, `cli.py`, `navegador_chrome.py`, `config.py`, `logging_config.py`, `__init__.py`, `pyproject.toml`, `uv.lock` e respectivos testes.

**Interfaces:** `current_date() -> date` em config, `current_year()` deriva dela. `select_scdp_page(browser: Browser) -> Page` em navegador valida contexto, prefere aba SCDP/gov.br existente e cria uma só quando não há páginas. CLI/extrator continuam responsáveis por login e navegação conforme sua ação.

- [ ] Escrever `test_paths_are_checkout_relative_from_other_cwd` com `chdir` temporário e configuração mockada: JSON, planilha, template e logs devem continuar sob REPO_ROOT. Rodar o caso para confirmar a falha atual.

Asserções de resolução dos destinos, após mudar para outro cwd e mockar a gravação de configuração:

```python
args = parse_args([])
self.assertEqual(args.output, REPO_ROOT / "output" / "viagens_scdp_2026.json")
self.assertEqual(args.workbook, REPO_ROOT / "output" / "gastos_scdp_2026.xlsx")
self.assertEqual(DEFAULT_REFERENCE, REPO_ROOT / "input" / "gastos_scdp_template.xlsx")
```

- [ ] Escrever testes de seleção de página: aba SCDP existente entre outras abas, aba gov.br, contexto sem páginas e navegador sem contexto. `--abrir-navegador` deve abrir SCDP; `--login` pode continuar na aba gov.br. Não adicionar ação de navegação dentro do helper de seleção.
- [ ] Centralizar caminhos e data; remover defaults anuais calculados por import quando não tiverem consumidor necessário. Tornar `__init__.py` livre de importação da CLI; `__main__.py` e o entry point do pyproject continuam chamando `cli.main`.
- [ ] Substituir seleção duplicada da aba pelo helper, remover `extract_description`, `load_completed` e o lock local de `collect_pending_descriptions`. Preservar consultas sequenciais e checkpoint após cada consulta. Remover testes que apenas exercitavam funções retiradas; manter testes de extração DOM e retomada.
- [ ] Atualizar imports para seus módulos proprietários; retirar exports/adaptadores antigos. Conferir dependências realmente importadas. Se `python-dotenv` seguir apenas como dependência transitiva de pydantic-settings, executar `uv remove python-dotenv` e `uv sync`; manter testes de prioridade `.env` sobre USERNAME do Windows e de sigilo das credenciais. Para qualquer adição futura usar `uv add`.
- [ ] Rodar módulos de CLI/extrator, configuração, autenticação, navegador, relatório e logging, depois checks; commit `remove obsolete helpers and centralize runtime paths`.

## Tarefa 5: documentar e verificar a entrega integrada

**Arquivos:** modificar `README.md`, `AGENTS.md`; ajustar somente os testes finais afetados por achados da revisão.

**Interfaces:** documentação usa somente comandos públicos existentes e os contratos entregues nas tarefas 1–4.

- [ ] Atualizar README: template final obrigatório, erro para formatos antigos, preservação manual, M2, histórico conjunto entre anos, recuperação de Excel bloqueado e caminhos independentes do cwd. Retirar instruções de layout alternativo/importação antiga. Em AGENTS, registrar que o template é a fonte do layout e não se devem reintroduzir migradores ou geradores de apoio/resumo sem requisito explícito. Manter as regras sobre helpers, git mv e diagnósticos.
- [ ] Rodar `uv run python -m unittest discover -v`, `uv run ruff check .`, `uv run ruff format --check .`, `uv run ty check` na worktree. Registrar número de testes e skips reais; a retirada de testes legados não dispensa cobertura financeira/fidelidade.
- [ ] Rodar `uv run python -m scdp_automation --help` e `uv run scdp-extrair --help`; conferir opções públicas. Validar que importar `scdp_automation.relatorio` não carrega a CLI/Playwright da automação.
- [ ] Fazer revisão integrada contra cada critério de aceitação da spec, procurando imports mortos, referências a BD D&P, versões antigas, fórmulas regeneradas em apoio/resumo, caminhos do checkout principal e backups duplicados. Conservar proteções necessárias mesmo se elas não reduzirem linhas.
- [ ] Para qualquer verificação real de SCDP, usar Playwright conectado ao Chrome local visível da worktree, credenciais locais e CAPTCHA pelo operador, com destinos privados isolados. Se não houver esse ambiente, registrar a limitação e não afirmar validação operacional/visual no Windows. Não executar a automação contra output do checkout principal.
- [ ] Revisar `git diff --check`, diff completo e arquivos staged. Commit `document the final template extraction contract`, incluindo apenas documentação e eventuais correções verificadas desta tarefa. Nunca incluir código de diagnóstico.
- [ ] Entregar branch/worktree, commits, resultados dos checks e limitações reais. Manter a worktree para revisão; não mesclar, publicar ou remover automaticamente.

## Autorrevisão do plano

- Contrato final e retirada de legado: tarefa 2.
- Rateio, descrição vazia, M2 e preservação por PCDP: tarefas 1 e 2.
- Retenção, virada do ano, falhas, arquivos manuais e concorrência: tarefa 3.
- Caminhos, imports leves, funções sem consumidores e duplicação de navegador: tarefa 4.
- Cobertura financeira, documentação, confidencialidade e entrega na worktree: tarefa 5.
- A clonagem Windows/Linux, autenticação e paginação não são redesenhadas; sua suíte continua sendo requisito da entrega.
