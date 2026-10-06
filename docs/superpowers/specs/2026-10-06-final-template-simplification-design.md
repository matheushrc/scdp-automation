# Simplificação da automação para o template definitivo

Status: proposta para revisão; implementação posterior em worktree.

## Objetivo e requisitos do usuário

O formato atual da planilha é definitivo. O template estará em `input/gastos_scdp_template.xlsx`. A automação deve coletar o SCDP, atualizar os dados e preservar os preenchimentos manuais sem carregar a infraestrutura dos layouts anteriores.

Requisitos já estabelecidos:

- Manter `APOIO` e `RESUMO GASTOS` no formato do template.
- Deixar a planilha e o JSON atuais na raiz de `output/`.
- Guardar quatro pares anteriores em `output/backup/`, totalizando cinco execuções incluindo a atual.
- Atualizar `RESUMO GASTOS!M2` com a data da atualização dos valores em São Paulo.
- Usar `uv add` para adicionar bibliotecas e `git mv` para mover arquivos rastreados.
- Nunca commitar diagnósticos, dados privados ou resultados da automação.
- Executar a implementação depois, em uma worktree.

Decisão de preservação: numa atualização, a saída existente contém os preenchimentos manuais válidos. Ela será a base da atualização; o template será a base somente quando não houver saída. Não substituir orçamentos preenchidos pela versão do template em cada execução.

## Evidências da revisão

Baseline: commit `c13dd75`. Código de produção: 5.120 linhas; os nove módulos `xlsx_*` somam 3.119 linhas. Os arquivos locais de template e saída anunciam `SCDPLayoutVersion="9"`.

- `xlsx_output.py` cria um layout genérico quando não encontra o template e mantém migrações das versões 3, 5, 6/7 e 8.
- `xlsx_reference.py` importa a antiga `BD D&P`; `canonical_code` pertence a esse fluxo.
- `xlsx_code_summary.py` e `xlsx_description.py` transformam layouts anteriores. A geração das fórmulas de apoio/resumo serve à criação e migração, não à atualização normal do formato final.
- `extract_description`, `load_completed`, `support_range` e `base_range` não têm consumidores no código de produção. As duas últimas tampouco têm consumidores nos testes.
- `collect_pending_descriptions` usa um lock local dentro de um único laço sequencial; nenhuma outra tarefa pode disputar esse lock.
- CLI e extrator repetem a seleção da aba/contexto do navegador. `__init__.py` importa a CLI inteira ao importar qualquer submódulo.
- O JSON tem caminho relativo ao diretório de execução; a planilha usa caminho absoluto do repositório. A execução fora da raiz falha na criação de `OutputHistory`.
- A validação de rateio consulta a coluna 16, atualmente `Segmento`, embora o código de débito esteja na coluna 17. Uma reprodução sintética aceitou gasto de R$ 100 em `PPGEL +` com os três percentuais vazios.
- A retenção é por nome anual. Uma reprodução de seis execuções em cada um de dois anos deixou quatro arquivos atuais na raiz e 16 históricos.
- A revisão executou 160 testes: 159 passaram e um específico do Windows foi ignorado. Ruff e ty passaram. Os testes existentes não detectaram os dois últimos defeitos.

## Alternativas consideradas

1. **Remover somente funções sem chamadas:** baixo impacto, mas preserva quase toda a complexidade das migrações.
2. **Adotar exclusivamente o template final — escolhida:** remove criação alternativa, importação antiga e migrações; mantém validação, atualização e publicação segura.
3. **Reescrever a automação inteira:** aumenta o risco no login, no perfil Windows e na paginação que já funcionam, sem contribuir diretamente para o objetivo.

## Contrato do formato final

O template é obrigatório e somente leitura. Deve conter, nessa ordem, `BASE VIAGENS`, `APOIO` e `RESUMO GASTOS`, com o layout final `SCDPLayoutVersion="9"`. A saída existente deve cumprir o mesmo contrato. Template ausente, inválido ou formato antigo produz erro claro; a automação não tenta converter nem criar outro layout.

`BASE VIAGENS` tem as 18 colunas atuais. L:N são as três datas; O é descrição; P é segmento calculado; Q é código de débito; R é decisão de desconto. As tabelas e intervalos nomeados necessários devem existir e apontar para as colunas finais.

`APOIO` tem as 12 colunas finais: código, nome, segmento, distribuição de diárias/passagens, distribuição de transportes, recurso total, transporte agendado, transporte pago, três percentuais de rateio e total utilizado. O catálogo é lido de `APOIO`; não há catálogo de cursos duplicado em Python.

A validação admite categorias e linhas de resumo acrescentadas pelo operador. Não exige um número fixo de linhas, a ordem original das categorias ou identidade textual de todas as fórmulas manuais. Mantém as validações necessárias de códigos únicos, valores finitos, decisões e rateio.

## Fluxo de atualização

1. Resolver template, JSON, planilha e logs a partir da raiz da worktree/repositório ativo.
2. Validar template e saída sem migrar nem reparar suas worksheets.
3. Serializar a execução de extração por diretório de saída, antes de carregar o checkpoint e acessar o navegador.
4. Coletar a listagem completa e preservar descrições e datas já verificadas do JSON.
5. Arquivar o par anterior uma única vez nesta execução e salvar o JSON atomicamente.
6. Construir uma candidata a partir da saída válida; na primeira execução, copiar o template e limpar somente os dados de BASE.
7. Escrever uma linha por PCDP completa, preservar código e decisão pela chave da PCDP e conservar preenchimentos manuais de APOIO/resumo.
8. Atualizar M2 com uma data Excel, formato `dd/mm/yyyy`, no fuso `America/Sao_Paulo`.
9. Validar a candidata e substituir a saída atomicamente. Aplicar retenção somente depois da publicação bem-sucedida.
10. Consultar descrições pendentes com checkpoints individuais; publicar novamente sem criar outro histórico da mesma execução.

A ausência de uma PCDP já publicada continua impedindo a substituição da planilha. O JSON coletado continua disponível para reconciliação. Um erro não deve apagar a planilha anterior nem os históricos para cumprir artificialmente a retenção.

## Preservação da planilha

O caminho normal não gera nem reescreve fórmulas de `APOIO` ou `RESUMO GASTOS`. Conserva valores, fórmulas, estilos, tabelas, validações, nomes, mesclagens, dimensões, tema e configurações de impressão que o openpyxl suporta. As exceções deliberadas são M2, flags de recálculo do workbook e dados/controles da BASE.

A preservação é semântica: um salvamento XLSX pode reorganizar XML; não se exige identidade binária. Os testes precisam comparar conteúdo e propriedades das worksheets, inclusive fórmulas e referências nomeadas. Não substituir resultados de fórmulas por valores calculados em Python.

Datas e descrição vêm dos registros coletados. Preservar a distinção entre descrição `None` (não consultada) e `""` (consultada e vazia); uma descrição vazia não pode ressuscitar o texto antigo da planilha.

Quando `PPGEL +` tiver orçamento/transporte não zero ou PCDP com total não zero, os três percentuais devem ser números finitos entre 0 e 1 e somar 1 com tolerância absoluta de `1e-10`. Percentuais informados devem ser válidos mesmo sem gastos. A checagem de PCDP consulta Q, inclusive após escrever os dados novos na candidata.

## Histórico e publicação

A retenção considera todos os anos conjuntamente: um par atual e, no máximo, quatro pares históricos depois de uma publicação bem-sucedida. Cada par histórico tem identificador compartilhado de execução no nome da planilha e do JSON.

Na virada do ano, o par do ano anterior é arquivado em `output/backup/`; sai da raiz somente após a publicação bem-sucedida do novo par. Em caso de falha, os arquivos antigos permanecem recuperáveis. Na primeira execução de um ano novo, não copiar BASE ou classificações do ano anterior para o novo ano.

Só manipular arquivos reconhecidos como saídas da automação. Não apagar arquivos manuais desconhecidos, não inventar um JSON para uma planilha órfã e não destruir um arquivo órfão para completar pares.

Existe um único proprietário do arquivamento e da retenção: `OutputHistory`. Remover o backup de cada publicação em `xlsx_output`. O lock da execução fica fora de `output/`, em local temporário determinado pelo caminho absoluto do diretório de saída; não criar novos locks visíveis na raiz de output.

Se Excel impedir a substituição, preservar a saída e colocar a candidata validada em `output/backup/recovery.<id>.xlsx`, indicando o caminho no erro. Essa candidata não conta como execução completa e pode ser removida após uma publicação posterior bem-sucedida. Não criar outras subpastas.

## Estrutura de código pretendida

- `xlsx_models.py`: `TripSummary` e agregação de `Viagem`; remover `DebitCategory`, catálogo fixo e campo `review_required` quando seus consumidores legados forem retirados.
- `xlsx_validation.py`: contrato final, constantes de colunas, erro de validação e validação sem mutação. Renomear `xlsx_adjustments.py` com `git mv`, conservando somente o necessário.
- `xlsx_output.py`: carregar/copiar workbook, preservar entradas por PCDP, escrever BASE/M2, validar candidata e publicar. Remover os modos antigo/genérico e a geração de fórmulas de apoio/resumo.
- `output_history.py`: serialização da execução, arquivo anterior, virada do ano e retenção global.
- `navegador_chrome.py`: conexão existente e pequeno helper compartilhado de seleção da página/contexto.
- `config.py`: caminhos e data/fuso comuns; manter a seleção do perfil e o registro do ano sem criar um novo sistema de configuração.
- `tests/support/workbooks.py`: fixture sintética diretamente no formato final. Não montar o formato final passando pelas migrações que serão eliminadas.
- `tests/support/recalculation.py`: recálculo com LibreOffice usado para verificar resultados das fórmulas nos testes. Mover `xlsx_recalculate.py` com `git mv`; remover a opção de recálculo do publisher, pois a CLI não a usa.

Remover de produção `xlsx_reference.py`, `xlsx_code_summary.py`, `xlsx_description.py` e `xlsx_layout.py` depois de transferir os poucos contratos/validadores necessários. Não mover as migrações inteiras para outro lugar.

Remover as funções sem consumidores mencionadas na revisão, o lock local sequencial e o import da CLI em `__init__.py`. Unificar seleção da página e resolução de caminhos. `python-dotenv` deixa de ser dependência direta se nenhuma importação direta permanecer; a leitura de `.env` por pydantic-settings continua funcionando. Não adicionar bibliotecas.

## Limites e invariantes

- Python `>=3.14.5, <3.15`; manter uv, Playwright, openpyxl, Pydantic, filelock e loguru.
- Manter Windows e Linux, Chrome visível, cópia sem extensões e caminhos longos Windows.
- Preservar `--login`, `--abrir-navegador`, `--selecionar-perfil-chrome` e `--limite`.
- Preservar autenticação gov.br e resolução de CAPTCHA pelo operador.
- Preservar paginação, prevenção de páginas repetidas, parsing, decisões manuais, checkpoints e a proteção contra PCDPs ausentes.
- Não reescrever clonagem de perfil, protocolos de login ou seletores do SCDP nesta mudança.
- A execução usará a skill `superpowers:using-git-worktrees`; preferir ferramenta nativa quando disponível, senão `.worktrees/final-template-simplification` na branch `refactor/final-template-simplification`.
- Nunca usar input privado como fixture de teste ou commitar `.env`, perfis, input, output, logs e scripts de diagnóstico.
- Helpers compartilhados ficam em `tests/support/`; testes importam produção diretamente.

## Critérios de aceitação

1. Template final obrigatório; arquivo ausente/antigo é recusado sem alterar saída ou template.
2. Primeira extração parte do template com BASE vazia; atualização preserva códigos/decisões por PCDP completa e preenchimentos manuais das duas worksheets.
3. Nenhum fluxo executável importa/converte `BD D&P`, migra versões antigas ou gera um resumo alternativo.
4. Rateio inválido com gastos é recusado; descrição vazia permanece vazia; M2 recebe a data correta.
5. Raiz contém somente o par atual reconhecido; quatro pares anteriores ficam em backup, inclusive depois da virada do ano.
6. Falha de cópia/publicação não perde o estado anterior; checkpoints e mensagens permitem retomar.
7. Execução fora da raiz resolve os mesmos caminhos; execuções concorrentes não misturam checkpoints/históricos nem disputam a mesma página.
8. Os comandos públicos e as proteções Windows/Linux continuam funcionando.
9. Testes partem de dados sintéticos finais e verificam comportamento, preservação e resultados financeiros; a suíte, Ruff e ty passam.
10. Não perseguir uma meta arbitrária de linhas. A redução deve vir da exclusão dos fluxos obsoletos, sem retirar garantias necessárias.

## Verificação e entrega

Usar unittest, Ruff e ty conforme AGENTS.md. Os testes financeiros com LibreOffice usam cópias sintéticas temporárias e preservam testes de fórmulas, cancelamentos, direção/subcategorias, rateio e categorias acrescentadas. No Windows, conferir abertura no Excel e recálculo antes de afirmar compatibilidade visual completa; testes Linux não substituem essa conferência.

A execução posterior lê esta spec e o plano correspondente, cria a worktree e verifica o baseline. Os arquivos privados e o output do checkout principal não são compartilhados como destinos de escrita. Não implementar, mesclar ou publicar alterações nesta etapa de documentação.
