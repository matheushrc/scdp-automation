# SCDP

Extrator do relatório **Relatórios > Viagem** do SCDP. A automação usa a janela visível do Google Chrome e mantém a autenticação gov.br e o CAPTCHA sob controle do operador.

## Preparar e executar

Instale as dependências do projeto:

```sh
uv sync
```

Crie `.env` a partir do exemplo e restrinja as permissões do arquivo:

```sh
cp .env.example .env
chmod 600 .env
```

Configure as credenciais gov.br em `.env`, que é ignorado pelo Git:

```dotenv
USERNAME=seu-cpf
PASSWORD=sua-senha
```

Na primeira execução, o terminal lista os perfis Chrome encontrados para o
usuário atual. Use as setas e Enter para escolher; Escape cancela. O nome e o
e-mail exibidos vêm dos metadados locais do Chrome, e o e-mail fica vazio quando
o perfil não tem um associado. O script copia o perfil escolhido para
`.scdp-browser/` sem extensões e grava o caminho de origem e o e-mail em
`.scdp-config.toml`, ignorados pelo Git. Essa configuração local não substitui
as credenciais de `.env`.

Feche todas as janelas do Chrome antes da primeira cópia ou de trocar o perfil
salvo. Para trocar o perfil, execute:

```sh
uv run python -m scdp_automation --selecionar-perfil-chrome
```

Também é possível combinar essa opção com `--login` ou `--abrir-navegador`.
Nas execuções seguintes, o clone local é reutilizado.

O extrator preenche o CPF na página gov.br, aguarda você resolver o CAPTCHA
na janela visível, preenche a senha e envia o formulário. Também é possível
definir `USERNAME` e `PASSWORD` no ambiente como fallback quando as respectivas
chaves não estiverem presentes no `.env`.

Para abrir apenas o Chrome visível com o perfil local do projeto:

```sh
uv run python -m scdp_automation --abrir-navegador
```

Para executar a coleta, em outro terminal na raiz do projeto:

```sh
uv run python -m scdp_automation
```

Para autenticar sem iniciar a coleta, execute:

```sh
uv run python -m scdp_automation --login
```

Esse comando para após voltar ao SCDP autenticado. Ele não abre relatórios nem
cria ou atualiza o JSON de viagens.

Se for solicitado, o extrator clica em **Entrar com gov.br**. O CPF e a senha vêm da configuração local; resolva o CAPTCHA manualmente na janela visível. O extrator aguarda o retorno ao SCDP autenticado e a aparição do menu **RELATÓRIOS**. Em seguida, abre **RELATÓRIOS > Viagem**, confirma CCH — Campus Chapecó/SC (código `121766`), marca **Todas as viagens do ano de exercício** e pesquisa. Mantenha a janela do Chrome aberta enquanto a coleta estiver em execução.

## Saída e retomada

O resultado e checkpoint ficam em `output/viagens_scdp_2026.json`. O arquivo contém as solicitações, os trechos e os totais visíveis no relatório, mais `descricao_do_motivo_da_viagem`, consultada na tela **Situação da Solicitação**. Campos usam snake case, valores monetários e quantidade de diárias são números float, e todas as viagens são validadas com Pydantic.

Primeiro o script percorre todas as páginas do relatório e grava os dados. Depois consulta as descrições pendentes, atualizando o JSON atomicamente depois de cada uma. Uma descrição `null` indica que ainda não foi consultada; uma string vazia indica que o campo foi consultado e estava vazio. Em nova execução, descrições já concluídas são reutilizadas e as pendentes são retomadas. Para limitar a quantidade consultada em uma execução:

```sh
uv run python -m scdp_automation --limite 2
```

O limite afeta somente a consulta opcional das descrições. A listagem anual
completa é sempre salva no JSON e publicada em `output/gastos_scdp_2026.xlsx`.
Esse workbook tem três worksheets:

- `BASE VIAGENS` contém uma linha por PCDP. A coluna final `Código de débito` é
  recuperada da original na importação inicial e preservada pela PCDP completa.
  Viagens novas sem classificação conhecida são preenchidas manualmente;
  `Segmento` é obtido pelo código escolhido.
- `APOIO` contém o código, nome e segmento de cada curso, programa ou setor.
  Os campos de entrada são diárias e passagens distribuído, recurso total,
  transportes agendado e transportes pago. A distribuição de transportes é
  calculada como recurso total menos a distribuição de diárias e passagens.
  Valores existentes são importados da original; campos realmente ausentes
  ficam pendentes. Não repita o orçamento do grupo nas subcategorias.
  `PPGEL +` mantém o rateio da original entre PPGE, PPGEL e PPGH nas três
  colunas de percentual; os percentuais devem somar 100%.
- `RESUMO GASTOS` preserva o quadro visual da original: graduação,
  pós-graduação, administrativo, transportes e saldos. Os gastos são agrupados
  por segmento e código, com suplementos e detalhamentos no grupo correto.
  `Utilizado` de transportes corresponde ao agendado, como na fórmula original;
  o valor pago permanece disponível em APOIO. O Excel recalcula as fórmulas
  quando o workbook é aberto. Categorias adicionais e pendências ficam abaixo
  do quadro original e entram no total consolidado.

### Inclusão de categorias em APOIO

Use a próxima linha vazia da tabela de `APOIO`. Preencha código único, nome,
segmento e `Grupo no resumo`. O grupo deve ter o mesmo nome de uma linha do
resumo; assim, a categoria entra nas somas existentes sem editar fórmulas.
Preencha os valores de orçamento e transporte que forem conhecidos.

As colunas F (transportes distribuído) e M (total utilizado por categoria)
são calculadas. As colunas J:L são exclusivas do rateio `PPGEL +`.
Os gastos entram ao classificar as viagens em `BASE VIAGENS` com o novo código.
Categorias com grupos inexistentes aparecem como pendência no resumo; os gastos
continuam no total consolidado. Criar uma linha própria para um novo grupo no
quadro visual ainda exige ajustar esse quadro.

As buscas usam `VLOOKUP` (`PROCV`) com correspondência exata. Os agrupamentos
usam `SUMIF`, `SUMIFS`, `SUMPRODUCT`, `COUNTIFS` e `ISNUMBER` sobre intervalos,
sem `XLOOKUP`, matrizes dinâmicas ou referências estruturadas nas fórmulas.
O Excel antigo com essas funções e o Google Sheets oferecem as funções usadas;
a importação do arquivo específico no Sheets ainda deve ser conferida.
Os intervalos de APOIO e BASE VIAGENS usam `OFFSET` (`DESLOC`) até a
última categoria ou viagem preenchida, sem reservar 1.000 ou 10.000 linhas.
A identificação da última linha considera espaços vazios no meio dos dados.
RESUMO GASTOS soma esses intervalos dinâmicos.

Ao inserir uma linha em APOIO ou BASE VIAGENS, copie uma linha existente para
preservar as fórmulas das colunas calculadas e preencha os dados novos.
`OFFSET` amplia os intervalos de consulta e soma; o preenchimento das fórmulas
da nova linha depende dessa cópia ou do preenchimento automático da tabela.

Fontes: [SUMIFS no LibreOffice](https://help.libreoffice.org/latest/en-US/text/scalc/01/func_sumifs.html),
[VLOOKUP no Google Sheets](https://support.google.com/docs/answer/3093318?hl=en).

O gerador existente publica em `output/gastos_scdp_2026.xlsx`, substituindo a
planilha anterior com backup. A criação e atualização usam o formato original
no mesmo fluxo de extração. A referência é somente lida; não é alterada.

Duplicidades equivalentes são consolidadas por PCDP; classificações e valores
conflitantes impedem a importação. Um workbook anterior no formato genérico é
migrado na atualização quando a referência está disponível, preservando códigos
e alocações manuais. A reconstrução usa o conjunto da original; a extração
posterior pode conter viagens diferentes e exigir reconciliação.

Na primeira publicação, o arquivo é criado sem backup. Atualizações seguintes
criam um backup com data e hora antes de substituir o workbook. Os códigos
manuais continuam vinculados à PCDP completa e as alocações de `APOIO` são
preservadas entre as atualizações.

## Logs

O Loguru grava mensagens operacionais no terminal e em `logs/scdp/`. Os arquivos giram diariamente e são mantidos por 30 dias. Os logs não registram conteúdo de formulários, credenciais, cookies, corpos de requisição/resposta ou identificadores de viagem.

## Desenvolvimento

Os módulos do pacote usam imports absolutos, por exemplo `from scdp_automation.relatorio import Viagem`. O Ruff verifica essa regra e, com a extensão Ruff do VS Code, o workspace usa Ruff para formatar Python e aplicar correções de imports ao salvar. A ação de correção habilita os fixes inseguros usados pela regra `TID252`; revise as alterações sugeridas. A formatação isolada (`ruff format`) não altera imports. O ty valida tipos com `uv run ty check`; não formata código.

Execute verificações da raiz do projeto:

```sh
uv run python -m unittest discover -v
uv run ruff check .
uv run ruff format .
uv run ruff format --check .
uv run ty check
```

O ty valida tipos; ele não formata nem reescreve imports. Os arquivos em `input/`, `output/`, `logs/`, o perfil `.scdp-browser/`, `.scdp-config.toml` e `.env` são locais e ignorados pelo Git. Não publique planilhas, resultados de viagem, credenciais ou dados de sessão.
