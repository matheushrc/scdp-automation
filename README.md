# SCDP

Sistema para extrair o relatório **Relatórios > Viagem** do SCDP e atualizar a planilha de gastos. A automação usa uma janela visível do Google Chrome; o operador resolve o CAPTCHA e outros desafios de autenticação.

## Preparação

Instale o [uv](https://docs.astral.sh/uv/) e o Google Chrome. Execute os comandos na raiz do projeto. O projeto requer Python 3.14.5 ou superior, dentro da série 3.14.

Instale as dependências:

```sh
uv sync
```

Crie o arquivo de credenciais a partir do exemplo. Em Linux, restrinja as permissões do arquivo:

```sh
cp .env.example .env
chmod 600 .env
```

Preencha o CPF e a senha do gov.br em `.env`:

```dotenv
USERNAME=seu-cpf
PASSWORD=sua-senha
```

As variáveis `USERNAME` e `PASSWORD` do ambiente também são aceitas quando a respectiva chave não estiver presente em `.env`.

Na primeira execução, escolha um perfil do Chrome com as setas e Enter; Escape cancela. Feche todas as janelas do Chrome para permitir a cópia do perfil. O sistema cria uma cópia local sem extensões em `.scdp-browser/` e salva a seleção em `.scdp-config.toml`. Nas próximas execuções, essa cópia é reutilizada.

## Comandos da CLI

### Extrair viagens e atualizar a planilha

```sh
uv run python -m scdp_automation
```

O comando abre ou conecta ao Chrome local, aguarda a autenticação e consulta **RELATÓRIOS > Viagem** para CCH — Campus Chapecó/SC, código `121766`, com a opção **Todas as viagens do ano de exercício**. Se necessário, o sistema clica em **Entrar com gov.br** e preenche as credenciais. Resolva os desafios de autenticação na janela visível e mantenha o Chrome aberto durante a coleta.

A execução salva a listagem completa, atualiza a planilha, consulta os motivos das solicitações que precisam de verificação e atualiza a planilha novamente ao terminar.

### Limitar as consultas individuais

```sh
uv run python -m scdp_automation --limite 2
```

O limite afeta somente a quantidade de solicitações consultadas individualmente para obter o motivo da viagem. A listagem anual completa continua sendo coletada e publicada. Use `--limite 0`, ou omita a opção, para consultar todas as solicitações que precisam de verificação.

### Autenticar sem extrair dados

```sh
uv run python -m scdp_automation --login
```

O comando termina depois de retornar ao SCDP autenticado. Ele não consulta relatórios nem atualiza o JSON ou a planilha.

### Abrir somente o navegador

```sh
uv run python -m scdp_automation --abrir-navegador
```

Mantenha esse comando em execução enquanto usar a janela. Para iniciar a extração usando o mesmo navegador, execute o comando de coleta em outro terminal.

### Trocar o perfil do Chrome

```sh
uv run python -m scdp_automation --selecionar-perfil-chrome --login
```

Feche todas as janelas do Chrome para permitir a cópia do perfil escolhido. A opção `--selecionar-perfil-chrome` pode ser combinada com `--login` ou `--abrir-navegador`; usada sozinha, seleciona o perfil e inicia a extração.

### Consultar a ajuda

```sh
uv run python -m scdp_automation --help
```

Após instalar as dependências, também é possível usar `uv run scdp-extrair` no lugar de `uv run python -m scdp_automation`, com as mesmas opções.

## Resultados e retomada

- `output/viagens_scdp_2026.json`: viagens, trechos, valores, descrição do motivo e data da última verificação.
- `output/gastos_scdp_2026.xlsx`: planilha de gastos com as abas `BASE VIAGENS`, `APOIO` e `RESUMO GASTOS`.
- `logs/scdp/`: registros de execução, com rotação diária e retenção de 30 dias.

O JSON é salvo depois da coleta da listagem e após cada consulta individual concluída. Se a execução for interrompida, execute novamente para aproveitar o progresso salvo. A descrição `null` indica uma consulta ainda não concluída; uma descrição vazia indica que o campo foi consultado e estava vazio.

Uma solicitação é consultada individualmente quando a viagem começa depois de hoje, quando falta a descrição ou quando falta a data da última verificação. Viagens cujo início já ocorreu, inclusive hoje, deixam de ser consultadas individualmente quando a descrição e a data da verificação estão salvas. A data considerada é a de início, e a comparação usa o fuso de São Paulo. A listagem geral continua sendo atualizada em todas as execuções.

## Uso da planilha

Abra `output/gastos_scdp_2026.xlsx` para conferir e preencher os dados. As fórmulas são configuradas para recalcular ao abrir o arquivo no Excel. Feche a planilha antes de executar a extração para permitir a substituição do arquivo.

### BASE VIAGENS

Cada linha representa uma PCDP. Confira os valores extraídos e use as últimas três colunas:

- `Segmento`: calculado a partir do código de débito cadastrado em `APOIO`.
- `Código de débito`: escolha a categoria responsável pela despesa na lista suspensa.
- `Descontar do curso?`: para viagens canceladas, selecione `Sim` para incluir o valor nos gastos ou `Não` para excluí-lo. Para as demais situações, o gasto é incluído independentemente dessa escolha.

Uma viagem cancelada sem decisão em `Descontar do curso?` deixa os cálculos correspondentes como `Pendente` e destaca a linha em vermelho. Uma linha com PCDP preenchida e segmento vazio fica amarela; se as duas condições ocorrerem, prevalece o vermelho.

As colunas de início da viagem, término da viagem e última verificação ficam ocultas. As demais colunas usam texto ajustado à largura e alinhamento centralizado.

### APOIO

Cadastre os códigos de débito, os nomes e os segmentos das categorias. `PPGH/PPGDH` é um único código para Mestrado e Doutorado em História.

Preencha os valores conhecidos nas colunas `Diárias e passagens distribuído (R$)`, `Transportes distribuído (R$)`, `Transportes agendado (R$)` e `Transportes pago (R$)`. Informe `0` quando o valor for conhecido e igual a zero; deixe vazio apenas o que ainda não foi informado. O recurso total, na coluna F, é calculado pela soma das duas distribuições. O total utilizado por categoria, na coluna L, também é calculado.

As colunas I a K contêm o rateio de `PPGEL +` entre PPGE, PPGEL e História. Preencha percentuais que somem 100% quando essa categoria tiver valores ou gastos. Não repita o orçamento de uma categoria principal nas suas subcategorias.

Para incluir uma categoria, copie uma linha existente da tabela, mantenha as fórmulas e preencha um código único, nome, segmento e os valores disponíveis. Depois, classifique as viagens com esse código em `BASE VIAGENS`.

### RESUMO GASTOS

Confira os recursos distribuídos, os gastos, os transportes e os saldos por categoria. O utilizado de transportes corresponde ao valor agendado; o valor pago permanece disponível em `APOIO`. As subcategorias com código iniciado por `DIREÇÃO -` alimentam o total de Direção, e `PPGEL +` é distribuído conforme os percentuais informados em `APOIO`.

Não existe uma coluna `Grupo no resumo`. As linhas do resumo identificam a categoria pelo nome por extenso ou pelo código de débito cadastrado em `APOIO`. Para apresentar uma categoria nova, copie uma linha de categoria do resumo, mantenha as fórmulas e troque o nome pelo nome ou código correspondente. Categorias sem linha no resumo são sinalizadas como pendência.

Ao adicionar linhas em `BASE VIAGENS` ou `APOIO`, copie uma linha existente para preservar as fórmulas calculadas. Os intervalos de consulta e soma acompanham a última PCDP ou categoria preenchida, mesmo com linhas vazias entre os dados.

### Atualizações e backups

A extração usa a planilha existente para atualizar `BASE VIAGENS`, preservando os códigos de débito e as decisões de desconto pela PCDP completa, além das entradas de `APOIO` e dos ajustes de `RESUMO GASTOS`. Não é necessário executar um criador de planilha separadamente.

Se a planilha não existir, o sistema cria uma nova. Quando disponível, a referência `input/.Diárias-Pass-Transp 2026 - Consulta Saldos.xlsx` fornece o formato original e os dados iniciais; o arquivo de referência é somente lido. Mantenha a planilha ajustada no caminho de saída para que ela seja usada nas próximas atualizações.

Antes de substituir uma planilha existente, o sistema valida a atualização e cria um backup com data e hora em `output/`. Se a nova listagem não contiver alguma PCDP já publicada, a atualização é interrompida para permitir a conferência, preservando a planilha anterior.

## Desenvolvimento

Execute as verificações na raiz do projeto:

```sh
uv run python -m unittest discover -v
uv run ruff check .
uv run ruff format --check .
uv run ty check
```

Para aplicar a formatação Python, use `uv run ruff format .`. Após alterar dependências, execute `uv sync`.

Os arquivos em `input/`, `output/` e `logs/`, as cópias de perfil `.scdp-browser*`, `.scdp-config.toml` e `.env` são locais e ignorados pelo Git. Não publique credenciais, perfis de navegador, resultados ou planilhas privadas.
