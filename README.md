# SCDP

Sistema para extrair o relatório **Relatórios > Viagem** do SCDP e atualizar a planilha de gastos. A automação usa uma janela visível do Google Chrome; o operador resolve o CAPTCHA e outros desafios de autenticação.

## Preparação

Instale o [uv](https://docs.astral.sh/uv/) e o Google Chrome. Execute os comandos na raiz do projeto. O projeto requer Python 3.14.5 ou superior, dentro da série 3.14.

Antes de extrair, coloque o template final obrigatório em `input/gastos_scdp_template.xlsx`. Ele deve declarar `SCDPLayoutVersion="9"` e conter, nesta ordem, `BASE VIAGENS`, `APOIO` e `RESUMO GASTOS`. O template é somente leitura; arquivos ausentes, inválidos ou de layouts antigos são recusados, inclusive quando já existe uma saída. Não há conversão de layouts nem geração de uma planilha alternativa.

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

A extração consulta sempre o ano atual, considerando o fuso de São Paulo. Ao iniciar, o sistema atualiza automaticamente o campo `ano` no início de `.scdp-config.toml`, antes da seção `[chrome_profile]`. Exemplo para 2026:

```toml
ano = 2026
```

O nome da planilha segue sempre `output/gastos_scdp_{ano}.xlsx`, e o JSON segue `output/viagens_scdp_{ano}.json`. Na virada do ano, o sistema parte do template para o novo ano, sem copiar viagens ou classificações do ano anterior. O par anterior é arquivado e sai da raiz somente depois da publicação bem-sucedida. A seleção do perfil Chrome é preservada.

Template, JSON, planilha, logs, credenciais e configuração são resolvidos a partir da raiz do projeto/worktree ativo, independentemente do diretório de execução. Execute `uv run --project /caminho/para/scdp python -m scdp_automation` quando estiver fora da raiz. Worktrees usam seus próprios arquivos locais.

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

Cada linha representa uma PCDP. A coluna `Descrição do pedido` mostra o motivo detalhado extraído do SCDP e fica vazia enquanto essa informação não estiver disponível. Confira os valores extraídos e use as últimas três colunas:

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

Se a planilha não existir, o sistema copia o template final obrigatório, limpa os dados de `BASE VIAGENS` e escreve somente as viagens coletadas. Nas atualizações seguintes, a saída existente fornece os preenchimentos manuais. A automação preserva valores, fórmulas, estilos, tabelas, validações, nomes, mesclagens, dimensões, tema e impressão de `APOIO` e `RESUMO GASTOS` que o openpyxl suporta. As exceções são `RESUMO GASTOS!M2`, gravada como data Excel com formato `dd/mm/yyyy`, as flags de recálculo e os dados/controles de BASE. O salvamento pode reorganizar o XML interno do XLSX.

A raiz de `output/` contém a planilha e o JSON atuais. Cada execução arquiva o par anterior uma única vez em `output/backup/`, com data e hora no nome. Após a publicação bem-sucedida, mantém os quatro pares anteriores de todos os anos conjuntamente, totalizando cinco execuções incluindo a atual. Planilha e JSON de cada histórico compartilham o identificador da execução; arquivos manuais e órfãos não são apagados nem completados artificialmente. A célula `RESUMO GASTOS!M2` recebe a data da atualização dos valores, no fuso de São Paulo. Antes de substituir a planilha, o sistema valida a atualização. Se a nova listagem não contiver alguma PCDP já publicada, a atualização é interrompida para permitir a conferência, preservando a planilha anterior.

A execução mantém um lock por diretório de saída desde antes de carregar o checkpoint e acessar o navegador até terminar a extração. Duas execuções no mesmo destino não misturam checkpoints ou históricos. As consultas individuais e a segunda publicação da mesma execução não criam outro histórico.

Se o Excel bloquear a substituição, a planilha anterior permanece e a candidata validada fica em `output/backup/recovery.<id>.xlsx`; o erro informa o caminho. Se esse deslocamento também falhar, o erro indica onde a candidata sobreviveu. Feche o arquivo no Excel e execute novamente. O JSON coletado continua disponível para retomada ou reconciliação; falhas não provocam poda dos históricos. Candidatas de recuperação em backup podem ser removidas após uma publicação posterior bem-sucedida.

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

### Erros de planilha

A mensagem identifica o caminho e distingue o template obrigatório da saída publicada. O layout aceito é somente a versão 9; não altere o marcador de versão para contornar a validação. Se o template estiver ausente ou inválido, restaure `input/gastos_scdp_template.xlsx` com um template final válido.

`git pull` e `git reset` não atualizam os resultados locais ignorados pelo Git. Se uma saída antiga tiver layout incompatível e você quiser começar novamente, preserve a pasta `output` inteira, incluindo JSON e históricos, fora da saída ativa antes de executar de novo. Na raiz do projeto, o PowerShell pode arquivá-la sem apagar os dados:

```powershell
Move-Item -LiteralPath .\output -Destination (Join-Path . ("output-arquivado-" + (Get-Date -Format "yyyyMMdd-HHmmss-fffffff")))
```

Confira os preenchimentos manuais no arquivo arquivado antes de preencher a nova saída. Para erros de acesso, feche a planilha no Excel e confira as permissões do caminho indicado. Uma falha de publicação informa onde a candidata validada foi preservada para recuperação.
