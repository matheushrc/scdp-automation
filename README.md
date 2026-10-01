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

O ty valida tipos; ele não formata nem reescreve imports. Os arquivos em `input/`, `output/`, `logs/`, o perfil `.scdp-browser/` e `.env` são locais e ignorados pelo Git. Não publique planilhas, resultados de viagem, credenciais ou dados de sessão.
