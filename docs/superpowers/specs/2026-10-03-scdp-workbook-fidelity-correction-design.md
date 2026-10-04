# Correção da planilha de gastos SCDP

## Pedido e status

Desenho aprovado e implementado na worktree. Corrige o escopo da spec e do
plano de 2026-10-02. O usuário confirmou que o preenchimento esperado abrange
as três abas: `BASE VIAGENS`, `APOIO` e `RESUMO GASTOS`.

O resultado deve refazer a planilha original com uma entrada simplificada,
preservar o formato e a estética de `RESUMO GASTOS` e usar os dados já
existentes na original para validar a reconstrução. A referência permanece
intacta em `input/.Diárias-Pass-Transp 2026 - Consulta Saldos.xlsx`.

## Diagnóstico confirmado

A implementação está na worktree `.worktrees/xlsx-report-output`, branch
`feature/xlsx-report-output`. Não está incorporada à branch `main`.

O arquivo gerado observado tem 116 viagens, zero códigos de débito preenchidos
e zero alocações preenchidas. O código cria uma tabela de resumo de 18 colunas
com `TableStyleMedium2`; não copia o desenho original.

Evidências no código dessa worktree:

- `scdp_automation/xlsx_output.py:158`: define um novo cabeçalho de 18 colunas;
- `scdp_automation/xlsx_output.py:210`: aplica o estilo genérico azul;
- `scdp_automation/xlsx_output.py:386`: inicia todas as alocações vazias;
- `scdp_automation/xlsx_output.py:637`: só recupera códigos de um workbook
  anteriormente gerado; não importa as classificações da original.

O plano anterior, em `docs/superpowers/plans/2026-10-02-scdp-official-xlsx-update.md:16`,
exclui transportes. Sua tarefa de construção não especifica a preservação do
layout original. Os testes sintéticos desse plano não demonstram equivalência
visual nem preenchimento a partir da referência.

A leitura da aba `BD D&P` encontrou 124 ocorrências de viagens, correspondentes
a 111 PCDPs distintas. Há 12 chaves repetidas. Os valores financeiros das
ocorrências de cada chave são iguais; não há conflito entre classificações
preenchidas da mesma chave. Todas as 111 chaves têm ao menos uma ocorrência
classificada. Linhas sem classificação em uma ocorrência repetida não devem
apagar a classificação encontrada em outra ocorrência.

O conjunto da original e o conjunto do workbook gerado diferem: uma PCDP da
original está ausente do gerado e seis do gerado não estão na original. Por
isso, a reconstrução de validação deve usar o conjunto da original. A
reconciliação posterior com a extração recente deve explicitar essas diferenças.

## Resultado visual

`RESUMO GASTOS` mantém o quadro original em `B2:M58`: títulos, identificação,
blocos de graduação, pós-graduação e administrativo, totais e detalhamento de
Outros. Mantém as colunas de recurso total, distribuído/utilizado/saldo de
diárias e passagens, distribuído/utilizado/saldo de transportes e saldo geral.

Copiar da referência os estilos completos, cores, fontes, bordas, formatos
numéricos, alinhamentos, mesclagens, larguras, alturas, espaços entre blocos,
configuração de impressão e formatação condicional aplicável. Não substituir
esse quadro por uma tabela genérica. O conteúdo auxiliar à direita não deve
alargar a apresentação principal.

As fórmulas devem passar a apontar para as entradas simplificadas. A cópia de
aparência não pode deixar vínculos para abas removidas nem conservar resultados
estáticos que deixem de acompanhar a entrada.

## Entrada simplificada

### BASE VIAGENS

Uma linha por PCDP completa, preservando o sufixo complementar. Manter os campos
financeiros agregados e o código de débito como última coluna com lista suspensa.
O segmento deriva desse código. A reconstrução inicial importa valores e
classificações da original, sem exigir que o operador classifique novamente.

Deduplicar apenas ocorrências comprovadamente equivalentes nos valores
financeiros. Aproveitar a classificação preenchida quando outra ocorrência da
mesma chave estiver vazia. Conflitos entre valores ou classificações preenchidas
devem interromper a importação com diagnóstico local, sem escolher por posição.

### APOIO

Uma linha por categoria, contendo código, nome e segmento, os recursos anuais
necessários ao quadro original e os valores de transportes alimentados fora do
relatório SCDP. O input distingue recurso total, distribuição para diárias e
passagens, distribuição para transportes e utilização de transportes. Se a
referência distingue agendado e pago, manter ambos com indicação explícita de
qual alimenta o resumo, seguindo a fórmula original.

Importar os valores existentes das células que efetivamente alimentam o resumo.
Não tratar como intercambiáveis números de áreas auxiliares ou de snapshots
antigos da mesma aba. Preservar zero informado e distinguir ausência de zero.

Manter categorias adicionais para uso futuro, inclusive `PPGEL +`, os
detalhamentos de Direção e `AFAST PAÍS`. Não duplicar uma alocação de categoria
principal nas suas subcategorias. Descrições ambíguas permanecem sinalizadas,
sem inventar nomes.

### RESUMO GASTOS

Reproduzir a apresentação original com fórmulas atualizáveis. Categorias de
detalhe devem alimentar o grupo correto sem duplicar gastos. Registros que não
cabem nas linhas originais devem ficar visíveis em detalhamento no mesmo padrão
visual e entrar em totais explicitamente definidos.

Classificações realmente novas e valores não informados aparecem como pendência;
não são preenchidos por inferência. A importação da referência não deve produzir
pendências de classificação para as 111 PCDPs cuja classificação já existe.

## Validação com os dados existentes

1. Criar uma reconstrução separada a partir da original, sem substituir o arquivo
   gerado que está aberto e sem modificar a referência.
2. Confirmar 111 PCDPs únicas importadas e classificadas, com valores de origem
   rastreáveis e deduplicação das 124 ocorrências documentada.
3. Reconciliar a base simplificada com os subtotais e rodapés da original.
   Comparar também `CONSULTA ANALÍTICA`, reportando diferenças de snapshots,
   duplicidades ou tabelas dinâmicas desatualizadas; não forçar igualdade.
4. Comparar as alocações e transportes com as células que alimentam o quadro
   original, por categoria. Nenhum dado disponível pode desaparecer.
5. Recalcular uma cópia descartável em um mecanismo de planilhas. Ler os
   resultados efetivos para detectar células vazias, erros de fórmula e vínculos
   quebrados. Conferir descontos, restituições, reembolsos e cancelamentos sem
   dupla contagem.
6. Comparar o quadro reconstruído com a original: posições, mesclagens, estilos,
   dimensões, impressão e apresentação renderizada. Conferir todos os blocos.
7. Confirmar que mudanças no input recalculam utilização e saldos e que uma
   atualização preserva as entradas manuais por chave.
8. Rodar testes de regressão e os checks Ruff e ty previstos em `AGENTS.md`.

A validação deve separar discrepâncias da fonte de defeitos da reconstrução.
Um arquivo que abre e contém fórmulas não basta como evidência de aceite.

## Atualização e entrega

Manter a publicação com candidato validado, backup e substituição atômica, e o
checkpoint JSON anterior à publicação. A importação inicial é distinta da
atualização por extração: não reaplicar valores antigos sobre edições manuais.

Entregar um workbook local para revisão, com as três abas preenchidas a partir
da referência, e evidências de reconciliação. Depois, confrontar a extração
recente com esse conjunto, mantendo pendentes as classificações sem fonte.

Não commitar planilhas privadas, checkpoints, resultados ou credenciais. Não
incluir nomes, PCDPs nem registros financeiros individuais em documentação
pública. A entrega deve informar o caminho do arquivo, o que foi reconciliado e
qualquer diferença real que ainda dependa do operador.
