# Leitura de Planilha de Transportadora e Processamento de Viagens

Este repositório contém um código de automação para ler dados de uma planilha Excel, exibir mensagens personalizadas para cada linha e executar um fluxo externo para processar informações de viagem.

## Funcionalidades

* Leitura de dados de uma planilha Excel.
* Exibição de mensagens personalizadas com nome e origem da viagem.
* Execução de um fluxo externo para processar dados da viagem (origem, destino, transportadora).
* Tratamento de erros durante a execução do fluxo externo.

## Estrutura do Código

1.  **Leitura da Planilha:**
    * A função "LerPlanilha" lê os dados de uma planilha Excel e os armazena na variável `ExcelData`.

2.  **Loop de Processamento das Linhas:**
    * Um loop `FOREACH` itera sobre cada linha em `ExcelData`.
    * A variável `LinhaAtual` representa a linha atual sendo processada.

3.  **Exibição de Mensagem Personalizada:**
    * Uma caixa de diálogo exibe uma mensagem com o nome e a origem da viagem, formatada dinamicamente com os dados da linha atual.

4.  **Execução do Fluxo Externo:**
    * O fluxo externo "AulaFilaDeTrabalhoTransportadora" (ID: `d2f2af84-0ff7-45a7-a896-4a78e175d728`) é executado.
    * Os dados da viagem (origem, destino, transportadora e prioridade) são passados como parâmetros para o fluxo externo.
    * A execução aguarda a conclusão do fluxo externo antes de continuar.

5.  **Tratamento de Erros:**
    * Um bloco `ON ERROR` trata erros que possam ocorrer durante a execução do fluxo externo.
    * Em caso de erro, a função "SeErro" é chamada e o erro é relançado.

