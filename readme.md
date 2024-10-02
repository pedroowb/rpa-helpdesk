# Automação RPA para Abertura de Chamados

Este projeto é um protótipo de uma automação RPA (Robotic Process Automation) feita em Python, desenvolvida para agilizar o processo de abertura de chamados com base nos dados extraídos de relatórios em planilhas de Helpdesk. As informações da empresa foram removidas para este exemplo, apresentando uma versão simplificada para fins de exemplificação.

## Funcionalidades

- **Leitura de planilhas**: A automação lê as planilhas de relatórios do Helpdesk utilizando a biblioteca `openpyxl`, extraindo informações relevantes para a criação de chamados.
- **Interação com a interface**: Utilizando `pyautogui`, a automação simula ações de teclado e mouse para abrir chamados automaticamente em uma interface gráfica de sistema de Helpdesk.
- **Comparação e otimização de dados**: A biblioteca `fuzzywuzzy` é utilizada para realizar comparações de texto e garantir que as informações de chamados existentes estejam alinhadas corretamente.
- **Gerenciamento de arquivos**: A biblioteca `OS` é utilizada para manipular diretórios e gerenciar arquivos de maneira eficiente dentro do projeto.

## Bibliotecas Utilizadas

- `openpyxl`: Para manipulação de planilhas no formato `.xlsx`.
- `pyautogui`: Para automação das interações com o sistema, simulando cliques e entradas de teclado.
- `fuzzywuzzy`: Para comparação e tratamento de strings, identificando semelhanças entre textos, como nomes de clientes ou descrições de problemas.
- `OS`: Para lidar com arquivos e diretórios do sistema operacional.

## Pré-requisitos

Antes de rodar o script, certifique-se de ter instaladas as bibliotecas necessárias:

```bash
pip install openpyxl pyautogui fuzzywuzzy python-Levenshtein
```

## Como Funciona

1. **Carregamento dos Dados**: A automação carrega o relatório da planilha de Helpdesk, contendo informações como nome do cliente, número do chamado e descrição do problema.
2. **Abertura Automática de Chamados**: O script, utilizando `pyautogui`, acessa a interface gráfica do sistema de chamados e simula a abertura de novos tickets com base nas informações fornecidas na planilha.
3. **Tratamento de Informações**: Se houver inconsistências nas informações dos chamados, a automação utiliza `fuzzywuzzy` para comparar e corrigir os dados, garantindo maior precisão.
4. **Logs e Feedback**: O sistema gera logs de cada chamado aberto para acompanhamento posterior.

## Execução

Para rodar o script, basta executar o seguinte comando no terminal:

```bash
python scripts/automacao_abertura_chamados.py
```

Certifique-se de que as janelas do sistema de Helpdesk estejam abertas e corretamente posicionadas para que a automação funcione conforme esperado.

## Limitações

Este protótipo foi desenvolvido como uma versão simplificada para fins de demonstração. O projeto original foi limpo de todas as informações confidenciais e personalizações específicas da empresa.
