# Projeto de Automação de Processamento de Arquivos

Este projeto foi desenvolvido para automatizar o processamento de arquivos de diferentes formatos (.xls, .xlsx, .pdf) e consolidar os dados em uma única planilha Excel. O código está dividido em dois scripts principais, `Main.py` e `Main2.py`, que tratam de casos diferentes de arquivos de dados.

## Funcionalidades

### Main.py - Casos Grupo 1

Este script é responsável por processar arquivos que não exigem iteração linha a linha. Ele realiza as seguintes operações:

- **Conversão de Arquivos**: Converte arquivos `.xls` em `.xlsx` e arquivos `.pdf` em `.xlsx`.
- **Limpeza de Dados**: Remove linhas que contenham textos relacionados a "total", linhas em branco, e linhas de somas.
- **Seleção de Colunas**: Identifica e seleciona colunas de interesse (Loja, Marca, EAN, Valor) com base em padrões predefinidos.
- **Consolidação de Dados**: Consolida os dados de todos os arquivos processados em uma única planilha Excel.

### Main2.py - Casos Grupo 2

Este script trata de arquivos que requerem iteração nas linhas e possui funcionalidades adicionais:

- **Localização de Colunas**: Identifica a linha de cabeçalho e ajusta o DataFrame para começar a partir dela.
- **Processamento de Dados**: Segue etapas semelhantes ao `Main.py` para a limpeza, seleção de colunas e consolidação de dados.

## Estrutura de Diretórios

- **Main.py e Main2.py**: Scripts principais que realizam o processamento.
- **Pasta de entrada**: Local onde os arquivos a serem processados são armazenados.
- **Pasta de saída**: Local onde os arquivos processados e a planilha consolidada são salvos.
