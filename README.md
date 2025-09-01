# Gerador de Relatório de Acesso

Este projeto tem como objetivo **ler bases de dados em Excel de
funcionários e visitantes** e gerar um relatório consolidado em formato
**Excel (.xlsx)**.\
A aplicação possui uma interface gráfica simples utilizando
**PySimpleGUI**, permitindo que o usuário selecione arquivos de entrada,
configure empresas e exporte o relatório final.

------------------------------------------------------------------------

## 🚀 Funcionalidades

-   Leitura da base de **funcionários** (planilha Excel).
-   Leitura da base de **visitantes** (planilha Excel).
-   Consolidação dos dados em um único relatório:
    -   Quantidade de pessoas por dia.
    -   Classificação por empresa.
    -   Contagem de visitantes vinculados ou não a empresas.
    -   Total diário de pessoas e percentual de ocupação.
-   Exportação para um arquivo `relatorio_acesso.xlsx`.
-   Interface gráfica amigável para:
    -   Seleção dos arquivos de entrada.
    -   Escolha do mês e ano.
    -   Configuração da ordem das empresas (adicionar, remover e
        reordenar).
    -   Salvamento das configurações em `config.json`.

------------------------------------------------------------------------

## 📦 Requisitos

Certifique-se de ter o **Python 3.9+** instalado.\
As bibliotecas necessárias estão listadas abaixo:

``` bash
pip install PySimpleGUI openpyxl
```

------------------------------------------------------------------------

## 📂 Estrutura do Projeto

    Tratadados.py         # Código principal da aplicação
    config.json           # Arquivo gerado automaticamente com a ordem das empresas
    relatorio_acesso.xlsx # Relatório de saída (gerado pela aplicação)

------------------------------------------------------------------------

## ▶️ Como Executar

No terminal, execute:

``` bash
python Tratadados.py
```

A janela principal será aberta. Nela você poderá: 1. Selecionar o
arquivo de funcionários (Excel). 2. Selecionar o arquivo de visitantes
(Excel). 3. Informar o mês e ano desejados. 4. Clicar em **Gerar
Relatório**.

O relatório será salvo automaticamente como **`relatorio_acesso.xlsx`**
na pasta do programa.

------------------------------------------------------------------------

## ⚙️ Configurações Avançadas

Na aba **Configurações**, você pode: - Adicionar novas empresas. -
Remover empresas. - Reordenar a lista de empresas. - Salvar suas
alterações (ficarão registradas em `config.json`).

------------------------------------------------------------------------

## 📊 Exemplo de Saída

O relatório gerado possui colunas como:

-   **Dia**
-   **Empresas cadastradas**
-   **Visitante WTNU**
-   **Total de pessoas**
-   **% de ocupação**

Além disso, ao final da planilha, é incluída a linha **ENTRADA**, que
mostra a soma de todos os acessos por empresa.

------------------------------------------------------------------------

## 🛠️ Tecnologias Utilizadas

-   [Python](https://www.python.org/)\
-   [PySimpleGUI](https://pysimplegui.readthedocs.io/en/latest/)\
-   [OpenPyXL](https://openpyxl.readthedocs.io/en/stable/)

------------------------------------------------------------------------

## 📌 Observações

-   Apenas arquivos **Excel (.CSV)** são aceitos como entrada.\
-   O cálculo de **% de ocupação** é baseado em um limite fixo de
    **8.931 pessoas** (valor definido no código).\
-   Caso o usuário não selecione os arquivos corretamente, uma mensagem
    de erro será exibida.
