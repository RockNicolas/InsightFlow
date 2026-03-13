# 📊 InsightFlow - Automação de Relatórios de Frota

O **InsightFlow** é uma ferramenta desenvolvida em Python para simplificar a gestão de frotas. O sistema extrai dados de utilização (Horas Trabalhadas e Quilometragem) diretamente de folhas de cálculo Excel complexas e gera relatórios profissionais em PDF, prontos para apresentação.

## 🚀 Funcionalidades

- **Extração Inteligente**: Varre folhas de cálculo ignorando erros de fórmulas e capturando dados de múltiplas frotas (Máquinas Pesadas, Camiões e Ligeiros).
- **Processamento de Dados**: Soma automaticamente as horas diárias, filtrando ruídos e valores inválidos.
- **Relatórios Customizados**: Gera PDFs com logótipo, tabelas formatadas e destaque visual (cor vermelha) para itens críticos ou de baixa produtividade.
- **Configuração Simples**: Gestão de variáveis (ficheiros, abas e caminhos) via arquivo `.env`.

## 🛠️ Tecnologias Utilizadas

O projeto foi construído utilizando as melhores práticas de programação modular e as seguintes bibliotecas:

- [Python 3.13](https://www.python.org/): Linguagem base do projeto.
- [Pandas](https://pandas.pydata.org/): Manipulação e análise de dados de alto desempenho.
- [FPDF](http://www.fpdf.org/): Geração de documentos PDF de forma programática.
- [Python-Dotenv](https://pypi.org/project/python-dotenv/): Gestão de configurações e segurança de ambiente.
- [Openpyxl](https://openpyxl.readthedocs.io/): Motor de leitura para ficheiros Excel (.xlsx).


## 📁 Estrutura do Projeto

```text
InsightFlow/
│
├── main.py              # Orquestrador principal do sistema
├── .env                 # Configurações de caminhos e nomes de ficheiros
├── assets/
        company/
        logo.png          # Logótipo da empresa para o relatório
├── modules/
│   ├── data_processor.py # Lógica de extração e limpeza de dados
│   └── pdf_generator.py  # Design e geração do layout do PDF
    └── observation_report.py  # Design e geração do layout do relatorio de anotações
├── inputs/              # Pasta para colocar as folhas de cálculo Excel
└── outputs/             # Pasta onde os relatórios gerados são guardados
```

AUTOR: Nicolas Rock

# 📊 InsightFlow - Fleet Report Automation

**InsightFlow** is a tool developed in Python to simplify fleet management. The system extracts usage data (Hours Worked and Mileage) directly from complex Excel spreadsheets and generates professional PDF reports, ready for presentation.

## 🚀 Features

- **Intelligent Extraction**: Scans spreadsheets ignoring formula errors and capturing data from multiple fleets (Heavy Machinery, Trucks, and Light Vehicles).

- **Data Processing**: Automatically sums daily hours, filtering out noise and invalid values.

- **Customized Reports**: Generates PDFs with logos, formatted tables, and visual highlighting (red color) for critical or low-productivity items.

- **Simple Configuration**: Management of variables (files, tabs, and paths) via `.env` file.

## 🛠️ Technologies Used

The project was built using best practices in modular programming and the following libraries:

- [Python 3.13](https://www.python.org/): Base language of the project.

- [Pandas](https://pandas.pydata.org/): High-performance data manipulation and analysis.

- [FPDF](http://www.fpdf.org/): Programmatic generation of PDF documents.

- [Python-Dotenv](https://pypi.org/project/python-dotenv/): Environment configuration and security management.

- [Openpyxl](https://openpyxl.readthedocs.io/): Reading engine for Excel files (.xlsx).

## 📁 Project Structure

```text
InsightFlow/
│
├── main.py # Main system orchestrator
├── .env # File path and name configurations
├── assets/
    company/
    logo.png # Company logo for the report
├── modules/
│ ├── data_processor.py # Data extraction and cleaning logic
│ └── pdf_generator.py # PDF layout design and generation
├── inputs/ # Folder to place Excel spreadsheets
└── outputs/ # Folder where generated reports are saved
```

AUTHOR: Nicolas Rock