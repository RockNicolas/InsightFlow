# 📊 InsightFlow - Automação de Relatórios de Frota

O **InsightFlow** é uma ferramenta desenvolvida em Python para simplificar a gestão de frotas. O sistema extrai dados de utilização (Horas Trabalhadas e Quilometragem) diretamente de folhas de cálculo Excel complexas e gera relatórios profissionais em PDF, prontos para apresentação.

## 🚀 Funcionalidades

- **Extração Inteligente**: Varre folhas de cálculo ignorando erros de fórmulas e capturando dados de múltiplas frotas (Máquinas Pesadas, Camiões e Ligeiros).
- **Processamento de Dados**: Soma automaticamente as horas diárias, filtrando ruídos e valores inválidos.
- **Relatórios Customizados**: Gera PDFs com logótipo, tabelas formatadas e destaque visual (cor vermelha) para itens críticos ou de baixa produtividade.
- **Interface Web Local**: Envie a planilha Excel, escolha as abas e gere os relatórios em uma interface moderna no navegador.
- **Seleção Explícita**: O sistema não usa mais `EXCEL_FILENAME` nem `SELECTED_SHEET` no `.env`; a escolha do arquivo e das abas é feita na tela.
- **Perfis de Frota**: `Saneamento` e `Itarema` agora têm fluxos separados, com saídas organizadas por frota.

## ▶️ Como usar

1. Instale dependências Python: `pip install -r requirements.txt`
2. Tenha [Node.js](https://nodejs.org/) instalado (para o frontend)
3. Execute **`python main.py`** — o script instala/compila o React automaticamente e abre o navegador
4. Envie a planilha Excel, escolha a frota e as abas, e gere os relatórios

| Comando | O que faz |
|---------|-----------|
| `python main.py` | Compila o frontend se precisar e abre **http://127.0.0.1:5000** (tudo em um lugar) |
| `python main.py --dev` | Sobe Flask + Vite com hot reload em **http://127.0.0.1:5173** |
| `python main.py --rebuild` | Força `npm install` + `npm run build` antes de iniciar |

O `.env` agora é usado apenas para configurações da aplicação, como a pasta de saída `OUTPUT_FOLDER`.

## 🛠️ Tecnologias Utilizadas

O projeto foi construído utilizando as melhores práticas de programação modular e as seguintes bibliotecas:

- [Python 3.13](https://www.python.org/): Linguagem base do projeto.
- [Pandas](https://pandas.pydata.org/): Manipulação e análise de dados de alto desempenho.
- [FPDF](http://www.fpdf.org/): Geração de documentos PDF de forma programática.
- [Python-Dotenv](https://pypi.org/project/python-dotenv/): Gestão de configurações e segurança de ambiente.
- [Openpyxl](https://openpyxl.readthedocs.io/): Motor de leitura para ficheiros Excel (.xlsx).
- [React](https://react.dev/) + [TypeScript](https://www.typescriptlang.org/): Interface web (pasta `frontend/`).


## 📁 Estrutura do Projeto

```text
InsightFlow/
│
├── main.py              # Inicia a interface web local
├── web_app.py           # API Flask + serve o build React
├── frontend/            # Interface React + TypeScript (Vite)
├── app/
│   └── web/
│       └── server.py    # Ponto de entrada organizado da camada web
├── .env                 # Configurações opcionais da aplicação
├── assets/
        company/
        logo.png          # Logótipo da empresa para o relatório
├── modules/
│   ├── processors/       # Processamento e leitura de dados
│   ├── reports/          # Geração dos relatórios e PDFs
│   ├── fleet/            # Perfis, roteamento e organização por frota
│   └── __init__.py       # Pacote principal dos módulos
└── outputs/             # Relatórios gerados, separados por frota
```

AUTOR: Nicolas Rock

# 📊 InsightFlow - Fleet Report Automation

**InsightFlow** is a tool developed in Python to simplify fleet management. The system extracts usage data (Hours Worked and Mileage) directly from complex Excel spreadsheets and generates professional PDF reports, ready for presentation.

## 🚀 Features

- **Intelligent Extraction**: Scans spreadsheets ignoring formula errors and capturing data from multiple fleets (Heavy Machinery, Trucks, and Light Vehicles).

- **Data Processing**: Automatically sums daily hours, filtering out noise and invalid values.

- **Customized Reports**: Generates PDFs with logos, formatted tables, and visual highlighting (red color) for critical or low-productivity items.

- **Local Web Interface**: Upload the Excel file, choose the sheets, and generate reports in a modern browser interface.
- **Explicit Selection**: The app no longer uses `EXCEL_FILENAME` or `SELECTED_SHEET` in `.env`; file and sheet selection now happens in the UI.
- **Fleet Profiles**: `Saneamento` and `Itarema` now have separate processing flows, with outputs organized by fleet.

## ▶️ How to use

1. Install Python deps: `pip install -r requirements.txt`
2. Build the frontend (first time): `cd frontend && npm install && npm run build`
3. Run `python main.py` and open `http://127.0.0.1:5000`
4. Upload the Excel workbook, choose fleet and sheets, and generate reports

**Frontend dev:** run `python main.py` in one terminal and `cd frontend && npm run dev` in another (`http://localhost:5173`).

The `.env` file is now used only for application settings, such as the output folder `OUTPUT_FOLDER`.

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
├── main.py # Starts the local web interface
├── web_app.py # Main web server
├── app/
│   └── web/
│       └── server.py # Organized entrypoint for the web layer
├── .env # Optional application settings
├── assets/
    company/
    logo.png # Company logo for the report
├── modules/
│ ├── processors/ # Data processing and Excel readers
│ ├── reports/ # Report and PDF generation
│ ├── fleet/ # Fleet profiles, routing, and fleet packages
│ └── __init__.py # Main modules package
├── frontend/ # React + TypeScript web interface (Vite)
└── outputs/ # Generated reports, split by fleet
```

AUTHOR: Nicolas Rock