# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

BYM Relatório Gerencial is a Python application that generates monthly construction project management reports as Excel workbooks. The UI is built with Streamlit, and Excel generation uses openpyxl.

## Commands

```bash
# Run the Streamlit web app (from relatorio_obras/)
streamlit run app.py

# Generate report from CLI
python main.py
python main.py --dados custom_data.json --saida output_report.xlsx
```

Dependencies: `pip install -r requirements.txt`

## Architecture

Three-layer pattern:

1. **[app.py](app.py)** — Streamlit UI with 10 tabs for data entry. Stores all form state in `st.session_state.dados`. On submit, instantiates `GeradorRelatorio` and offers the resulting bytes as an `.xlsx` download.

2. **[gerador_relatorio.py](gerador_relatorio.py)** — Core Excel generator. `GeradorRelatorio` class receives the data dict and builds a multi-sheet workbook (13 worksheets: cover page, summary, and 11 detailed sheets). Corporate color palette: dark blue `#1F3864`, orange `#ED7D31`.

3. **[main.py](main.py)** — Thin CLI wrapper around `GeradorRelatorio`. Reads JSON from `--dados`, writes `.xlsx` to `--saida`.

**Data flow:** Streamlit form → `st.session_state.dados` dict → `GeradorRelatorio(dados)` → Excel bytes → browser download (or file via CLI).

**[dados_relatorio.json](dados_relatorio.json)** defines the canonical data schema with all expected keys and default empty values. Use it as the reference when adding new data fields.

## Key Conventions

- All UI labels, form fields, and report text are in Portuguese.
- Date conversion helpers `_str_para_date()` / `_date_para_str()` are used throughout both files — keep date handling consistent with these.
- `st.data_editor` (pandas DataFrames) is used for tabular inputs; changes must be synced back to `st.session_state.dados` before calling the generator.
