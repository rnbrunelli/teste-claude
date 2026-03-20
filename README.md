# BYM — Gerador de Relatório Gerencial Mensal

Ferramenta para geração de relatórios gerenciais mensais de obras em formato Excel (.xlsx). Os dados são inseridos via interface web (Streamlit) ou fornecidos como JSON pela linha de comando.

---

## Instalação

```bash
pip install -r requirements.txt
```

**Dependências:** `streamlit >= 1.32.0`, `openpyxl >= 3.1.2`, `pandas >= 2.0.0`

---

## Como usar

### Interface Web (recomendado)

```bash
streamlit run app.py
```

Acesse `http://localhost:8501`. Preencha os dados nas 10 abas, clique em **Gerar Relatório Excel** e baixe o arquivo `.xlsx`.

### Linha de comando

```bash
# Usa dados_relatorio.json e gera arquivo com timestamp
python main.py

# Caminhos personalizados
python main.py --dados meu_projeto.json --saida RelatorioMarco26.xlsx
```

---

## Estrutura de dados (JSON)

O arquivo `dados_relatorio.json` é o schema de referência. Todas as chaves esperadas pelo gerador estão definidas nele com valores padrão vazios.

| Chave raiz | Conteúdo |
|---|---|
| `projeto` | Identificação do empreendimento (nome, incorporadora, áreas, estrutura) |
| `relatorio` | Período de referência (mês, ano, data do relatório) |
| `controle_prazo` | Datas contratuais vs. tendência e marcos do cronograma |
| `avanco_fisico` | Percentuais de avanço, IPF e quantidade de serviços do mês |
| `farol_metas` | Lista de metas do período com status por farol (1=Vermelho, 2=Amarelo, 3=Verde) |
| `metas_proximo_mes` | Lista de metas planejadas para o próximo período |
| `histograma_mao_obra` | Headcount por empresa e semana do mês |
| `fluxo_caixa` | Previsto vs. realizado mensal com INCC |
| `tabela_aporte` | Aportes acumulados previsto vs. realizado |
| `analise_financeira` | Orçamento por centro de custo (grupos e subitens) |
| `gerenciamento_contratacoes` | Cronograma de contratação de serviços |
| `controle_mapas_contratacoes` | Valor contratado vs. orçamento por fornecedor |
| `cronograma_suprimentos` | Fases de suprimento (carta convite → contratação) |
| `legalizacao` | Documentos iniciais e datas do Habite-se |
| `datas_marco_prototipo` | Marcos de execução do protótipo por local |

Datas são armazenadas como string `"YYYY-MM-DD"`.

---

## Abas do relatório Excel gerado

| Aba | Seção |
|---|---|
| CAPA | Identificação visual do projeto |
| Sumario | Índice de todas as seções |
| 1-4) Resumo e Indicadores | Dados do empreendimento, controle de prazo e avanço físico |
| 3.2) Fluxo de Caixa | Previsto vs. realizado mês a mês |
| 3.3) Analise Financeira | Orçamento por centro de custo |
| 5-6) Prototipo e Histograma | Marcos do protótipo e headcount de mão de obra |
| 8) Farol de Metas | Lista de metas do período com indicadores de farol |
| 9) Metas Proximo Mes | Metas planejadas para o próximo período |
| 10) Legalizacao | Documentos e cronograma do Habite-se |
| 13) Tabela de Aporte | Aportes previsto vs. realizado acumulados |
| 14) Gerenc. Contratacoes | Cronograma de contratações com farol |
| 15) Mapas de Contratacoes | Controle financeiro de contratos |
| 16) Cronograma de Suprimentos | Fases de suprimento por serviço |

---

## Arquivos

| Arquivo | Função |
|---|---|
| `app.py` | Interface Streamlit com 10 abas de entrada de dados |
| `gerador_relatorio.py` | Classe `GeradorRelatorio` — gera o workbook Excel |
| `main.py` | Entrada CLI — lê JSON e salva `.xlsx` |
| `dados_relatorio.json` | Schema com todas as chaves e valores padrão |
| `.streamlit/config.toml` | Tema visual da interface (cores BYM) |

---

## Salvamento e carregamento de dados

- **Salvar JSON:** clique em **Salvar JSON** na interface para exportar o estado atual como `dados_relatorio.json`.
- **Carregar JSON:** use o uploader na barra lateral para importar um JSON salvo anteriormente e retomar o preenchimento.

---

## Paleta de cores

| Nome | Hex | Uso |
|---|---|---|
| Azul escuro | `#1F3864` | Cabeçalhos principais, capa |
| Azul médio | `#2E75B6` | Cabeçalhos de seção |
| Azul claro | `#BDD7EE` | Linhas de destaque, totais |
| Laranja | `#ED7D31` | Período no cabeçalho, realizado |
| Verde | `#70AD47` / `#C6EFCE` | Farol verde / fundo OK |
| Amarelo | `#FFD966` | Farol amarelo (atenção) |
| Vermelho | `#FF4B4B` / `#FFCCCC` | Farol vermelho / fundo desvio negativo |
