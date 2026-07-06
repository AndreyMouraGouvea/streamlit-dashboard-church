# AGENTS.md — Projeto EscalaIgreja

## Visão Geral
Gerador de escala de trabalho para igreja. Duas interfaces:
- **`app.py`** — Web app (Streamlit), multi-dia/posto configurável
- **`index.py`** — Desktop app (customtkinter), fixo em Terças/Sábados com 4 postos

## Estrutura
```
app.py          — Streamlit (421+ linhas)
index.py        — customtkinter (361 linhas)
index.spec      — PyInstaller build spec para index.exe
escala.xlsx     — Planilha de entrada (colunas: Nome, Data Indisponível)
dist/           — Build output (index.exe + escalas geradas)
build/          — PyInstaller cache
```

## Dependências
`streamlit`, `pandas`, `openpyxl`, `fpdf`, `customtkinter`, `pywin32` (só index.py)

## Algoritmo de Seleção (ambos os arquivos)
- **Contador de vezes** que cada pessoa foi escalada
- A cada dia: ordena disponíveis por contagem (menor primeiro), agrupa por nível, shuffle dentro de cada grupo, pega os primeiros N
- Incrementa contagem de quem foi escalado
- Balanceamento justo + aleatório

## Convenções de Código
- Funções com `snake_case`, constantes em `UPPER_SNAKE_CASE`
- Docstrings em português (opcionais)
- Streamlit: session state gerenciado manualmente com `st.session_state`
- customtkinter: sem classes, funções soltas + closures
- Nomes de variáveis em português (ex: `postos`, `dias`, `nomes`, `indisponiveis`)

## Regras de Negócio
- Se nº pessoas < nº postos, erro
- Se não há dias no mês para os weekdays selecionados, erro
- Datas indisponíveis: formato `dd/mm/aaaa`, uma por linha (app.py) ou coluna no Excel (index.py)

## Build
- Windows: `pyinstaller index.spec` gera `dist/index.exe`
- Não há linter/typecheck configurado
