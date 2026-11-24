# GURI-DataExtraction

Script em Python para automatizar o acesso ao GURI (UNIPAMPA), extrair disciplinas do Diário de Classe e gerar uma planilha Excel com faltas registradas e faltas possíveis.

## Funcionalidades

- Login automático no GURI via Selenium/Chrome.
- Listagem das disciplinas do Diário de Classe.
- Filtro por ano e semestre/período, ou exportação completa com `--all`.
- Extração de faltas por disciplina.
- Geração de planilha `xlsx` com: ID, Nome, Carga Horária, Faltas e Faltas Possíveis.
- Execução em modo headless e logs por CLI.

## Requisitos

- Python 3.10+ (testado com Python 3.12)
- Google Chrome/Chromium instalado

## Instalação

```bash
git clone <url-do-repositorio>
cd GURI-DataExtraction-main

python3 -m venv .venv
source .venv/bin/activate  # Linux/macOS
# .venv\Scripts\activate   # Windows

pip install -r requirements.txt
```

## Configuração

Edite o `config.json` com suas credenciais e filtros:

```json
{
  "login": "SEU_USUARIO",
  "senha": "SUA_SENHA",
  "ano": "2026",
  "semestre": 1
}
```

- `login` e `senha`: credenciais do GURI.
- `ano`: ano a filtrar (ex.: `"2026"`).
- `semestre`: aceita apenas `1` ou `2`.

> Não faça commit do `config.json` com credenciais reais. Se quiser manter um modelo no repositório, use placeholders como no exemplo acima.

## Uso

Exibir ajuda:

```bash
python app.py --help
```

Extrair um ano/semestre específico:

```bash
python app.py --ano 2026 --semestre "2. Semestre"
```

Executar em headless:

```bash
python app.py --ano 2026 --semestre "2. Semestre" --headless
```

Exportar todas as disciplinas:

```bash
python app.py --all --headless
```

Definir arquivo de saída e timeout:

```bash
python app.py --ano 2026 --semestre "2. Semestre" --output resultado.xlsx --timeout 30
```

## Opções da CLI

| Opção | Descrição |
|---|---|
| `--ano` | Ano a filtrar (ex.: `2026`). Também pode ser definido no `config.json`. |
| `--semestre` | Semestre a filtrar: `1` ou `2`. Também pode ser definido no `config.json`. |
| `--all` | Ignora filtros e exporta todas as disciplinas. |
| `--output` | Arquivo Excel de saída. Padrão: `cadeiras.xlsx`. |
| `--config` | Caminho do `config.json`. Padrão: `config.json`. |
| `--headless` | Executa o Chrome em modo headless. |
| `--timeout` | Timeout padrão em segundos. Padrão: `20`. |
| `--verbose` | Ativa logs detalhados. |

## Saída

O arquivo Excel gerado contém as colunas:

- `ID`
- `Nome`
- `Carga Horária`
- `Faltas`
- `Faltas Possíveis` (fórmula: `=(C{linha} * 0.25 * 60) / 50`)

## Observações

- O GURI pode alterar a estrutura da página; se isso ocorrer, os seletores podem precisar de ajuste.
- O Selenium Manager resolve o driver do Chrome automaticamente nas versões recentes do Selenium.
- Evite executar muitas requisições em paralelo contra o GURI.

## Estrutura do projeto

```text
.
├── .gitignore
├── app.py
├── config.json
├── README.md
└── requirements.txt
```
