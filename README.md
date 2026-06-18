# Padronizador de Base Diária de Credenciados
**PUCPR / Grupo Marista**

Script Python que converte automaticamente a extração diária do sistema de controle de acesso (`.csv`) em uma planilha Excel padronizada com o guia de estilo visual da PUCPR.

---

## Estrutura do projeto

```
dashboard - monitoramento check-in/
├── bases/
│   ├── 17_06_2026.csv       ← base exportada do sistema
│   └── 17_06_2026.xlsx      ← gerado automaticamente
├── codigo/
│   ├── main.py
│   └── README.md
└── Monitoramento.pbix
```

---

## Pré-requisitos

Python 3.8 ou superior e a biblioteca `openpyxl`:

```bash
pip install openpyxl
```

---

## Como usar

1. Coloque o(s) arquivo(s) `.csv` exportados do sistema na pasta `bases/`
2. Pelo terminal, entre na pasta `codigo/` e execute:

```bash
python main.py
```

O script detecta todos os CSVs da pasta `bases/` automaticamente e gera o `.xlsx` correspondente na mesma pasta, com o mesmo nome.

```
17_06_2026.csv  →  17_06_2026.xlsx
18_06_2026.csv  →  18_06_2026.xlsx
```

---

## O que o script faz

- Lê os CSVs no encoding `ISO-8859-1` (padrão da extração do sistema)
- Remove linhas de controle (cabeçalhos repetidos, numeração de página, etc.)
- Normaliza todos os dados em **caixa alta**
- Mapeia as colunas corretamente: `Cartão`, `Credencial`, `Data`, `Evento`, `Usuário`, `Descrição`, `Grupo`
- Aplica a identidade visual da PUCPR conforme o guia de estilo
- Ativa **filtro automático** e **painel congelado** no cabeçalho

---

## Estrutura do Excel gerado

| Linha | Conteúdo |
|-------|----------|
| 1 | Nome da unidade |
| 2 | Título do relatório |
| 3 | Emissão, Período e Usuário da extração |
| 4 | Espaçador |
| 5 | Cabeçalho das colunas (com filtro) |
| 6+ | Dados — linhas alternando branco e cinza claro |

---

## Identidade visual aplicada

Paleta de cores conforme o **Guia de Estilo Visual PUCPR**:

| Elemento | Cor | Hex |
|----------|-----|-----|
| Cabeçalho principal | Primary Pure | `#8A0538` |
| Subtítulo | Primary Dark | `#570013` |
| Metadados | Primary Darkest | `#300000` |
| Linha par | Light Pure | `#FFFFFF` |
| Linha ímpar | Light Pure 02 | `#F0F2F2` |
| Texto dos dados | Dark 02 | `#404040` |

Fontes utilizadas: **Poppins** (cabeçalhos) e **Source Sans Pro** (dados).

---

## Observação sobre a coluna Usuário

No CSV exportado pelo sistema, a coluna `Usuário` pode conter tanto o nome de uma pessoa quanto a identificação de um veículo (ex: `FLUENCE AYB-5768`). Esse é o comportamento original do sistema — o script preserva o valor como está.