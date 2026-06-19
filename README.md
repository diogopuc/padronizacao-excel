# Padronizador de Base Diária de Credenciados
**PUCPR / Grupo Marista**

Script Python que converte automaticamente a extração diária do sistema de controle de acesso (`.csv`) em uma planilha Excel padronizada com o guia de estilo visual da PUCPR, além de manter um histórico consolidado por unidade (Curitiba/Toledo).

---

## Estrutura do projeto

```
dashboard - monitoramento check-in/
├── bases/
│   ├── 17_06_2026.csv       ← base exportada do sistema
│   ├── 17_06_2026.xlsx      ← gerado automaticamente (individual)
│   └── consolidado.xlsx     ← gerado/atualizado automaticamente (histórico por unidade)
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

O script detecta todos os CSVs da pasta `bases/` automaticamente e gera o `.xlsx` correspondente na mesma pasta, com o mesmo nome. Em seguida, atualiza automaticamente o `consolidado.xlsx` com o histórico por unidade — esse passo roda sempre, mesmo quando não há CSV novo, para absorver qualquer XLSX já presente em `bases/` que ainda não tenha entrado no histórico.

```
17_06_2026.csv  →  17_06_2026.xlsx  →  consolidado.xlsx (aba da unidade)
18_06_2026.csv  →  18_06_2026.xlsx  →  consolidado.xlsx (aba da unidade)
```

---

## O que o script faz

- Lê os CSVs no encoding `ISO-8859-1` (padrão da extração do sistema)
- Remove linhas de controle (cabeçalhos repetidos, numeração de página, etc.)
- Normaliza todos os dados em **caixa alta**
- Mapeia as colunas corretamente: `Cartão`, `Credencial`, `Data`, `Evento`, `Usuário`, `Descrição`, `Grupo`
- Aplica a identidade visual da PUCPR conforme o guia de estilo
- Ativa **filtro automático** e **painel congelado** no cabeçalho
- Consolida o histórico de todas as bases em `consolidado.xlsx`, separado por unidade, sem duplicar registros já existentes

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

## Consolidado histórico (`consolidado.xlsx`)

Além dos XLSX individuais, o script mantém um arquivo `bases/consolidado.xlsx` com o histórico acumulado de todas as bases já processadas, separado em duas abas fixas:

- `PUCPR - CURITIBA`
- `PUCPR - TOLEDO`

**Como a unidade é identificada:** o script lê o texto de unidade (linha 1 do XLSX individual) e procura, de forma case-insensitive e sem considerar acentos, pelas palavras `PARANÁ` (→ Curitiba) ou `TOLEDO` (→ Toledo). Bases de unidades não reconhecidas são ignoradas na consolidação (com aviso no terminal), mas continuam gerando seu XLSX individual normalmente.

**Como funciona a leitura:** o script processa **qualquer** arquivo `.xlsx` da pasta `bases/` (exceto o próprio `consolidado.xlsx`), independente do nome. O cabeçalho é localizado dinamicamente (procurando a coluna "Cartão"), então uma coluna nova que surja na origem futuramente é incorporada automaticamente ao consolidado, sem quebrar o processamento.

**Como evita duplicidade:** antes de inserir, cada registro é comparado por uma chave composta de `Cartão + Credencial + Data + Evento`. Só registros inéditos são adicionados — rodar o script várias vezes sobre a mesma base não duplica linhas.

**Como é atualizado:** a cada execução, o histórico é acumulado (novas linhas sempre abaixo das existentes); nada do histórico anterior é sobrescrito ou removido.

---

## Observação sobre a coluna Usuário

No CSV exportado pelo sistema, a coluna `Usuário` pode conter tanto o nome de uma pessoa quanto a identificação de um veículo (ex: `FLUENCE AYB-5768`). Esse é o comportamento original do sistema — o script preserva o valor como está.