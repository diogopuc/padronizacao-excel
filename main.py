"""
Padronizador de base diária de credenciados — PUCPR / Grupo Marista
Lê todos os CSVs de ../bases/ e gera os XLSXs na mesma pasta.
Uso: python main.py
"""

import os
import glob
import unicodedata
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ── Paleta PUCPR ──────────────────────────────────────────────────────────────
PRIMARY_PURE     = "8A0538"
PRIMARY_DARK     = "570013"
PRIMARY_DARKEST  = "300000"
PRIMARY_LIGHTEST = "E5C3D0"
LIGHT_PURE_02    = "F0F2F2"
LIGHT_PURE       = "FFFFFF"
DARK_02          = "404040"

SKIP_KEYWORDS = ['Cartão', 'Credencial', 'Página', 'UNIDADE', 'Relatório',
                 'Emissão', 'Período', 'Usuário:', 'Cartão:']

# ── Caminho da pasta de bases (relativo ao script) ────────────────────────────
PASTA_BASES = os.path.join(os.path.dirname(__file__), '..', 'bases')

# ── Consolidação histórica por unidade ────────────────────────────────────────
NOME_CONSOLIDADO = 'consolidado.xlsx'
ABA_CURITIBA     = 'PUCPR - CURITIBA'
ABA_TOLEDO       = 'PUCPR - TOLEDO'
CHAVE_DEDUP      = ['CARTÃO', 'CREDENCIAL', 'DATA', 'EVENTO']


def fill(hex_color):
    return PatternFill('solid', start_color=hex_color)


def _parse_csv(filepath):
    with open(filepath, 'r', encoding='ISO-8859-1') as f:
        lines = f.readlines()

    meta = {'unidade': '', 'emissao': '', 'periodo': '', 'usuario': ''}
    for line in lines[:7]:
        cols = line.strip().split(';')
        joined = ' '.join(c.strip() for c in cols if c.strip())
        if 'UNIDADE'   in joined: meta['unidade'] = joined.strip()
        elif 'Emissão'  in joined: meta['emissao'] = joined.replace('Emissão:', '').strip()
        elif 'Período'  in joined: meta['periodo'] = joined.replace('Período:', '').replace('até', 'ATÉ').strip()
        elif 'Usuário:' in joined: meta['usuario'] = joined.replace('Usuário:', '').strip()

    rows = []
    for line in lines:
        cols = line.strip().split(';')
        while len(cols) < 16:
            cols.append('')
        if any(kw in ';'.join(cols) for kw in SKIP_KEYWORDS):
            continue
        if not any(c.strip() for c in cols):
            continue

        cartao     = (cols[1].strip() or cols[2].strip()).upper()
        credencial = (cols[3].strip() if cols[1].strip() else cols[2].strip()).upper()
        data       = cols[4].strip().upper()
        evento     = cols[6].strip().upper()
        usuario    = cols[10].strip().upper()
        descricao  = cols[12].strip().upper()
        grupo      = cols[15].strip().upper()

        if data and '/' in data:
            rows.append([cartao, credencial, data, evento, usuario, descricao, grupo])

    return meta, rows


def _build_xlsx(meta, rows, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = os.path.splitext(os.path.basename(output_path))[0]

    borda = Border(
        left=Side(style='thin', color='E4E4E4'),
        right=Side(style='thin', color='E4E4E4'),
        top=Side(style='thin', color='E4E4E4'),
        bottom=Side(style='thin', color='E4E4E4')
    )
    centro = Alignment(horizontal='center', vertical='center')
    esq    = Alignment(horizontal='left',   vertical='center')

    ws.merge_cells('A1:G1')
    ws['A1'] = meta['unidade'].upper()
    ws['A1'].font      = Font(name='Poppins', bold=True, size=13, color='FFFFFF')
    ws['A1'].fill      = fill(PRIMARY_PURE)
    ws['A1'].alignment = centro
    ws.row_dimensions[1].height = 24

    ws.merge_cells('A2:G2')
    ws['A2'] = 'RELATÓRIO DE CARTÕES DE CREDENCIADOS'
    ws['A2'].font      = Font(name='Poppins', bold=True, size=11, color='FFFFFF')
    ws['A2'].fill      = fill(PRIMARY_DARK)
    ws['A2'].alignment = centro
    ws.row_dimensions[2].height = 20

    ws.merge_cells('A3:G3')
    ws['A3'] = (
        f"EMISSÃO: {meta['emissao'].upper()}"
        f"   |   PERÍODO: {meta['periodo'].upper()}"
        f"   |   USUÁRIO: {meta['usuario'].upper()}"
    )
    ws['A3'].font      = Font(name='Poppins', size=9, color=PRIMARY_LIGHTEST)
    ws['A3'].fill      = fill(PRIMARY_DARKEST)
    ws['A3'].alignment = centro
    ws.row_dimensions[3].height = 16

    ws.merge_cells('A4:G4')
    ws['A4'].fill = fill(LIGHT_PURE_02)
    ws.row_dimensions[4].height = 6

    headers = ['CARTÃO', 'CREDENCIAL', 'DATA', 'EVENTO', 'USUÁRIO', 'DESCRIÇÃO', 'GRUPO']
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=5, column=col_idx, value=h)
        cell.font      = Font(name='Poppins', bold=True, size=10, color='FFFFFF')
        cell.fill      = fill(PRIMARY_PURE)
        cell.alignment = centro
        cell.border    = Border(bottom=Side(style='medium', color=PRIMARY_DARK))
    ws.row_dimensions[5].height = 20

    for i, row in enumerate(rows):
        excel_row = i + 6
        bg = LIGHT_PURE if i % 2 == 0 else LIGHT_PURE_02
        for col_idx, val in enumerate(row, start=1):
            cell = ws.cell(row=excel_row, column=col_idx, value=val)
            cell.font      = Font(name='Source Sans Pro', size=9, color=DARK_02)
            cell.fill      = fill(bg)
            cell.border    = borda
            cell.alignment = esq
        ws.row_dimensions[excel_row].height = 15

    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 22
    ws.column_dimensions['E'].width = 44
    ws.column_dimensions['F'].width = 14
    ws.column_dimensions['G'].width = 24

    ws.auto_filter.ref = f"A5:G{len(rows) + 5}"
    ws.freeze_panes    = 'A6'

    wb.save(output_path)


def processar(csv_path):
    nome_base   = os.path.splitext(os.path.basename(csv_path))[0]
    output_path = os.path.join(os.path.dirname(csv_path), f"{nome_base}.xlsx")

    print(f"Processando: {os.path.basename(csv_path)}")
    meta, rows = _parse_csv(csv_path)
    _build_xlsx(meta, rows, output_path)
    print(f"  ✓ {len(rows)} registros → {os.path.basename(output_path)}")


# ── Consolidação histórica por unidade ────────────────────────────────────────

def _normalizar(texto):
    """Remove acentos e caixa para comparação robusta (ex.: 'PARANÁ' -> 'PARANA')."""
    sem_acento = unicodedata.normalize('NFKD', texto or '')
    return ''.join(c for c in sem_acento if not unicodedata.combining(c)).upper()


def obter_unidade(texto_unidade):
    """Identifica a aba de destino (Curitiba/Toledo) a partir do texto de unidade do relatório."""
    texto = _normalizar(texto_unidade)
    if 'TOLEDO' in texto:
        return ABA_TOLEDO
    if 'PARANA' in texto:
        return ABA_CURITIBA
    return None


def _localizar_linha_cabecalho(ws, limite_busca=10):
    """Procura nas primeiras linhas a que contém o cabeçalho 'CARTÃO', sem depender de posição fixa
    (arquivos gerados por versões anteriores do script podem não ter a linha espaçadora)."""
    for r in range(1, limite_busca + 1):
        valor = ws.cell(row=r, column=1).value
        if valor and _normalizar(str(valor)) == 'CARTAO':
            return r
    return None


def _ler_individual(xlsx_path):
    """Lê unidade, cabeçalho e dados de um XLSX já gerado por _build_xlsx, localizando o
    cabeçalho dinamicamente para suportar colunas novas ou layouts ligeiramente diferentes."""
    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb.active

    unidade_texto = ws['A1'].value or ''

    linha_cabecalho = _localizar_linha_cabecalho(ws)
    if linha_cabecalho is None:
        return unidade_texto, [], []

    headers = []
    col = 1
    while ws.cell(row=linha_cabecalho, column=col).value:
        headers.append(str(ws.cell(row=linha_cabecalho, column=col).value).strip())
        col += 1

    rows = []
    for linha in ws.iter_rows(min_row=linha_cabecalho + 1, max_col=len(headers), values_only=True):
        valores = [v if v is not None else '' for v in linha]
        if any(str(v).strip() for v in valores):
            rows.append(valores)

    return unidade_texto, headers, rows


def carregar_consolidado(caminho):
    """Abre o consolidado existente ou cria um novo, garantindo as duas abas obrigatórias."""
    if os.path.exists(caminho):
        wb = load_workbook(caminho)
    else:
        wb = Workbook()
        wb.remove(wb.active)

    for nome_aba in (ABA_CURITIBA, ABA_TOLEDO):
        if nome_aba not in wb.sheetnames:
            wb.create_sheet(nome_aba)

    return wb


def _garantir_cabecalho(ws, headers_origem):
    """Escreve o cabeçalho na aba se ainda não existir; acrescenta colunas novas ao final, se houver."""
    cabecalho_atual = [c.value for c in ws[1] if c.value] if ws['A1'].value else []

    novas_colunas = [h for h in headers_origem if h not in cabecalho_atual]
    inicio = len(cabecalho_atual) + 1
    for offset, h in enumerate(novas_colunas):
        cell = ws.cell(row=1, column=inicio + offset, value=h)
        cell.font = Font(name='Poppins', bold=True, size=10, color='FFFFFF')
        cell.fill = fill(PRIMARY_PURE)
    cabecalho_atual.extend(novas_colunas)

    if cabecalho_atual:
        ws.freeze_panes = 'A2'

    return cabecalho_atual


def registro_ja_existe(chave, chaves_existentes):
    return chave in chaves_existentes


def inserir_registros(ws, cabecalho, headers_origem, rows):
    """Insere apenas registros inéditos abaixo do histórico já presente na aba."""
    indices_chave = [cabecalho.index(c) for c in CHAVE_DEDUP if c in cabecalho]
    mapa_colunas  = [cabecalho.index(h) for h in headers_origem]

    chaves_existentes = set()
    for linha in ws.iter_rows(min_row=2, values_only=True):
        chave = tuple(str(linha[i]).strip().upper() if i < len(linha) and linha[i] is not None else ''
                       for i in indices_chave)
        chaves_existentes.add(chave)

    proxima_linha = ws.max_row + 1
    inseridos = 0
    for row in rows:
        linha_destino = [''] * len(cabecalho)
        for col_origem, col_destino in enumerate(mapa_colunas):
            linha_destino[col_destino] = row[col_origem]

        chave = tuple(str(linha_destino[i]).strip().upper() for i in indices_chave)
        if registro_ja_existe(chave, chaves_existentes):
            continue

        for col_idx, valor in enumerate(linha_destino, start=1):
            cell = ws.cell(row=proxima_linha, column=col_idx, value=valor)
            cell.font = Font(name='Source Sans Pro', size=9, color=DARK_02)

        chaves_existentes.add(chave)
        proxima_linha += 1
        inseridos += 1

    return inseridos


def consolidar_planilhas(pasta_bases):
    """Acumula no consolidado.xlsx os dados de todos os XLSX da pasta, separados por unidade."""
    caminho_consolidado = os.path.join(pasta_bases, NOME_CONSOLIDADO)
    arquivos = sorted(
        f for f in glob.glob(os.path.join(pasta_bases, '*.xlsx'))
        if os.path.basename(f) != NOME_CONSOLIDADO
    )

    if not arquivos:
        return

    wb_consolidado = carregar_consolidado(caminho_consolidado)
    total_inseridos = {ABA_CURITIBA: 0, ABA_TOLEDO: 0}

    for arquivo in arquivos:
        unidade_texto, headers, rows = _ler_individual(arquivo)
        nome_aba = obter_unidade(unidade_texto)
        if nome_aba is None:
            print(f"  [AVISO] Unidade não identificada em {os.path.basename(arquivo)} — ignorado na consolidação")
            continue

        ws = wb_consolidado[nome_aba]
        cabecalho = _garantir_cabecalho(ws, headers)
        total_inseridos[nome_aba] += inserir_registros(ws, cabecalho, headers, rows)

    for nome_aba in (ABA_CURITIBA, ABA_TOLEDO):
        ws = wb_consolidado[nome_aba]
        if ws.max_row >= 2 and ws.max_column >= 1:
            ultima_coluna = ws.cell(row=1, column=ws.max_column).column_letter
            ws.auto_filter.ref = f"A1:{ultima_coluna}{ws.max_row}"

    wb_consolidado.save(caminho_consolidado)
    print(f"\nConsolidado atualizado: +{total_inseridos[ABA_CURITIBA]} CURITIBA, "
          f"+{total_inseridos[ABA_TOLEDO]} TOLEDO ({NOME_CONSOLIDADO})")


if __name__ == '__main__':
    pasta = os.path.abspath(PASTA_BASES)

    if not os.path.isdir(pasta):
        print(f"[ERRO] Pasta não encontrada: {pasta}")
        exit(1)

    arquivos = glob.glob(os.path.join(pasta, '*.csv'))

    if arquivos:
        print(f"Encontrados {len(arquivos)} arquivo(s) em: {pasta}\n")
        for arquivo in sorted(arquivos):
            processar(arquivo)
    else:
        print(f"Nenhum CSV encontrado em: {pasta}")

    # A consolidação roda sempre, mesmo sem CSV novo, para absorver XLSX já
    # presentes em bases/ que ainda não tenham sido incorporados ao histórico.
    consolidar_planilhas(pasta)

    print("\nConcluído.")