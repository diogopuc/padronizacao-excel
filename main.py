"""
Padronizador de base diária de credenciados — PUCPR / Grupo Marista
Lê todos os CSVs de ../bases/ e gera os XLSXs na mesma pasta.
Uso: python main.py
"""

import os
import glob
from openpyxl import Workbook
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


if __name__ == '__main__':
    pasta = os.path.abspath(PASTA_BASES)

    if not os.path.isdir(pasta):
        print(f"[ERRO] Pasta não encontrada: {pasta}")
        exit(1)

    arquivos = glob.glob(os.path.join(pasta, '*.csv'))

    if not arquivos:
        print(f"Nenhum CSV encontrado em: {pasta}")
        exit(0)

    print(f"Encontrados {len(arquivos)} arquivo(s) em: {pasta}\n")
    for arquivo in sorted(arquivos):
        processar(arquivo)

    print("\nConcluído.")