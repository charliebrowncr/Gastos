"""Actualización diaria de facturas Canavida.

Toma el Excel acumulado del día anterior (reportes/facturas_<fecha>.xlsx),
agrega las facturas nuevas (JSON) y genera:
  - reportes/facturas_<hoy>.xlsx  (hojas: Facturas, Resumen por mes, Gastos Diarios (nuevas))
  - reportes/resumen_facturas_<hoy>.pdf

Uso: python3 scripts/actualizar_facturas.py <excel_anterior> <nuevas.json> <AAAA-MM-DD>
"""
import json
import sys
from collections import defaultdict
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.fonts import addMapping

pdfmetrics.registerFont(TTFont('DV', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
pdfmetrics.registerFont(TTFont('DVB', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))
addMapping('DV', 0, 0, 'DV'); addMapping('DV', 1, 0, 'DVB')
addMapping('DV', 0, 1, 'DV'); addMapping('DV', 1, 1, 'DVB')

TC = 459.0  # ₡ por US$, mismo tipo de cambio usado en los reportes anteriores
HEAD = ['Mes', 'Fecha', 'Proveedor', 'Categoría', 'Descripción', 'Método', 'N° factura',
        'Monto ₡', 'Monto USD', 'Fila Gastos Diarios', 'PDF', 'Nota']
GD_HEAD = ['Fecha', 'Categoría', 'Descripción', 'Proveedor', 'Método de pago', 'Monto (USD)',
           'Nº factura / Notas']
BLUE = PatternFill('solid', fgColor='1F4E78')


def load_prev(path):
    ws = openpyxl.load_workbook(path)['Facturas']
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    return [list(r) for r in rows if r[0]]


def new_rows(path):
    out = []
    for f in json.load(open(path)):
        fecha = datetime.strptime(f['fecha'], '%Y-%m-%d')
        if f.get('usd') is not None:
            usd = f['usd']
            crc = round(usd * TC, 2)
        else:
            crc = f['crc']
            usd = round(crc / TC, 2)
        out.append([fecha.strftime('%Y-%m'), fecha, f['proveedor'], f['categoria'], f['descripcion'],
                    f['metodo'], f['numero'], crc, usd, None, f['pdf'], f.get('nota')])
    return out


def header(ws, cols):
    ws.append(cols)
    for c in ws[1]:
        c.font = Font(bold=True, color='FFFFFF')
        c.fill = BLUE
        c.alignment = Alignment(vertical='center', wrap_text=True)
    ws.freeze_panes = 'A2'


def resumen(rows):
    agg = defaultdict(lambda: [0, 0.0, 0.0])
    tot = defaultdict(lambda: [0, 0.0, 0.0])
    for r in rows:
        for d, k in ((agg, (r[0], r[3])), (tot, r[0])):
            d[k][0] += 1
            d[k][1] += r[7]
            d[k][2] += r[8]
    return agg, tot


def build_xlsx(rows, nuevas, out):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Facturas'
    header(ws, HEAD)
    for r in sorted(rows, key=lambda r: (r[1], r[2])):
        ws.append(r)
    for row in ws.iter_rows(min_row=2):
        row[1].number_format = 'DD/MM/YYYY'
        row[7].number_format = '#,##0.00'
        row[8].number_format = '#,##0.00'
    for col, w in zip('ABCDEFGHIJKL', [9, 11, 30, 26, 50, 20, 22, 14, 12, 10, 55, 55]):
        ws.column_dimensions[col].width = w

    agg, tot = resumen(rows)
    ws2 = wb.create_sheet('Resumen por mes')
    header(ws2, ['Mes', 'Categoría', 'Documentos', 'Monto ₡', 'Monto USD'])
    for (m, c), v in sorted(agg.items()):
        ws2.append([m, c, v[0], round(v[1], 2), round(v[2], 2)])
    for m, v in sorted(tot.items()):
        ws2.append([m, f'TOTAL {m}', v[0], round(v[1], 2), round(v[2], 2)])
        for c in ws2[ws2.max_row]:
            c.font = Font(bold=True)
    for col, w in zip('ABCDE', [9, 32, 12, 16, 14]):
        ws2.column_dimensions[col].width = w

    # Filas listas para pegar en la hoja "Gastos Diarios" del libro oficial
    ws3 = wb.create_sheet('Gastos Diarios (nuevas)')
    header(ws3, GD_HEAD)
    for r in nuevas:
        nota = r[6] + (f' · {r[11]}' if r[11] else '')
        ws3.append([r[1], r[3], r[4], r[2], r[5], r[8], nota])
        ws3.cell(ws3.max_row, 1).number_format = 'DD/MM/YYYY'
    for col, w in zip('ABCDEFG', [11, 26, 50, 30, 20, 12, 70]):
        ws3.column_dimensions[col].width = w
    wb.save(out)
    return tot


def build_pdf(rows, nuevas, tot, hoy, out):
    st = getSampleStyleSheet()
    for k in ('Title', 'BodyText', 'Heading2'):
        st[k].fontName = 'DVB' if k != 'BodyText' else 'DV'
    small = st['BodyText'].clone('s', fontName='DV', fontSize=7.5, leading=9)
    doc = SimpleDocTemplate(out, pagesize=landscape(letter), leftMargin=1.2 * cm, rightMargin=1.2 * cm,
                            topMargin=1.2 * cm, bottomMargin=1.2 * cm)
    el = [Paragraph(f'Canavida · Actualización diaria de facturas · {hoy:%d/%m/%Y}', st['Title'])]
    crc = sum(r[7] for r in nuevas)
    usd = sum(r[8] for r in nuevas)
    el.append(Paragraph(f'<b>Facturas nuevas hoy: {len(nuevas)}</b> · ₡{crc:,.2f} ≈ US${usd:,.2f} '
                        f'(TC ₡{TC:.0f})', st['BodyText']))
    el.append(Spacer(1, 8))

    def tabla(data, widths):
        t = Table(data, colWidths=widths, repeatRows=1)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E78')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTSIZE', (0, 0), (-1, -1), 7.5),
            ('GRID', (0, 0), (-1, -1), 0.3, colors.grey), ('FONTNAME', (0, 0), (-1, -1), 'DV'), ('FONTNAME', (0, 0), (-1, 0), 'DVB'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('ALIGN', (5, 1), (6, -1), 'RIGHT'),
        ]))
        return t

    data = [['Fecha', 'Proveedor', 'Categoría', 'Descripción', 'N° factura', 'Monto ₡', 'USD', 'Nota']]
    for r in nuevas:
        data.append([f'{r[1]:%d/%m/%Y}', Paragraph(r[2], small), Paragraph(r[3], small),
                     Paragraph(r[4], small), Paragraph(r[6], small), f'{r[7]:,.2f}', f'{r[8]:,.2f}',
                     Paragraph(r[11] or '', small)])
    el.append(tabla(data, [1.8 * cm, 3.6 * cm, 3 * cm, 6 * cm, 2.6 * cm, 2.2 * cm, 1.7 * cm, 5 * cm]))
    el.append(Spacer(1, 12))

    el.append(Paragraph('Acumulado por mes', st['Heading2']))
    data = [['Mes', 'Documentos', 'Monto ₡', 'Monto USD']]
    for m, v in sorted(tot.items()):
        data.append([m, v[0], f'{v[1]:,.2f}', f'{v[2]:,.2f}'])
    t = Table(data, colWidths=[3 * cm, 3 * cm, 4 * cm, 4 * cm])
    t.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E78')),
                           ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                           ('GRID', (0, 0), (-1, -1), 0.3, colors.grey), ('FONTNAME', (0, 0), (-1, -1), 'DV'), ('FONTNAME', (0, 0), (-1, 0), 'DVB'),
                           ('ALIGN', (1, 1), (-1, -1), 'RIGHT'), ('FONTSIZE', (0, 0), (-1, -1), 8.5)]))
    el.append(t)
    el.append(Spacer(1, 12))

    mes = f'{hoy:%Y-%m}'
    el.append(Paragraph(f'Detalle por categoría · {mes}', st['Heading2']))
    agg, _ = resumen(rows)
    data = [['Categoría', 'Documentos', 'Monto ₡', 'Monto USD']]
    for (m, c), v in sorted(agg.items()):
        if m == mes:
            data.append([c, v[0], f'{v[1]:,.2f}', f'{v[2]:,.2f}'])
    t = Table(data, colWidths=[6 * cm, 3 * cm, 4 * cm, 4 * cm])
    t.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E78')),
                           ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                           ('GRID', (0, 0), (-1, -1), 0.3, colors.grey), ('FONTNAME', (0, 0), (-1, -1), 'DV'), ('FONTNAME', (0, 0), (-1, 0), 'DVB'),
                           ('ALIGN', (1, 1), (-1, -1), 'RIGHT'), ('FONTSIZE', (0, 0), (-1, -1), 8.5)]))
    el.append(t)
    doc.build(el)


def main():
    prev, nuevas_json, hoy = sys.argv[1], sys.argv[2], datetime.strptime(sys.argv[3], '%Y-%m-%d')
    rows = load_prev(prev)
    nuevas = new_rows(nuevas_json)
    vistos = {(r[2], r[6]) for r in rows}
    nuevas = [n for n in nuevas if (n[2], n[6]) not in vistos]
    rows += nuevas
    tag = f'{hoy:%Y-%m-%d}'
    tot = build_xlsx(rows, nuevas, f'reportes/facturas_{tag}.xlsx')
    build_pdf(rows, nuevas, tot, hoy, f'reportes/resumen_facturas_{tag}.pdf')
    m = f'{hoy:%Y-%m}'
    print(json.dumps({'nuevas': len(nuevas), 'crc': round(sum(r[7] for r in nuevas), 2),
                      'usd': round(sum(r[8] for r in nuevas), 2),
                      'mes': [tot[m][0], round(tot[m][1], 2), round(tot[m][2], 2)]}))


if __name__ == '__main__':
    main()
