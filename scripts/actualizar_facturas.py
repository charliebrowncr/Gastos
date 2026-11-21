"""Actualización diaria de facturas Canavida.

Toma el Excel acumulado anterior (reportes/facturas_<fecha>.xlsx, hoja 'Facturas'),
agrega las facturas nuevas (JSON) y genera:
  - reportes/facturas_<hoy>.xlsx  (hojas: Facturas, una hoja por mes AAAA-MM, Resumen por mes)
  - reportes/resumen_facturas_<hoy>.pdf  (nuevas del día, resumen por mes/categoría, detalle por mes)

Uso: python3 scripts/actualizar_facturas.py <excel_anterior> <nuevas.json> <AAAA-MM-DD> ["nota para revisar"]
"""
import json
import sys
from collections import defaultdict
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from reportlab.lib import colors
from reportlab.lib.fonts import addMapping
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

pdfmetrics.registerFont(TTFont('DV', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
pdfmetrics.registerFont(TTFont('DVB', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))
addMapping('DV', 0, 0, 'DV'); addMapping('DV', 1, 0, 'DVB'); addMapping('DV', 0, 1, 'DV'); addMapping('DV', 1, 1, 'DVB')

TC = 459.0  # ₡ por US$, mismo tipo de cambio de los reportes anteriores
HEAD = ['Mes', 'Fecha', 'Proveedor', 'Categoría', 'Descripción', 'Método', 'N° factura',
        'Monto ₡', 'Monto USD', 'Fila Gastos Diarios', 'PDF', 'Nota']
WIDTHS = [9, 12, 32, 28, 55, 20, 24, 14, 12, 10, 62, 60]
FILL = PatternFill('solid', fgColor='305496')
AZUL = colors.HexColor('#305496')


def load_prev(path):
    ws = openpyxl.load_workbook(path)['Facturas']
    return [list(r) for r in ws.iter_rows(min_row=2, values_only=True) if r[0]]


def new_rows(path):
    out = []
    for f in json.load(open(path)):
        fecha = datetime.strptime(f['fecha'], '%Y-%m-%d')
        if f.get('usd') is not None:
            usd, crc = f['usd'], round(f['usd'] * TC, 2)
        else:
            crc, usd = f['crc'], round(f['crc'] / TC, 2)
        out.append([fecha.strftime('%Y-%m'), fecha, f['proveedor'], f['categoria'], f['descripcion'],
                    f['metodo'], f['numero'], crc, usd, None, f['pdf'], f.get('nota')])
    return out


def header(ws, cols, widths):
    ws.append(cols)
    for c in ws[1]:
        c.font = Font(bold=True, color='FFFFFF')
        c.fill = FILL
        c.alignment = Alignment(vertical='center', wrap_text=True)
    ws.freeze_panes = 'A2'
    for i, w in enumerate(widths):
        ws.column_dimensions[chr(65 + i)].width = w


def fmt(ws):
    for row in ws.iter_rows(min_row=2):
        row[1].number_format = 'DD/MM/YYYY'
        row[7].number_format = '#,##0.00'
        row[8].number_format = '#,##0.00'


def resumen(rows):
    agg = defaultdict(lambda: [0, 0.0, 0.0])
    tot = defaultdict(lambda: [0, 0.0, 0.0])
    for r in rows:
        for d, k in ((agg, (r[0], r[3])), (tot, r[0])):
            d[k][0] += 1; d[k][1] += r[7]; d[k][2] += r[8]
    return agg, tot


def build_xlsx(rows, out):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Facturas'
    header(ws, HEAD, WIDTHS)
    for r in rows:
        ws.append(r)
    fmt(ws)
    for m in sorted({r[0] for r in rows}):
        wm = wb.create_sheet(m)
        header(wm, HEAD, WIDTHS)
        for r in rows:
            if r[0] == m:
                wm.append(r)
        fmt(wm)
    agg, tot = resumen(rows)
    w2 = wb.create_sheet('Resumen por mes')
    header(w2, ['Mes', 'Categoría', 'Documentos', 'Monto ₡', 'Monto USD'], [10, 32, 12, 16, 14])
    for (m, c), v in sorted(agg.items()):
        w2.append([m, c, v[0], round(v[1], 2), round(v[2], 2)])
    for m, v in sorted(tot.items()):
        w2.append([m, f'TOTAL {m}', v[0], round(v[1], 2), round(v[2], 2)])
        for c in w2[w2.max_row]:
            c.font = Font(bold=True)
    for row in w2.iter_rows(min_row=2):
        row[3].number_format = row[4].number_format = '#,##0.00'
    wb.save(out)
    return agg, tot


def build_pdf(rows, nuevas, agg, tot, hoy, nota, out):
    st = getSampleStyleSheet()
    st['Title'].fontName = st['Heading2'].fontName = 'DVB'
    st['BodyText'].fontName = 'DV'
    sm = st['BodyText'].clone('sm', fontName='DV', fontSize=7.5, leading=9)
    doc = SimpleDocTemplate(out, pagesize=landscape(letter), leftMargin=1.2 * cm, rightMargin=1.2 * cm,
                            topMargin=1.2 * cm, bottomMargin=1.2 * cm, title='Resumen de facturas Canavida')
    base = [('BACKGROUND', (0, 0), (-1, 0), AZUL), ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('GRID', (0, 0), (-1, -1), 0.3, colors.grey), ('FONTNAME', (0, 0), (-1, -1), 'DV'),
            ('FONTNAME', (0, 0), (-1, 0), 'DVB'), ('FONTSIZE', (0, 0), (-1, -1), 7.5), ('VALIGN', (0, 0), (-1, -1), 'TOP')]

    def detalle(rs):
        data = [['Fecha', 'Proveedor', 'Categoría', 'Descripción', 'Método', 'N°', 'Monto ₡', 'USD']]
        for r in rs:
            data.append([f'{r[1]:%d/%m}', Paragraph(r[2], sm), Paragraph(r[3], sm), Paragraph(r[4], sm),
                         Paragraph(r[5], sm), Paragraph(r[6], sm), f'{r[7]:,.2f}', f'{r[8]:,.2f}'])
        t = Table(data, colWidths=[1.3*cm, 4*cm, 3.2*cm, 8*cm, 2.6*cm, 3*cm, 2.4*cm, 1.8*cm], repeatRows=1)
        t.setStyle(TableStyle(base + [('ALIGN', (6, 1), (7, -1), 'RIGHT')]))
        return t

    crc = sum(r[7] for r in nuevas); usd = sum(r[8] for r in nuevas)
    el = [Paragraph(f'Canavida · Actualización diaria de facturas — {hoy:%d/%m/%Y}', st['Title']),
          Paragraph(f'<b>Nuevas hoy: {len(nuevas)} facturas · ₡{crc:,.2f} ≈ US${usd:,.2f}</b> (TC ₡{TC:.0f})', st['BodyText']),
          Spacer(1, 6), detalle(nuevas), Spacer(1, 6)]
    if nota:
        el.append(Paragraph(f'<b>Para revisar:</b> {nota}', sm))
    el += [Spacer(1, 8), Paragraph('Resumen por mes y categoría', st['Heading2'])]
    data = [['Mes', 'Categoría', 'Docs', 'Monto ₡', 'Monto USD']]
    for (m, c), v in sorted(agg.items()):
        data.append([m, c, v[0], f'{v[1]:,.2f}', f'{v[2]:,.2f}'])
    nt = len(data)
    for m, v in sorted(tot.items()):
        data.append([m, f'TOTAL {m}', v[0], f'{v[1]:,.2f}', f'{v[2]:,.2f}'])
    t = Table(data, colWidths=[2.2*cm, 6*cm, 1.6*cm, 3.4*cm, 3*cm])
    t.setStyle(TableStyle(base + [('ALIGN', (2, 1), (-1, -1), 'RIGHT'), ('FONTNAME', (0, nt), (-1, -1), 'DVB'),
                                  ('BACKGROUND', (0, nt), (-1, -1), colors.HexColor('#D9E1F2'))]))
    el += [t, PageBreak()]
    for m in sorted(tot):
        el += [Paragraph(f'Detalle {m}', st['Heading2']), detalle([r for r in rows if r[0] == m]), Spacer(1, 8)]
    doc.build(el)


def main():
    prev, nj, hoy = sys.argv[1], sys.argv[2], datetime.strptime(sys.argv[3], '%Y-%m-%d')
    nota = sys.argv[4] if len(sys.argv) > 4 else ''
    rows = load_prev(prev)
    vistos = {(r[2], r[6]) for r in rows}
    nuevas = [n for n in new_rows(nj) if (n[2], n[6]) not in vistos]
    rows = sorted(rows + nuevas, key=lambda r: (r[1], r[2]))
    tag = f'{hoy:%Y-%m-%d}'
    agg, tot = build_xlsx(rows, f'reportes/facturas_{tag}.xlsx')
    build_pdf(rows, nuevas, agg, tot, hoy, nota, f'reportes/resumen_facturas_{tag}.pdf')
    m = f'{hoy:%Y-%m}'
    print(json.dumps({'nuevas': len(nuevas), 'crc': round(sum(r[7] for r in nuevas), 2),
                      'usd': round(sum(r[8] for r in nuevas), 2),
                      'mes': [tot[m][0], round(tot[m][1], 2), round(tot[m][2], 2)]}))


if __name__ == '__main__':
    main()
