#!/usr/bin/env python3
"""Agrega facturas extraídas (JSON) a la hoja 'Gastos Diarios' y genera el reporte PDF.

Uso:
  python3 automatizacion/actualizar_gastos.py <carpeta_json> <salida_dir> [tipo_cambio_default]

- Lee cada <carpeta_json>/*.json (una factura PDF por archivo).
- Omite facturas ya registradas (mismo Nº de factura en columna G, o message_id ya
  procesado según automatizacion/procesados.json).
- Inserta las nuevas filas después del último gasto, ordenadas por fecha (por mes).
- Montos en colones se registran como fórmula =CRC/TC, igual que los registros manuales.
- Genera <salida_dir>/Canavida_Gastos_<fecha>.xlsx y Reporte_Gastos_<fecha>.pdf.
"""
import datetime as dt
import glob
import json
import os
import shutil
import sys
from collections import defaultdict

import openpyxl
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX = os.path.join(REPO, "Copy of Canavida Gastos Ingresos Oficial (2).xlsx")
STATE = os.path.join(REPO, "automatizacion", "procesados.json")
SHEET = "Gastos Diarios"
FIRST_ROW = 6
INICIO = "2026-10-01"
MESES = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto",
         "Septiembre", "Octubre", "Noviembre", "Diciembre"]


def to_usd(amount, tc):
    """Evalúa un monto de la hoja (número o fórmula simple '=a/b')."""
    if amount is None:
        return 0.0
    if isinstance(amount, (int, float)):
        return float(amount)
    s = str(amount).lstrip("=")
    try:
        if "/" in s:
            a, b = s.split("/", 1)
            return float(a) / float(b)
        return float(s)
    except ValueError:
        return 0.0


def main():
    src, outdir = sys.argv[1], sys.argv[2]
    tc_default = float(sys.argv[3]) if len(sys.argv) > 3 else 459.0
    os.makedirs(outdir, exist_ok=True)
    hoy = dt.date.today()

    state = {"message_ids": [], "consecutivos": []}
    if os.path.exists(STATE):
        state = json.load(open(STATE))

    wb = openpyxl.load_workbook(XLSX)
    ws = wb[SHEET]
    cats = [c.value for c in wb["Listas"]["A"][4:40] if c.value]

    # Última fila con gasto real (columna B o F con dato)
    last = FIRST_ROW - 1
    existing = []  # (fecha, Nº factura)
    for r in range(FIRST_ROW, ws.max_row + 1):
        if ws.cell(r, 2).value or ws.cell(r, 6).value:
            last = r
            g = ws.cell(r, 7).value
            f = ws.cell(r, 1).value
            if isinstance(g, (int, float)):
                existing.append((f.date() if isinstance(f, dt.datetime) else None, int(g)))

    def ya_registrada(d):
        c = d.get("consec_corto")
        if c is None:
            return False
        c = int(c)
        fd = dt.date.fromisoformat(d["fecha"])
        for f, g in existing:
            if g == c and (f is None or abs((f - fd).days) <= 45):
                return True
            # Nº escrito abreviado (ej. 687 por 20687) con fecha cercana
            if len(str(g)) >= 3 and str(c).endswith(str(g)) and f and abs((f - fd).days) <= 5:
                return True
        return False

    nuevos, omitidos, revisar = [], [], []
    for path in sorted(glob.glob(os.path.join(src, "*.json"))):
        d = json.load(open(path))
        if "total" not in d or not d.get("fecha"):
            if d.get("nota"):
                omitidos.append((d.get("thread_id") or d.get("message_id"), d.get("nota")))
            else:
                revisar.append(d)
            continue
        key = d.get("consecutivo") or f'{d.get("proveedor")}-{d.get("consec_corto")}'
        if key in state["consecutivos"]:
            omitidos.append((key, "ya procesado en una corrida anterior"))
            continue
        if d.get("consec_corto") in (None, "") or d.get("total") in (None, ""):
            omitidos.append((d.get("proveedor"), "no es factura (recibo/comprobante/sin total) – no se registra"))
            continue
        if d["fecha"] < INICIO:
            omitidos.append((key, f"fecha {d['fecha']} fuera del periodo de la hoja"))
            continue
        if ya_registrada(d):
            omitidos.append((key, "Nº de factura ya registrado en la hoja"))
            state["consecutivos"].append(key)
            continue
        if any(n.get("consecutivo") == d.get("consecutivo") and d.get("consecutivo") for n in nuevos):
            omitidos.append((key, "PDF duplicado en otro correo"))
            continue
        if d.get("categoria") not in cats:
            d["categoria"] = "Imprevistos y otros"
        nuevos.append(d)

    nuevos.sort(key=lambda d: (d["fecha"], d.get("proveedor", "")))

    # Limpia filas vacías con solo fecha (marcadores) a partir de last+1
    r = last + 1
    for d in nuevos:
        fecha = dt.datetime.strptime(d["fecha"], "%Y-%m-%d")
        signo = -1 if "crédito" in (d.get("tipo_doc") or "").lower() else 1
        total = float(d["total"]) * signo
        if (d.get("moneda") or "CRC").upper() == "USD":
            monto = round(total, 2)
        else:
            tc = d.get("tipo_cambio") or tc_default
            if tc and float(tc) < 100:  # tipo de cambio inválido
                tc = tc_default
            monto = f"={round(total, 2)}/{int(tc) if float(tc).is_integer() else tc}"
        nota = "Auto: correo " + (hoy.strftime("%d/%m/%Y"))
        if d.get("tipo_doc") and "factura" not in d["tipo_doc"].lower():
            nota += f" · {d['tipo_doc']}"
        if d.get("confianza") == "baja":
            nota += " · REVISAR: " + (d.get("nota") or "lectura dudosa")
            revisar.append(d)
        vals = [fecha, d["categoria"], d.get("descripcion"), d.get("proveedor"),
                d.get("forma_pago") or "Tarjeta", monto, d.get("consec_corto"), nota]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(r, c)
            cell.value = v
            src_cell = ws.cell(last, c)
            if src_cell.has_style:
                cell._style = src_cell._style
        d["_usd"] = to_usd(monto, None)
        r += 1
        state["consecutivos"].append(d.get("consecutivo") or f'{d.get("proveedor")}-{d.get("consec_corto")}')
    # Borra marcadores vacíos restantes (fecha sin datos) justo después
    while ws.cell(r, 1).value and not any(ws.cell(r, c).value for c in range(2, 9)):
        ws.cell(r, 1).value = None
        r += 1

    stamp = hoy.strftime("%Y-%m-%d")
    wb.save(XLSX)
    xlsx_out = os.path.join(outdir, f"Canavida_Gastos_{stamp}.xlsx")
    shutil.copy(XLSX, xlsx_out)
    json.dump(state, open(STATE, "w"), ensure_ascii=False, indent=1)

    # ---------- Resumen por mes (toda la hoja) ----------
    wb2 = openpyxl.load_workbook(XLSX)
    ws2 = wb2[SHEET]
    por_mes = defaultdict(lambda: defaultdict(float))
    cnt_mes = defaultdict(int)
    for row in ws2.iter_rows(min_row=FIRST_ROW, values_only=True):
        if not row[0] or not row[1] or row[5] is None:
            continue
        f = row[0]
        k = (f.year, f.month)
        por_mes[k][row[1]] += to_usd(row[5], None)
        cnt_mes[k] += 1

    pdf_out = os.path.join(outdir, f"Reporte_Gastos_{stamp}.pdf")
    st = getSampleStyleSheet()
    small = st["BodyText"].clone("small", fontSize=8, leading=10)
    doc = SimpleDocTemplate(pdf_out, pagesize=landscape(letter), leftMargin=36, rightMargin=36,
                            topMargin=36, bottomMargin=36)
    el = [Paragraph("Canavida Villas &amp; Resort – Actualización diaria de gastos", st["Title"]),
          Paragraph(f"Fecha: {hoy.strftime('%d/%m/%Y')} · Facturas nuevas agregadas: {len(nuevos)} · "
                    f"Omitidas (ya registradas/duplicadas): {len(omitidos)} · Para revisar: {len(revisar)}",
                    st["Normal"]), Spacer(1, 10)]

    def tabla(data, widths, total_row=False):
        t = Table(data, colWidths=widths, repeatRows=1)
        sty = [("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e5f")),
               ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
               ("FONTSIZE", (0, 0), (-1, -1), 8),
               ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
               ("VALIGN", (0, 0), (-1, -1), "TOP")]
        if total_row:
            sty += [("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
                    ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#e8f0f2"))]
        t.setStyle(TableStyle(sty))
        return t

    # Nuevas facturas agrupadas por mes
    el.append(Paragraph("Facturas nuevas agregadas hoy (por mes)", st["Heading2"]))
    if not nuevos:
        el.append(Paragraph("No se encontraron facturas nuevas.", st["Normal"]))
    grupos = defaultdict(list)
    for d in nuevos:
        y, m = int(d["fecha"][:4]), int(d["fecha"][5:7])
        grupos[(y, m)].append(d)
    for k in sorted(grupos):
        el.append(Paragraph(f"{MESES[k[1]]} {k[0]}", st["Heading3"]))
        data = [["Fecha", "Proveedor", "Descripción", "Categoría", "Nº", "Moneda", "Total", "USD"]]
        sub = 0
        for d in grupos[k]:
            sub += d["_usd"]
            data.append([d["fecha"], Paragraph(str(d.get("proveedor")), small),
                         Paragraph(str(d.get("descripcion")), small), Paragraph(d["categoria"], small),
                         str(d.get("consec_corto")), d.get("moneda"), f"{float(d['total']):,.2f}",
                         f"{d['_usd']:,.2f}"])
        data.append(["", "", "", "Subtotal", "", "", "", f"{sub:,.2f}"])
        el.append(tabla(data, [60, 140, 150, 130, 70, 45, 70, 60], total_row=True))
        el.append(Spacer(1, 8))

    # Resumen acumulado por mes y categoría
    el.append(Paragraph("Gasto acumulado por mes y categoría (USD, hoja Gastos Diarios)", st["Heading2"]))
    meses = sorted(por_mes)
    allcats = sorted({c for m in meses for c in por_mes[m]})
    data = [["Categoría"] + [f"{MESES[m][:3]} {y}" for y, m in meses] + ["Total"]]
    for c in allcats:
        fila = [por_mes[m].get(c, 0.0) for m in meses]
        data.append([c] + [f"{v:,.2f}" for v in fila] + [f"{sum(fila):,.2f}"])
    tot = [sum(por_mes[m].values()) for m in meses]
    data.append(["TOTAL"] + [f"{v:,.2f}" for v in tot] + [f"{sum(tot):,.2f}"])
    el.append(tabla(data, [170] + [80] * (len(meses) + 1), total_row=True))
    el.append(Paragraph("Montos en colones convertidos con el tipo de cambio de cada fila "
                        f"(por defecto ₡{tc_default:g} por $ cuando la factura no lo indica).", small))

    if revisar:
        el.append(Paragraph("Para revisar manualmente", st["Heading2"]))
        data = [["Fecha", "Proveedor", "Total", "Nota"]]
        for d in revisar:
            data.append([d.get("fecha"), Paragraph(str(d.get("proveedor")), small), str(d.get("total")),
                         Paragraph(str(d.get("nota")), small)])
        el.append(tabla(data, [70, 170, 80, 400]))
    doc.build(el)

    resumen = {"nuevos": len(nuevos), "omitidos": len(omitidos), "revisar": len(revisar),
               "usd_nuevos": round(sum(d["_usd"] for d in nuevos), 2),
               "por_mes_nuevos": {f"{MESES[k[1]]} {k[0]}": [len(v), round(sum(x["_usd"] for x in v), 2)]
                                  for k, v in sorted(grupos.items())},
               "xlsx": xlsx_out, "pdf": pdf_out, "omitidos_detalle": omitidos}
    print(json.dumps(resumen, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
