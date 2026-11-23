#!/usr/bin/env python3
"""Procesa facturas descargadas del correo y genera el reporte mensual.

Entrada:  facturas_staging/raw/<messageId>__<archivo>  (PDF y XML de Hacienda)
          facturas_staging/manifest_*.tsv              (metadatos del correo)
Salida:   facturas/AAAA-MM/<fecha>_<proveedor>_<consecutivo>.pdf
          facturas/registro_facturas.json              (estado, para detectar nuevas)
          reportes/Facturas_Canavida_<fecha>.xlsx      (formato hoja "Gastos Diarios", una hoja por mes)
          reportes/Facturas_Canavida_<fecha>.pdf       (resumen + todas las facturas, ordenadas por mes)

Uso: python3 tools/procesar_facturas.py [--tipo-cambio 459] [--fecha 2026-11-23]
"""
import argparse, csv, datetime as dt, glob, json, os, re, shutil, subprocess, sys, unicodedata
import xml.etree.ElementTree as ET

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pypdf import PdfReader, PdfWriter
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STAGING = os.path.join(ROOT, "facturas_staging")
OUT_FACT = os.path.join(ROOT, "facturas")
OUT_REP = os.path.join(ROOT, "reportes")
TEMPLATE = os.path.join(ROOT, "Copy of Canavida Gastos Ingresos Oficial (2).xlsx")
REGISTRO = os.path.join(OUT_FACT, "registro_facturas.json")

MESES = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto",
         "Septiembre", "Octubre", "Noviembre", "Diciembre"]
VERDE = "FF1F5C4A"

# Proveedor (palabra clave, sin tildes, minúsculas) -> categoría de la hoja Listas
CATEGORIAS = [
    (r"supermercado|ekono|walmart|pali|mas x menos|maxi|auto ?mercado|carniceria|panaderia|super |megasuper|coopeagri.*super|fruta|verdura|pequeno mundo", "Alimentos y bebidas"),
    (r"ferreteria|materiales|coopemadereros|deposito|construc|el lagar|epa\b|ferre|maderas|pintura|sur quimica|electric", "Mantenimiento y reparaciones"),
    (r"gasolin|servicentro|combustible|coopeagri|estacion|uno\b|delta|motochopper|llanta|repuesto|taller|automotr", "Combustible y transporte"),
    (r"kolbi|ice\b|telefon|tel:|claro|liberty|movistar|internet|wireless|cnfl|coopeguanacaste|electricidad|agua|aya\b|asada|acueducto|direccion de agua", "Servicios públicos e internet"),
    (r"accounting|contab|abogad|notari|legal|consult|crc accounting", "Imprevistos y otros"),
    (r"pricelabs|lodgify|software|google|microsoft|quo\b|openphone|zoom|canva|adobe|dropbox|apple", "Software y tecnología"),
    (r"airbnb|booking|expedia|vrbo|comision", "Comisiones de plataformas"),
    (r"lavander|limpieza|clean|quimicos", "Limpieza y lavandería"),
    (r"piscina|pool|jardin|vivero|cloro", "Jardinería y piscinas"),
    (r"seguro|\bins\b|aseguradora", "Seguros"),
    (r"ccss|caja costarricense", "Cargas sociales (CCSS/INS)"),
    (r"hacienda|municipalidad|patente|impuesto", "Impuestos y patentes"),
    (r"mueble|colchon|tapic", "Muebleria"),
    (r"marketing|publicidad|facebook|meta platforms|instagram", "Marketing y publicidad"),
    (r"tour|excursion|catamaran|kayak", "Costo de tours y actividades"),
    (r"amenit|hotel supply|blancos|toalla|sabanas", "Suministros y amenities"),
    (r"medic|clinica|farmacia|hospital", "Imprevistos y otros"),
]
# Código MedioPago de Hacienda -> método de pago de la hoja Listas (cheque/terceros/otros quedan vacíos)
MEDIO_PAGO = {"01": "Efectivo", "02": "Tarjeta", "03": "", "04": "Transferencia / SINPE",
              "05": "", "06": "Transferencia / SINPE", "99": ""}


def sin_tildes(s):
    return "".join(c for c in unicodedata.normalize("NFD", s or "") if unicodedata.category(c) != "Mn").lower()


def categoria(proveedor, descripcion, asunto):
    txt = sin_tildes(" ".join([proveedor or "", descripcion or "", asunto or ""]))
    for pat, cat in CATEGORIAS:
        if re.search(pat, txt):
            return cat
    return "Imprevistos y otros"


def local(tag):
    return tag.rsplit("}", 1)[-1]


def find(el, *path):
    """Busca por nombre local, ignorando namespaces."""
    cur = [el]
    for p in path:
        nxt = []
        for c in cur:
            nxt += [x for x in c.iter() if local(x.tag) == p and x is not c]
        if not nxt:
            return None
        cur = nxt[:1]
    return cur[0]


def txt(el, *path):
    x = find(el, *path)
    return x.text.strip() if x is not None and x.text else ""


def num(s):
    try:
        return float(str(s).replace(",", ""))
    except Exception:
        return None


def leer_xml(path):
    try:
        root = ET.parse(path).getroot()
    except Exception:
        return None
    tipo = local(root.tag)
    if not re.match(r"(FacturaElectronica|TiqueteElectronico|NotaCreditoElectronica|NotaDebitoElectronica|FacturaElectronicaCompra)", tipo):
        return None
    det = [txt(l, "Detalle") for l in root.iter() if local(l.tag) == "LineaDetalle"]
    det = [d for d in det if d]
    descr = det[0] if det else ""
    if len(det) > 1:
        descr += f" (+{len(det)-1} líneas)"
    fecha = txt(root, "FechaEmision")[:10]
    medio = ""
    for x in root.iter():
        if local(x.tag) in ("MedioPago", "TipoMedioPago") and x.text and x.text.strip() in MEDIO_PAGO:
            medio = MEDIO_PAGO[x.text.strip()]
            break
    total = num(txt(root, "ResumenFactura", "TotalComprobante"))
    if tipo.startswith("NotaCredito") and total:
        total = -total
    return {
        "tipo_doc": tipo,
        "clave": txt(root, "Clave"),
        "consecutivo": txt(root, "NumeroConsecutivo"),
        "fecha": fecha,
        "proveedor": txt(root, "Emisor", "Nombre"),
        "proveedor_nombre_comercial": txt(root, "Emisor", "NombreComercial"),
        "receptor": txt(root, "Receptor", "Nombre"),
        "moneda": txt(root, "ResumenFactura", "CodigoMoneda") or "CRC",
        "tipo_cambio_doc": num(txt(root, "ResumenFactura", "TipoCambio")),
        "total": total,
        "iva": num(txt(root, "ResumenFactura", "TotalImpuesto")),
        "descripcion": descr,
        "metodo_pago": medio,
        "fuente": "XML Hacienda",
    }


FECHA_RE = re.compile(r"(\d{1,2})[/-](\d{1,2})[/-](20\d\d)")
ISO_RE = re.compile(r"(20\d\d)-(\d\d)-(\d\d)")
MONTO = r"(?:₡|¢|\$|CRC|USD|US\$)?\s*([0-9]{1,3}(?:[.,\s][0-9]{3})*(?:[.,][0-9]{2})?)"


def parse_monto(s):
    s = s.strip().replace(" ", "")
    if re.search(r",\d{2}$", s) and "." in s:     # 1.234,56
        s = s.replace(".", "").replace(",", ".")
    elif re.search(r",\d{2}$", s):                 # 1234,56
        s = s.replace(",", ".")
    else:
        s = s.replace(",", "")
    return num(s)


def leer_pdf(path):
    try:
        t = subprocess.run(["pdftotext", "-layout", path, "-"], capture_output=True, text=True, timeout=60).stdout
    except Exception:
        t = ""
    if not t.strip():
        return {"fuente": "PDF (sin texto)"}
    d = {"fuente": "PDF"}
    up = sin_tildes(t).upper()
    if not re.search(r"FACTURA|INVOICE|RECEIPT|RECIBO|COMPROBANTE|TIQUETE|TOTAL", up):
        return {"no_factura": True}
    m = re.search(r"\b(\d{20})\b", t)
    if m:
        d["consecutivo"] = m.group(1)
    m = re.search(r"\b(506\d{47})\b", t)
    if m:
        d["clave"] = m.group(1)
    m = FECHA_RE.search(t)
    if m:
        dd, mm, yy = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if mm > 12:
            dd, mm = mm, dd
        try:
            d["fecha"] = dt.date(yy, mm, dd).isoformat()
        except ValueError:
            pass
    elif ISO_RE.search(t):
        d["fecha"] = "-".join(ISO_RE.search(t).groups())
    if "USD" in up or "DOLAR" in up or "US$" in t:
        d["moneda"] = "USD" if not ("₡" in t or "CRC" in up or "COLONES" in up) else ("USD" if re.search(r"MONEDA:?\s*USD|DOLARES", up) else "CRC")
    else:
        d["moneda"] = "CRC" if ("₡" in t or "CRC" in up or "COLON" in up) else "USD"
    best = None
    for line in t.splitlines():
        L = sin_tildes(line).upper()
        if re.search(r"TOTAL\s*(COMPROBANTE|FACTURA|A PAGAR|GENERAL|DOCUMENTO|:)|\bTOTAL\b\s*[₡$¢]|^\s*TOTAL\s", L) and not re.search(r"SUBTOTAL|IVA|IMPUESTO|DESCUENTO|GRAVAD|EXENT|EXONER|NO SUJETO", L):
            ms = re.findall(MONTO, line)
            if ms:
                v = parse_monto(ms[-1])
                if v and (best is None or v >= best):
                    best = v
    d["total"] = best
    lines = [l.strip() for l in t.splitlines() if l.strip()]
    d["proveedor"] = re.split(r"\s{2,}|\|", lines[0])[0] if lines else ""
    return d


def cargar_manifest():
    meta = {}
    for f in sorted(glob.glob(os.path.join(STAGING, "manifest_*.tsv"))):
        with open(f, encoding="utf-8", errors="replace") as fh:
            for row in csv.reader(fh, delimiter="\t"):
                if len(row) >= 5:
                    meta[row[1]] = {"thread": row[0], "fecha_correo": row[2][:10], "asunto": row[3], "remitente": row[4]}
    return meta


def slug(s, n=30):
    s = re.sub(r"[^A-Za-z0-9]+", "-", sin_tildes(s)).strip("-")
    return (s[:n] or "proveedor").strip("-")


def recolectar(tc_default):
    meta = cargar_manifest()
    raw = sorted(glob.glob(os.path.join(STAGING, "raw", "*")))
    por_msg = {}
    for p in raw:
        mid, _, name = os.path.basename(p).partition("__")
        por_msg.setdefault(mid, []).append((name, p))
    facturas = []
    for mid, files in por_msg.items():
        pdfs = [p for n, p in files if n.lower().endswith(".pdf")]
        if not pdfs:
            continue
        xmls = [leer_xml(p) for n, p in files if n.lower().endswith(".xml")]
        xmls = [x for x in xmls if x and "canavida" not in sin_tildes(x.get("proveedor", ""))]
        m = meta.get(mid, {})
        if xmls:
            # Un XML por factura; empareja cada XML con el PDF que contiene su clave/consecutivo.
            for x in xmls:
                pdf = next((p for p in pdfs if x["clave"] and x["clave"] in os.path.basename(p)), None) \
                   or next((p for p in pdfs if x["consecutivo"] and x["consecutivo"] in os.path.basename(p)), None) \
                   or (pdfs[0] if len(xmls) == 1 else None)
                if pdf is None:
                    for p in pdfs:
                        tx = subprocess.run(["pdftotext", p, "-"], capture_output=True, text=True).stdout
                        if x["consecutivo"] and x["consecutivo"] in tx:
                            pdf = p
                            break
                if pdf is None:
                    pdf = pdfs[0]
                facturas.append(dict(x, pdf=pdf, message_id=mid, **m))
        else:
            for p in pdfs:
                d = leer_pdf(p)
                if d.get("no_factura"):
                    print("Omitido (no parece factura):", os.path.basename(p), file=sys.stderr)
                    continue
                d.setdefault("fecha", m.get("fecha_correo", ""))
                if not d.get("proveedor") or len(d["proveedor"]) < 4 or re.match(r"(?i)(clave|invoice|factura|recibo|receipt|fecha|page|pagina)\b", d["proveedor"]):
                    d["proveedor"] = m.get("remitente", "")
                facturas.append(dict(d, pdf=p, message_id=mid, descripcion=m.get("asunto", ""), metodo_pago="", **m))
    # Elimina duplicados (misma clave o mismo consecutivo+proveedor)
    vistos, unicas = set(), []
    for f in facturas:
        if f.get("clave"):
            k = f["clave"]
        elif f.get("consecutivo"):
            k = (f["consecutivo"], f.get("proveedor"), f.get("total"))
        else:
            k = ("pdf", os.path.basename(f["pdf"]).split("__", 1)[-1], f.get("total"))
        if k in vistos:
            continue
        vistos.add(k)
        unicas.append(f)
    for f in unicas:
        f["fecha"] = f.get("fecha") or f.get("fecha_correo") or ""
        f["moneda"] = (f.get("moneda") or "CRC").upper()
        tc = f.get("tipo_cambio_doc") if f["moneda"] == "CRC" and (f.get("tipo_cambio_doc") or 0) > 100 else tc_default
        f["tipo_cambio"] = tc
        tot = f.get("total")
        if tot is None:
            f["monto_usd"] = None
        elif f["moneda"] == "USD":
            f["monto_usd"] = round(tot, 2)
        elif f["moneda"] == "CRC":
            f["monto_usd"] = round(tot / tc, 2)
        else:
            f["monto_usd"] = round(tot * (f.get("tipo_cambio_doc") or 1) / tc_default, 2)
        prov = f.get("proveedor_nombre_comercial") or f.get("proveedor") or ""
        f["proveedor_mostrar"] = prov
        f["categoria"] = categoria(prov + " " + (f.get("proveedor") or ""), f.get("descripcion"), f.get("asunto"))
        f["mes"] = f["fecha"][:7] if re.match(r"20\d\d-\d\d", f["fecha"]) else "sin-fecha"
        f["revisar"] = "; ".join(x for x in [
            "monto no leído" if tot is None else "",
            "fecha no leída" if f["mes"] == "sin-fecha" else "",
            "solo PDF (verificar)" if f.get("fuente", "").startswith("PDF") else "",
        ] if x)
    unicas.sort(key=lambda f: (f["fecha"], f["proveedor_mostrar"]))
    return unicas


def organizar_pdfs(facturas):
    for f in facturas:
        carpeta = os.path.join(OUT_FACT, f["mes"])
        os.makedirs(carpeta, exist_ok=True)
        cons = f.get("consecutivo") or f["message_id"]
        nombre = f"{f['fecha'] or 'sin-fecha'}_{slug(f['proveedor_mostrar'])}_{cons[-10:]}.pdf"
        dest = os.path.join(carpeta, nombre)
        if not os.path.exists(dest):
            shutil.copy2(f["pdf"], dest)
        f["archivo"] = os.path.relpath(dest, ROOT)


def ya_en_gastos_diarios():
    """Nº de factura ya anotados a mano en la hoja Gastos Diarios del archivo oficial."""
    try:
        wb = openpyxl.load_workbook(TEMPLATE, read_only=True, data_only=True)
        ws = wb["Gastos Diarios"]
        s = set()
        for row in ws.iter_rows(min_row=6, max_col=7, values_only=True):
            if row[0] and row[6] not in (None, ""):
                try:
                    s.add(int(float(str(row[6]).strip())))
                except ValueError:
                    pass
        return s
    except Exception:
        return set()


def nombre_mes(m):
    if m == "sin-fecha":
        return "Sin fecha"
    y, mm = m.split("-")
    return f"{MESES[int(mm)]} {y}"


def excel(facturas, fecha_rep, nuevas_ids, path):
    registradas = ya_en_gastos_diarios()
    wb = openpyxl.Workbook()
    res = wb.active
    res.title = "Resumen"
    thin = Side(style="thin", color="FFD9D9D9")
    head_font = Font(bold=True, color="FFFFFFFF")
    head_fill = PatternFill("solid", fgColor=VERDE)
    meses = sorted({f["mes"] for f in facturas})

    for mes in meses:
        ws = wb.create_sheet(nombre_mes(mes)[:31])
        ws["A1"] = f"Facturas recibidas por correo – {nombre_mes(mes)}"
        ws["A1"].font = Font(bold=True, size=16, color=VERDE)
        ws["A2"] = ("Mismo formato que la hoja 'Gastos Diarios'. Montos en CRC convertidos a USD con el tipo de cambio de la columna K. "
                    "Columna M indica si el Nº de factura ya aparece en 'Gastos Diarios' (para no registrarla dos veces).")
        ws["A2"].font = Font(italic=True, size=9, color="FF595959")
        cols = ["Fecha", "Categoría", "Descripción", "Proveedor", "Método de pago", "Monto (USD)", "Nº factura / Notas",
                "Moneda original", "Monto original", "IVA original", "Tipo de cambio", "Archivo PDF", "¿Ya en Gastos Diarios?", "Nueva hoy", "Revisar"]
        widths = [13, 30, 42, 33, 22, 15, 24, 10, 16, 14, 11, 52, 12, 9, 22]
        for i, (c, w) in enumerate(zip(cols, widths), 1):
            cell = ws.cell(5, i, c)
            cell.font, cell.fill = head_font, head_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(i)].width = w
        r = 6
        filas = [f for f in facturas if f["mes"] == mes]
        for f in filas:
            cons = f.get("consecutivo") or ""
            try:
                corto = int(cons[-10:]) if cons else None
            except ValueError:
                corto = None
            fecha = dt.date.fromisoformat(f["fecha"]) if re.match(r"\d{4}-\d\d-\d\d$", f["fecha"]) else f["fecha"]
            vals = [fecha, f["categoria"], f.get("descripcion", "")[:200], f["proveedor_mostrar"], f.get("metodo_pago", ""),
                    f["monto_usd"], f"Fact. {cons}" if cons else "", f["moneda"], f.get("total"), f.get("iva"),
                    f["tipo_cambio"] if f["moneda"] != "USD" else 1, f.get("archivo", ""),
                    "Sí" if corto is not None and corto in registradas else "No",
                    "Sí" if f["message_id"] in nuevas_ids else "", f["revisar"]]
            for i, v in enumerate(vals, 1):
                c = ws.cell(r, i, v)
                c.border = Border(bottom=thin)
            ws.cell(r, 1).number_format = "d/M/yyyy"
            ws.cell(r, 6).number_format = '$#,##0.00;"($"#,##0.00);-'
            ws.cell(r, 9).number_format = "#,##0.00"
            ws.cell(r, 10).number_format = "#,##0.00"
            if f["revisar"]:
                ws.cell(r, 15).font = Font(color="FFC00000")
            r += 1
        ws["A3"], ws["B3"] = "Facturas:", len(filas)
        ws["E3"], ws["F3"] = "TOTAL USD", f"=SUM(F6:F{max(r-1,6)})"
        for c in ("A3", "E3"):
            ws[c].font = Font(bold=True)
        ws["F3"].number_format = '$#,##0.00'
        ws["F3"].font = Font(bold=True, color=VERDE)
        ws.freeze_panes = "A6"
        ws.auto_filter.ref = f"A5:O{max(r-1,5)}"

    # Resumen por categoría y mes (fórmulas sobre las hojas de cada mes)
    res["A1"] = "Canavida Villas & Resort – Facturas recibidas por correo"
    res["A1"].font = Font(bold=True, size=16, color=VERDE)
    res["A2"] = f"Actualizado: {fecha_rep}. Fuente: buzón admicanavida@gmail.com (adjuntos PDF/XML de facturas). Moneda: USD."
    res["A2"].font = Font(italic=True, size=9, color="FF595959")
    res.cell(4, 1, "Categoría")
    for j, mes in enumerate(meses, 2):
        res.cell(4, j, nombre_mes(mes))
    res.cell(4, len(meses) + 2, "Total")
    for j in range(1, len(meses) + 3):
        c = res.cell(4, j)
        c.font, c.fill = head_font, head_fill
        c.alignment = Alignment(horizontal="center")
    cats = sorted({f["categoria"] for f in facturas})
    for i, cat in enumerate(cats, 5):
        res.cell(i, 1, cat)
        for j, mes in enumerate(meses, 2):
            sh = nombre_mes(mes)[:31]
            res.cell(i, j, f"=SUMIFS('{sh}'!$F$6:$F$2000,'{sh}'!$B$6:$B$2000,$A{i})").number_format = '$#,##0.00'
        L = get_column_letter(len(meses) + 1)
        res.cell(i, len(meses) + 2, f"=SUM(B{i}:{L}{i})").number_format = '$#,##0.00'
    tr = 5 + len(cats)
    res.cell(tr, 1, "TOTAL").font = Font(bold=True)
    for j in range(2, len(meses) + 3):
        L = get_column_letter(j)
        c = res.cell(tr, j, f"=SUM({L}5:{L}{tr-1})")
        c.number_format, c.font = '$#,##0.00', Font(bold=True)
    res.cell(tr + 1, 1, "Nº de facturas")
    for j, mes in enumerate(meses, 2):
        res.cell(tr + 1, j, sum(1 for f in facturas if f["mes"] == mes))
    res.cell(tr + 1, len(meses) + 2, len(facturas))
    res.column_dimensions["A"].width = 34
    for j in range(2, len(meses) + 3):
        res.column_dimensions[get_column_letter(j)].width = 18
    wb.save(path)


def pdf(facturas, fecha_rep, nuevas_ids, path, tc):
    tmp = path + ".resumen.pdf"
    ss = getSampleStyleSheet()
    h = ParagraphStyle("h", parent=ss["Title"], textColor=colors.HexColor("#1F5C4A"))
    small = ParagraphStyle("s", parent=ss["Normal"], fontSize=7, leading=8.5)
    doc = SimpleDocTemplate(tmp, pagesize=landscape(letter), leftMargin=28, rightMargin=28, topMargin=28, bottomMargin=28)
    el = [Paragraph("Canavida – Facturas recibidas por correo", h),
          Paragraph(f"Actualizado {fecha_rep} · {len(facturas)} facturas · nuevas hoy: {len(nuevas_ids)} · "
                    f"tipo de cambio por defecto ₡{tc:g}/$", ss["Normal"]), Spacer(1, 10)]
    meses = sorted({f["mes"] for f in facturas})
    resumen = [["Mes", "Facturas", "Total USD"]]
    for m in meses:
        fs = [f for f in facturas if f["mes"] == m]
        resumen.append([nombre_mes(m), str(len(fs)), f"${sum(f['monto_usd'] or 0 for f in fs):,.2f}"])
    t = Table(resumen, colWidths=[160, 70, 100])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F5C4A")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                           ("ALIGN", (1, 0), (-1, -1), "RIGHT"), ("GRID", (0, 0), (-1, -1), 0.25, colors.lightgrey)]))
    el += [t, Spacer(1, 12)]
    for m in meses:
        fs = [f for f in facturas if f["mes"] == m]
        el.append(Paragraph(nombre_mes(m), ss["Heading2"]))
        data = [["Fecha", "Proveedor", "Categoría", "Descripción", "Nº factura", "Moneda", "Monto orig.", "USD", "Nueva"]]
        for f in fs:
            data.append([f["fecha"], Paragraph(f["proveedor_mostrar"][:45], small), f["categoria"], Paragraph((f.get("descripcion") or "")[:70], small),
                         (f.get("consecutivo") or "")[-10:], f["moneda"],
                         f"{f['total']:,.2f}" if f.get("total") is not None else "?",
                         f"${f['monto_usd']:,.2f}" if f.get("monto_usd") is not None else "?",
                         "●" if f["message_id"] in nuevas_ids else ""])
        data.append(["", "", "", "", "", "", "TOTAL", f"${sum(f['monto_usd'] or 0 for f in fs):,.2f}", ""])
        t = Table(data, colWidths=[56, 130, 110, 170, 62, 40, 70, 62, 30], repeatRows=1)
        t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F5C4A")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                               ("FONTSIZE", (0, 0), (-1, -1), 7), ("ALIGN", (6, 0), (7, -1), "RIGHT"), ("VALIGN", (0, 0), (-1, -1), "TOP"),
                               ("GRID", (0, 0), (-1, -1), 0.25, colors.lightgrey), ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold")]))
        el += [t, Spacer(1, 10)]
    doc.build(el)

    w = PdfWriter()
    w.append(PdfReader(tmp), outline_item="Resumen")
    for m in meses:
        parent = None
        for f in [f for f in facturas if f["mes"] == m]:
            try:
                r = PdfReader(f["pdf"])
                start = len(w.pages)
                w.append(r)
                if parent is None:
                    parent = w.add_outline_item(nombre_mes(m), start)
                w.add_outline_item(f"{f['fecha']} {f['proveedor_mostrar'][:40]}", start, parent=parent)
            except Exception as e:
                print("No se pudo anexar", f["pdf"], e, file=sys.stderr)
    w.compress_identical_objects() if hasattr(w, "compress_identical_objects") else None
    with open(path, "wb") as fh:
        w.write(fh)
    os.remove(tmp)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tipo-cambio", type=float, default=459.0)
    ap.add_argument("--fecha", default=dt.date.today().isoformat())
    a = ap.parse_args()
    facturas = recolectar(a.tipo_cambio)
    organizar_pdfs(facturas)
    prev = {}
    if os.path.exists(REGISTRO):
        prev = json.load(open(REGISTRO))
    previos = set(prev.get("message_ids", []))
    nuevas = {f["message_id"] for f in facturas} - previos
    os.makedirs(OUT_REP, exist_ok=True)
    base = os.path.join(OUT_REP, f"Facturas_Canavida_{a.fecha}")
    excel(facturas, a.fecha, nuevas, base + ".xlsx")
    pdf(facturas, a.fecha, nuevas, base + ".pdf", a.tipo_cambio)
    keep = ["fecha", "mes", "proveedor_mostrar", "categoria", "consecutivo", "clave", "moneda", "total", "monto_usd",
            "tipo_cambio", "archivo", "message_id", "asunto", "fuente", "revisar"]
    json.dump({"actualizado": a.fecha, "message_ids": sorted(previos | {f["message_id"] for f in facturas}),
               "facturas": [{k: f.get(k) for k in keep} for f in facturas]},
              open(REGISTRO, "w"), ensure_ascii=False, indent=1)
    print(json.dumps({"facturas": len(facturas), "nuevas": len(nuevas), "excel": base + ".xlsx", "pdf": base + ".pdf",
                      "por_mes": {m: round(sum(f["monto_usd"] or 0 for f in facturas if f["mes"] == m), 2)
                                  for m in sorted({f["mes"] for f in facturas})},
                      "revisar": [(f["fecha"], f["proveedor_mostrar"], f["revisar"]) for f in facturas if f["revisar"]]},
                     ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
