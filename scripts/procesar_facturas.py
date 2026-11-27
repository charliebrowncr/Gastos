#!/usr/bin/env python3
"""Procesa facturas descargadas de Gmail (PDF + XML de Hacienda CR) y genera:
  - Excel actualizado con el formato de la plantilla (hoja 'Gastos Diarios' + hoja 'Facturas por Mes')
  - PDF con resumen por mes seguido de las facturas ordenadas por mes

Uso: python3 procesar_facturas.py <dir_descargas> <plantilla.xlsx> <dir_salida> [--tc 459]
El directorio de descargas contiene <msgid>__<archivo>.pdf/.xml y meta_*.jsonl.
"""
import sys, os, re, json, glob, subprocess, datetime as dt, argparse, copy
import xml.etree.ElementTree as ET
from collections import defaultdict, OrderedDict

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from pypdf import PdfWriter, PdfReader
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak

CANAVIDA_IDS = {"3102777726"}
MESES = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto",
         "Septiembre", "Octubre", "Noviembre", "Diciembre"]
MEDIO = {"01": "Efectivo", "02": "Tarjeta", "03": "Transferencia / SINPE", "04": "Transferencia / SINPE",
         "05": "Transferencia / SINPE", "06": "Transferencia / SINPE", "07": "Transferencia / SINPE", "99": "Transferencia / SINPE"}

# (palabras clave en emisor/descripción) -> categoría de la plantilla
CATEGORIAS = [
    (r"supermercado|walmart|maxi pali|masxmenos|\bpali\b|automercado|bm |fruta|verdura|carnicer|panader|restaur|soda|burrito|mariscos|pizza|cafe|coffee|licor|spirits|wine|vino|cabernet|sauvignon|bebida|alimento|harina|café del valle", "Alimentos y bebidas"),
    (r"gasolin|combustible|servicentro|coopeagri|estacion de servicio|\bgs\b|taller|mecanic|repuesto|llanta|transporte|uber", "Combustible y transporte"),
    (r"ferreter|fixsur|material|construc|acarreo|agregado|piedra|arena|silicon|bisagra|colono|madera|coopemadereros|el lagar|\bepa\b|pintura|thinner|primer|diluyente|lijable|puerta|ventana|reparaci|mantenimiento|bambu|bambú|electric|plomer|cemento|vidrio", "Mantenimiento y reparaciones"),
    (r"ministerio de ambiente|direcci[oó]n de agua|\bice\b|instituto costarricense de electricidad|kolbi|telefon|celular|internet|cnfl|aya|acueducto|agua|energ|electricidad|liberty|claro|tigo|wireless", "Servicios públicos e internet"),
    (r"instituto nacional de seguros|\bins\b|seguro|poliza|póliza|assa|mapfre", "Seguros"),
    (r"ccss|caja costarricense", "Cargas sociales (CCSS/INS)"),
    (r"contador|accounting|contab|abogad|legal|notari|medic|farmac|clinic|urgent|hospital", "Imprevistos y otros"),
    (r"lavander|limpieza|cleaner|detergente|cloro", "Limpieza y lavandería"),
    (r"piscina|pool|jardin|vivero|botica|agro", "Jardinería y piscinas"),
    (r"software|tecnolog|smart506|computo|computadora|electronic|celular|apple|google|microsoft|pricelabs|breezeway|lodgify|\bquo\b|openphone|nomina|nómina|bit cloud", "Software y tecnología"),
    (r"mueble|muebler|colchon|colchón", "Muebleria"),
    (r"booking|airbnb|expedia|comision", "Comisiones de plataformas"),
    (r"municipal|hacienda|patente|impuesto", "Impuestos y patentes"),
    (r"publicidad|marketing|imprenta|rotulo", "Marketing y publicidad"),
    (r"tour|excursion|catamaran", "Costo de tours y actividades"),
    (r"amenit|suministro|bazar|libreria|librería|oficina|ekono|cemaco|pequeño mundo|universal", "Suministros y amenities"),
]


def lname(tag):
    return tag.rsplit("}", 1)[-1]


def find(root, *path):
    """Busca por nombres locales ignorando namespaces."""
    nodes = [root]
    for p in path:
        nxt = []
        for n in nodes:
            nxt += [c for c in n if lname(c.tag) == p]
        nodes = nxt
    return nodes


def text(root, *path):
    n = find(root, *path)
    return n[0].text.strip() if n and n[0].text else None


def parse_xml(path):
    try:
        root = ET.parse(path).getroot()
    except Exception:
        return None
    tipo = lname(root.tag)
    if tipo not in ("FacturaElectronica", "TiqueteElectronico", "NotaCreditoElectronica",
                    "NotaDebitoElectronica", "FacturaElectronicaCompra", "FacturaElectronicaExportacion"):
        return None
    d = {"tipo": tipo, "clave": text(root, "Clave"), "consecutivo": text(root, "NumeroConsecutivo")}
    f = text(root, "FechaEmision") or ""
    try:
        d["fecha"] = dt.datetime.fromisoformat(f[:19]).date()
    except Exception:
        d["fecha"] = None
    d["emisor"] = text(root, "Emisor", "NombreComercial") or text(root, "Emisor", "Nombre")
    d["emisor_legal"] = text(root, "Emisor", "Nombre")
    d["emisor_id"] = text(root, "Emisor", "Identificacion", "Numero")
    d["receptor_id"] = text(root, "Receptor", "Identificacion", "Numero")
    d["moneda"] = text(root, "ResumenFactura", "CodigoTipoMoneda", "CodigoMoneda") or "CRC"
    tc = text(root, "ResumenFactura", "CodigoTipoMoneda", "TipoCambio")
    d["tipo_cambio"] = float(tc) if tc else None
    tot = text(root, "ResumenFactura", "TotalComprobante")
    d["total"] = float(tot) if tot else None
    mp = text(root, "MedioPago") or text(root, "ResumenFactura", "MedioPago", "TipoMedioPago")
    d["medio_pago"] = MEDIO.get(mp or "", "Tarjeta" if mp is None else "Transferencia / SINPE")
    dets = [n.text.strip() for n in find(root, "DetalleServicio", "LineaDetalle", "Detalle") if n.text]
    d["detalle"] = "; ".join(dets[:3])[:120]
    return d


AMT = r"([\d]{1,3}(?:[.,\s]\d{3})*(?:[.,]\d{2})|\d+(?:[.,]\d{2}))"


def to_num(s):
    s = s.replace(" ", "")
    if re.search(r",\d{2}$", s):
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def parse_pdf(path):
    try:
        t = subprocess.run(["pdftotext", "-layout", path, "-"], capture_output=True, text=True, timeout=60).stdout
    except Exception:
        t = ""
    d = {"texto": t}
    m = re.search(r"(?<!\d)(506\d{47})(?!\d)", t) or re.search(r"(?<!\d)(506\d{47})(?!\d)", os.path.basename(path))
    if m:
        d["clave"] = m.group(1)
    tot = None
    for pat in [r"TOTAL\s+COMPROBANTE[^\d\n]*" + AMT, r"TOTAL\s+(?:A\s+PAGAR|FACTURA|GENERAL)[^\d\n]*" + AMT,
                r"\bTOTAL\b[^\d\n]{0,60}" + AMT]:
        ms = re.findall(pat, t, re.I)
        if ms:
            tot = to_num(ms[-1])
            break
    d["total"] = tot
    usd_ = re.search(r"\bUSD\s?\$?\s?\d|US\$|(?<![A-Za-z¢])\$\s?\d|Moneda:?\s*USD|D[oó]lares", t)
    crc_ = re.search(r"¢|₡|\bCRC\s?\d|Moneda:?\s*CRC|colones", t, re.I)
    d["moneda"] = "USD" if usd_ and not crc_ else "CRC"
    d["prefactura"] = bool(re.search(r"pre-?\s?factura|proforma", t[:1500], re.I))
    lineas = [l.strip() for l in t.splitlines() if len(l.strip()) > 3]
    d["primera_linea"] = re.split(r"\s{3,}", lineas[0])[0][:60] if lineas else None
    m = re.search(r"Fecha(?:\s+de)?\s+emisi[oó]n[:\s]*(\d{1,2})[/-](\d{1,2})[/-](\d{4})", t, re.I)
    if m:
        try:
            d["fecha"] = dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            pass
    m = re.search(r"Medio\s+(?:de\s+)?Pago[:\s]*(\d{2})?\s*([A-Za-zé]+)?", t, re.I)
    if m:
        w = (m.group(2) or "").lower()
        d["medio_pago"] = MEDIO.get(m.group(1) or "", "Efectivo" if "efect" in w else
                                    "Tarjeta" if "tarj" in w else "Transferencia / SINPE" if w else None)
    return d


def clave_info(clave):
    """Clave numérica de 50 dígitos: 506 DDMMAA cédula(12) consecutivo(20) situación código(8)."""
    if not clave or len(clave) != 50:
        return {}
    try:
        fecha = dt.date(2000 + int(clave[7:9]), int(clave[5:7]), int(clave[3:5]))
    except ValueError:
        fecha = None
    return {"fecha": fecha, "emisor_id": clave[9:21].lstrip("0"), "consecutivo": clave[21:41]}


def categoria(*txts):
    """Prueba cada texto por orden de confianza (proveedor, razón social, detalle, texto PDF)."""
    for t in txts:
        s = (t or "").lower()
        for pat, cat in CATEGORIAS:
            if s and re.search(pat, s):
                return cat
    return "Imprevistos y otros"


def emisor_from_subject(subj, sender):
    s = re.sub(r"^(RV|RE|FW|FWD|Fwd)\s*:\s*", "", subj or "", flags=re.I)
    for pat in [r"emitida por (.+?) para", r"Emisor:\s*(.+?)(?: Notificacion|$)", r"de (.+?) - \[", r"^(.+?) - Comprobantes",
                r"\((.+?)\s*\)", r"COOPEMADEREROS"]:
        m = re.search(pat, s, re.I)
        if m:
            return (m.group(1) if m.groups() else m.group(0)).strip()
    if sender:
        n = re.sub(r"<.*?>", "", sender).strip().strip('"')
        if n and "canavida" not in n.lower():
            return n
    return s[:60]


def load(dl):
    metas = {}
    for f in glob.glob(os.path.join(dl, "meta_*.jsonl")):
        for line in open(f, encoding="utf-8"):
            line = line.strip()
            if line:
                try:
                    m = json.loads(line)
                    metas[m["id"]] = m
                except Exception:
                    pass
    files = defaultdict(list)
    for p in glob.glob(os.path.join(dl, "*__*")):
        files[os.path.basename(p).split("__", 1)[0]].append(p)
    facturas = []
    for mid, paths in files.items():
        meta = metas.get(mid, {"id": mid})
        xmls = [parse_xml(p) for p in paths if p.lower().endswith(".xml")]
        xmls = [x for x in xmls if x]
        pdfs = [p for p in paths if p.lower().endswith(".pdf")]
        if not pdfs:
            continue
        used = set()
        for pdf in pdfs:
            pd_ = parse_pdf(pdf)
            tx = pd_.get("texto", "")
            es_factura = bool(pd_.get("clave")) or any(xx["clave"] and xx["clave"] in tx.replace(" ", "") for xx in xmls) or \
                bool(re.search(r"factura|invoice|receipt|recibo|comprobante|tiquete|ticket electr|nota de cr[eé]dito", tx[:4000], re.I)
                     and re.search(r"total", tx, re.I))
            if re.search(r"contrato|check-?in|guest messages|contrase[ñn]a|password", os.path.basename(pdf) + " " + tx[:500], re.I) \
                    and not pd_.get("clave"):
                es_factura = False
            if not es_factura:
                continue
            x = None
            for i, xx in enumerate(xmls):
                if i not in used and (pd_.get("clave") and xx["clave"] == pd_.get("clave") or len(xmls) == 1 or len(pdfs) == 1):
                    x = xx; used.add(i); break
            r = {"msgid": mid, "pdf": pdf, "asunto": meta.get("subject"), "fecha_correo": meta.get("date"),
                 "fuente": "XML" if x else "PDF"}
            ci = clave_info((x or {}).get("clave") or pd_.get("clave"))
            src = x or {}
            r["clave"] = src.get("clave") or pd_.get("clave")
            r["consecutivo"] = src.get("consecutivo") or ci.get("consecutivo")
            r["fecha"] = src.get("fecha") or pd_.get("fecha") or ci.get("fecha")
            if not r["fecha"] and meta.get("date"):
                try:
                    r["fecha"] = dt.datetime.fromisoformat(meta["date"].replace("Z", "+00:00")).date()
                except Exception:
                    r["fecha"] = None
            r["emisor"] = src.get("emisor") or emisor_from_subject(meta.get("subject"), meta.get("orig_sender"))
            if not src.get("emisor") and re.search(r"seguin|multiple|canavida|bundle", r["emisor"] or "", re.I) and pd_.get("primera_linea"):
                r["emisor"] = pd_["primera_linea"]
            r["emisor_id"] = src.get("emisor_id") or ci.get("emisor_id")
            r["receptor_id"] = src.get("receptor_id")
            r["moneda"] = src.get("moneda") or pd_.get("moneda") or "CRC"
            r["tipo_cambio"] = src.get("tipo_cambio")
            r["total"] = src.get("total") if src.get("total") is not None else pd_.get("total")
            r["medio_pago"] = src.get("medio_pago") or pd_.get("medio_pago") or "Tarjeta"
            r["detalle"] = src.get("detalle") or ""
            r["tipo"] = src.get("tipo") or ("PreFactura" if re.search(r"pre-?factura|proforma", meta.get("subject") or "", re.I) else "PDF")
            if not src and (pd_.get("prefactura") or re.search(r"pre-?factura|proforma", meta.get("subject") or "", re.I)):
                r["tipo"] = "PreFactura"
            r["categoria"] = categoria(r["emisor"], src.get("emisor_legal"), r["detalle"], pd_.get("texto", "")[:300])
            r["emitida_por_canavida"] = (r["emisor_id"] or "").lstrip("0") in CANAVIDA_IDS
            facturas.append(r)
    # deduplicar por clave (o consecutivo+emisor)
    seen, out = set(), []
    for r in sorted(facturas, key=lambda r: (r["fuente"] != "XML", r["msgid"])):
        if r["clave"]:
            k = r["clave"]
        elif r["consecutivo"]:
            k = (r["consecutivo"], r["emisor_id"])
        else:  # factura extranjera / sin clave: un cargo por correo y monto (p.ej. invoice + receipt)
            k = ("pdf", r["msgid"], r["total"])
        if k in seen:
            continue
        seen.add(k)
        out.append(r)
    # un PDF sin XML que repite monto (±1) y fecha (±5 días) de una factura con XML es la misma factura
    xmlrecs = [r for r in out if r["fuente"] == "XML"]
    def repetida(r):
        if r["fuente"] == "XML" or r["total"] is None:
            return False
        for x in xmlrecs:
            if x["total"] is not None and abs(x["total"] - r["total"]) <= 1 and r["fecha"] and x["fecha"] \
                    and abs((x["fecha"] - r["fecha"]).days) <= 5:
                return True
        return False
    out = [r for r in out if not repetida(r)]
    out.sort(key=lambda r: (r["fecha"] or dt.date(1900, 1, 1), r["emisor"] or ""))
    return out


def usd(r, tc):
    if r["total"] is None:
        return None
    sg = -1 if r.get("tipo") == "NotaCreditoElectronica" else 1
    if r["moneda"] == "USD":
        return sg * round(r["total"], 2)
    return sg * round(r["total"] / tc, 2)


def build_excel(facturas, plantilla, salida, tc):
    wb = openpyxl.load_workbook(plantilla)
    ws = wb["Gastos Diarios"]
    # facturas ya registradas a mano: col G = nº factura (cola del consecutivo)
    existentes = set()
    last = 5
    for i in range(6, ws.max_row + 1):
        if ws.cell(i, 2).value or ws.cell(i, 3).value or ws.cell(i, 6).value:
            last = i
            g = ws.cell(i, 7).value
            try:
                g = str(int(float(str(g).split()[-1])))
            except Exception:
                g = None
            f = ws.cell(i, 6).value
            try:
                m = re.fullmatch(r"=?\s*([\d.]+)\s*(?:/\s*([\d.]+))?\s*", str(f))
                f = float(m.group(1)) / (float(m.group(2)) if m.group(2) else 1)
            except Exception:
                f = None
            fd = ws.cell(i, 1).value
            fd = fd.date() if isinstance(fd, dt.datetime) else None
            existentes.add((fd, f, g))
    style_row = 6
    row = last + 1
    agregadas = 0
    for r in facturas:
        r["en_gastos"] = False
        if r["emitida_por_canavida"] or r["tipo"] == "PreFactura" or r["total"] is None or not r["fecha"]:
            continue
        num = int(r["consecutivo"][-10:]) if r["consecutivo"] and r["consecutivo"][-10:].isdigit() else None
        u = usd(r, tc)
        def igual(e):
            fd, f, g = e
            cerca = fd is not None and abs((fd - r["fecha"]).days) <= 3
            if cerca and g and r["consecutivo"] and r["consecutivo"].endswith(g) and len(g) >= 2:
                return True
            return cerca and f is not None and u is not None and abs(f - u) < 0.5
        if any(igual(e) for e in existentes):
            r["en_gastos"] = "ya registrada"
            continue
        if not (dt.date(2026, 10, 1) <= r["fecha"] <= dt.date(2026, 12, 31)):
            continue
        sg = -1 if r["tipo"] == "NotaCreditoElectronica" else 1
        desc = ("Nota de crédito: " if sg < 0 else "") + (r["detalle"] or "Factura")
        vals = [dt.datetime.combine(r["fecha"], dt.time()), r["categoria"], desc[:80], r["emisor"],
                r["medio_pago"],
                (sg * round(r["total"], 2)) if r["moneda"] == "USD" else f"={'-' if sg < 0 else ''}{r['total']:.2f}/{tc:g}",
                num if num is not None else r["consecutivo"],
                "Auto Gmail"]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row, c, v)
            src = ws.cell(style_row, c)
            if src.has_style:
                cell._style = copy.copy(src._style)
        row += 1
        agregadas += 1
        r["en_gastos"] = True

    # hoja de facturas por mes
    if "Facturas por Mes" in wb.sheetnames:
        del wb["Facturas por Mes"]
    fs = wb.create_sheet("Facturas por Mes", 1)
    fs["A1"] = "Facturas recibidas por correo (PDF) – organizadas por mes"
    fs["A1"].font = Font(bold=True, size=14)
    fs["A2"] = f"Actualizado: {dt.date.today():%d/%m/%Y} · Fuente: Gmail · Tipo de cambio usado para colones: ₡{tc:g} por $"
    hdr = ["Fecha", "Proveedor", "Categoría", "Descripción", "Método de pago", "Moneda", "Monto original", "Monto (USD)",
           "Nº consecutivo", "Tipo", "En Gastos Diarios", "Archivo PDF"]
    head_fill = PatternFill("solid", fgColor="1F4E78")
    mes_fill = PatternFill("solid", fgColor="DDEBF7")
    r_i = 4
    pormes = OrderedDict()
    for f in facturas:
        k = (f["fecha"].year, f["fecha"].month) if f["fecha"] else (0, 0)
        pormes.setdefault(k, []).append(f)
    for (y, m), items in sorted(pormes.items()):
        fs.cell(r_i, 1, f"{MESES[m]} {y}" if y else "Sin fecha").font = Font(bold=True, size=12)
        for c in range(1, len(hdr) + 1):
            fs.cell(r_i, c).fill = mes_fill
        r_i += 1
        for c, h in enumerate(hdr, 1):
            x = fs.cell(r_i, c, h)
            x.font = Font(bold=True, color="FFFFFF")
            x.fill = head_fill
        r_i += 1
        first = r_i
        for f in items:
            vals = [f["fecha"], f["emisor"], f["categoria"], f["detalle"], f["medio_pago"], f["moneda"], f["total"], usd(f, tc),
                    f["consecutivo"], ("Emitida por Canavida" if f["emitida_por_canavida"] else f["tipo"]),
                    {True: "Agregada", False: "No", "ya registrada": "Ya estaba"}[f.get("en_gastos", False)],
                    os.path.basename(f["pdf"]).split("__", 1)[1]]
            for c, v in enumerate(vals, 1):
                cell = fs.cell(r_i, c, v)
                if c == 1:
                    cell.number_format = "dd/mm/yyyy"
                if c in (7, 8):
                    cell.number_format = "#,##0.00"
            r_i += 1
        fs.cell(r_i, 7, "Total USD (sin pre-facturas):").font = Font(bold=True)
        fs.cell(r_i, 8, f'=SUMIFS(H{first}:H{r_i - 1},J{first}:J{r_i - 1},"<>PreFactura",J{first}:J{r_i - 1},"<>Emitida por Canavida")').font = Font(bold=True)
        fs.cell(r_i, 8).number_format = "$#,##0.00"
        r_i += 2
    for col, w in zip("ABCDEFGHIJKL", [11, 34, 26, 40, 20, 8, 14, 13, 24, 18, 16, 40]):
        fs.column_dimensions[col].width = w
    fs.freeze_panes = "A4"
    wb.save(salida)
    return agregadas, pormes


def build_pdf(facturas, pormes, salida, tc, max_mb=20):
    tmp = salida.replace("Facturas_por_Mes", "Resumen_Facturas")  # resumen corto (se conserva; cabe en un correo)
    doc = SimpleDocTemplate(tmp, pagesize=landscape(letter), leftMargin=30, rightMargin=30, topMargin=30, bottomMargin=30)
    st = getSampleStyleSheet()
    small = st["BodyText"].clone("s", fontSize=7, leading=8)
    el = [Paragraph("Canavida – Facturas por mes", st["Title"]),
          Paragraph(f"Actualizado {dt.date.today():%d/%m/%Y}. {len(facturas)} facturas PDF encontradas en Gmail. "
                    f"Montos en colones convertidos a USD a {tc:g} CRC por USD.", st["BodyText"]), Spacer(1, 8)]
    resumen = [["Mes", "Facturas", "Total USD"]]
    for (y, m), items in sorted(pormes.items()):
        resumen.append([f"{MESES[m]} {y}" if y else "Sin fecha", len(items),
                        f"${sum(usd(f, tc) or 0 for f in items if not f['emitida_por_canavida'] and f['tipo'] != 'PreFactura'):,.2f}"])
    t = Table(resumen, colWidths=[150, 70, 100])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                           ("GRID", (0, 0), (-1, -1), 0.4, colors.grey), ("ALIGN", (1, 0), (-1, -1), "RIGHT")]))
    el += [t, Spacer(1, 12)]
    for (y, m), items in sorted(pormes.items()):
        el.append(Paragraph(f"{MESES[m]} {y}" if y else "Sin fecha", st["Heading2"]))
        data = [["Fecha", "Proveedor", "Categoría", "Método", "Moneda", "Monto", "USD", "Consecutivo"]]
        for f in items:
            data.append([f["fecha"].strftime("%d/%m/%Y") if f["fecha"] else "", Paragraph(f["emisor"] or "", small),
                         Paragraph(f["categoria"], small), f["medio_pago"], f["moneda"],
                         f"{f['total']:,.2f}" if f["total"] is not None else "?",
                         f"{usd(f, tc):,.2f}" if usd(f, tc) is not None else "?", f["consecutivo"] or ""])
        tb = Table(data, colWidths=[55, 170, 110, 95, 40, 70, 60, 130], repeatRows=1)
        tb.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DDEBF7")), ("FONTSIZE", (0, 0), (-1, -1), 7),
                                ("GRID", (0, 0), (-1, -1), 0.3, colors.grey), ("ALIGN", (5, 1), (6, -1), "RIGHT"),
                                ("VALIGN", (0, 0), (-1, -1), "TOP")]))
        el += [tb, Spacer(1, 10)]
    doc.build(el)
    w = PdfWriter()
    for p in PdfReader(tmp).pages:
        w.add_page(p)
    size = os.path.getsize(tmp)
    incluidas = 0
    for f in facturas:
        s = os.path.getsize(f["pdf"])
        if size + s > max_mb * 1024 * 1024:
            break
        try:
            for p in PdfReader(f["pdf"]).pages:
                w.add_page(p)
            size += s
            incluidas += 1
        except Exception:
            pass
    for p in w.pages:
        p.compress_content_streams(level=9)
        for img in p.images:
            try:
                if img.image.width > 1200:
                    img.replace(img.image, quality=60)
            except Exception:
                pass
    w.compress_identical_objects(remove_duplicates=True, remove_unreferenced=True)
    with open(salida, "wb") as fh:
        w.write(fh)
    return incluidas


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("dl"); ap.add_argument("plantilla"); ap.add_argument("out")
    ap.add_argument("--tc", type=float, default=459.0)
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    facturas = load(a.dl)
    hoy = dt.date.today().isoformat()
    xlsx = os.path.join(a.out, f"Canavida_Gastos_Facturas_{hoy}.xlsx")
    pdf = os.path.join(a.out, f"Canavida_Facturas_por_Mes_{hoy}.pdf")
    agregadas, pormes = build_excel(facturas, a.plantilla, xlsx, a.tc)
    incl = build_pdf(facturas, pormes, pdf, a.tc)
    res = {"facturas": len(facturas), "agregadas_gastos": agregadas, "pdf_incluidas": incl, "xlsx": xlsx, "pdf": pdf,
           "sin_monto": [os.path.basename(f["pdf"]) for f in facturas if f["total"] is None],
           "por_mes": {f"{MESES[m]} {y}": [len(v), round(sum(usd(f, a.tc) or 0 for f in v if not f['emitida_por_canavida'] and f['tipo'] != 'PreFactura'), 2)]
                       for (y, m), v in sorted(pormes.items()) if y}}
    json.dump([{k: (str(v) if isinstance(v, dt.date) else v) for k, v in f.items()} for f in facturas],
              open(os.path.join(a.out, "facturas.json"), "w"), ensure_ascii=False, indent=1)
    print(json.dumps(res, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
