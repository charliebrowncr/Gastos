#!/usr/bin/env python3
"""Procesa facturas electronicas (PDF) descargadas del correo y genera:

  1. Una copia del libro oficial de Canavida con las facturas nuevas agregadas
     a la hoja "Gastos Diarios" (mismo formato: Fecha, Categoria, Descripcion,
     Proveedor, Metodo de pago, Monto USD, N factura) y una hoja por mes
     ("Facturas <Mes> <Año>") con el detalle de todas las facturas del correo.
  2. Un PDF con el resumen por mes seguido de todas las facturas ordenadas por mes.

Uso:
  procesar_facturas.py --nuevos m1.jsonl [m2.jsonl ...] --base-nuevos DIR \
      --repo RAIZ_REPO --plantilla libro.xlsx --salida DIR

Cada linea del manifest es un JSON con: threadId, messageId, emailDate,
subject, originalSender, filename, path (relativo a --base-dir).
"""
import argparse
import datetime as dt
import json
import os
import re
import subprocess
import unicodedata
from collections import defaultdict

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

INICIO_PERIODO = dt.date(2026, 10, 1)  # el libro registra gastos desde octubre 2026
TIPO_CAMBIO = 459  # colones por dolar, el mismo que usa el libro en octubre/noviembre

MESES = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
         "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

# (patron en texto/remitente/asunto, proveedor como aparece en el libro, categoria, descripcion)
PROVEEDORES = [
    (r"bm\.cr|grupo bm|bm supermercado", "BM Supermercado", "Alimentos y bebidas", "Pedido para alimentacion empleados"),
    (r"coopeagri", "Coopeagri", "Combustible y transporte", "Gasolina"),
    (r"colonos\.com|fejungla|el colono|colono", "Colono", "Mantenimiento y reparaciones", "Materiales de Construccion"),
    (r"grupomateriales|materiales la costa", "Grupo Materiales", "Mantenimiento y reparaciones", "Materiales de Construccion"),
    (r"gruposur|sur qu[ií]mica", "Sur Quimica", "Mantenimiento y reparaciones", "Materiales de Construccion"),
    (r"ferreter[ií]a palmares", "Ferreteria Palmares", "Mantenimiento y reparaciones", "Materiales de Construccion"),
    (r"natural clean", "Natural Cleaner", "Limpieza y lavandería", "Suministros Lavanderia"),
    (r"frutas y verduras", "Frutas y Verduras Generaleñas", "Alimentos y bebidas", "Pedido para alimentacion empleados"),
    (r"supermercados unidos", "Corporacion Supermercados Unidos", "Suministros y amenities", "Suministros"),
    (r"peque[nñ]omundo|peque[nñ]o mundo", "Pequeño Mundo", "Suministros y amenities", "Suministros"),
    (r"wireless comm", "Wireless Comunications KVF", "Software y tecnología", "Internet"),
    (r"bazar y libreria rivera|libreria rivera", "Bazar y Libreria Rivera", "Suministros y amenities", "Oficina"),
    (r"ana lorena", "Ana Lorena", "Imprevistos y otros", "Taller Mecanico"),
    (r"eduardo abarca|globacorp", "Eduardo Abarca", "Imprevistos y otros", "Servicios Abogado"),
    (r"gaspro|servicentro|zavillana|petrogas|3101622925", "Gasolinera", "Combustible y transporte", "Gasolina"),
    (r"motochopper", "Motochopper", "Imprevistos y otros", "Repuestos de moto"),
    (r"uvitaprint", "Uvitaprint", "Marketing y publicidad", "Impresion"),
    (r"coopemadereros", "Coopemadereros", "Mantenimiento y reparaciones", "Madera / materiales"),
    (r"los andes|cocinaslosandes", "Industria Los Andes", "Muebleria", "Muebles / cocina"),
    (r"envases exclusivos|alegra", "Repr. Centroamericanas de Envases", "Suministros y amenities", "Envases"),
    (r"life project education", "Life Project Education", "Imprevistos y otros", "Servicios"),
    (r"vive uvita", "Vive Uvita", "Alimentos y bebidas", "Gastos de representacion"),
    (r"moncavel|3102907395", "Moncavel", "Mantenimiento y reparaciones", "Materiales de Construccion"),
    (r"repuestos uvita", "Repuestos Uvita", "Mantenimiento y reparaciones", "Repuestos"),
    (r"pricelabs", "Pricelabs", "Software y tecnología", "Membresia"),
    (r"breezeway", "Breezeway", "Software y tecnología", "Membresia"),
    (r"smart506", "Smart506", "Software y tecnología", "Factura Electronica"),
    (r"cr contadores", "CR Contadores", "Imprevistos y otros", "Servicios Contables"),
    (r"vio ?gas", "Vio Gas", "Suministros y amenities", "Gas"),
    (r"ekono", "Ekono", "Suministros y amenities", "Articulos para habitaciones"),
    (r"\bice\b|instituto costarricense de electricidad", "ICE", "Servicios públicos e internet", "Energia / telefonia"),
]


def sin_acentos(s):
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


def a_numero(txt):
    """Convierte '1,234.56', '1.234,56', '2 843,35200' o '338,100.00' a float."""
    t = txt.strip().replace("\u00a0", " ").replace(" ", "")
    if "," in t and "." in t:
        if t.rfind(",") > t.rfind("."):
            t = t.replace(".", "").replace(",", ".")
        else:
            t = t.replace(",", "")
    elif "," in t:
        if t.count(",") == 1 and len(t.split(",")[1]) != 3:
            t = t.replace(",", ".")
        else:
            t = t.replace(",", "")
    elif t.count(".") > 1:
        t = t.replace(".", "")
    try:
        return float(t)
    except ValueError:
        return None


NUM = r"(\d{1,3}(?:[ .,]\d{3}(?!\d))+(?:[.,]\d+)?|\d+(?:[.,]\d+)?)"


def extraer_texto(path):
    try:
        out = subprocess.run(["pdftotext", "-layout", path, "-"], capture_output=True, timeout=60)
        return out.stdout.decode("utf-8", "replace")
    except Exception:
        return ""


def analizar(item, base_dir):
    path = os.path.join(base_dir, item["path"])
    texto = extraer_texto(path)
    plano = sin_acentos(texto).lower()
    contexto = " ".join([item.get("subject") or "", item.get("originalSender") or "",
                         item.get("filename") or "", texto[:3000]])
    contexto_l = sin_acentos(contexto).lower()

    r = {"archivo": item["path"], "threadId": item.get("threadId"), "messageId": item.get("messageId"),
         "asunto": item.get("subject"), "remitente": item.get("originalSender"),
         "emailDate": item.get("emailDate")}

    # Clave numerica de 50 digitos (Hacienda CR): 506 DD MM AA cedula(12) consecutivo(20) ...
    clave = None
    for fuente in (item.get("filename") or "", texto, item.get("subject") or ""):
        m = re.search(r"506(\d{47})", re.sub(r"[\s-]", "", fuente))
        if m:
            clave = "506" + m.group(1)
            break
    r["clave"] = clave
    consecutivo = clave[21:41] if clave else None
    if not consecutivo:
        m = re.search(r"\b(\d{20})\b", texto + " " + (item.get("subject") or ""))
        consecutivo = m.group(1) if m else None
    r["consecutivo"] = consecutivo
    r["tipo_doc"] = {"01": "Factura", "02": "Nota de débito", "03": "Nota de crédito", "04": "Tiquete",
                     "08": "Factura de compra", "09": "Factura de exportación"}.get(
        consecutivo[8:10] if consecutivo else "", "Factura")

    fecha = None
    if clave:
        try:
            fecha = dt.date(2000 + int(clave[7:9]), int(clave[5:7]), int(clave[3:5]))
        except ValueError:
            fecha = None
    if not fecha:
        m = re.search(r"(\d{1,2})[/-](\d{1,2})[/-](20\d{2})", texto)
        if m:
            try:
                fecha = dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            except ValueError:
                pass
    if not fecha and item.get("emailDate"):
        fecha = dt.date.fromisoformat(item["emailDate"][:10])
    r["fecha"] = fecha

    # Total del comprobante. Entre la etiqueta y el numero solo se permiten espacios,
    # simbolos de moneda, ':' o '*'; asi se evitan "Total en letras", "Total gravado", etc.
    total = None
    sep = r"(?:[ \t¢₡$:*]|crc|usd)*"
    plano = re.sub(r"sub\s*-?\s*total", "subtotal", plano)
    especificas = [r"total\s+comprobante", r"total\s+(?:de\s+la\s+)?factura", r"total\s+a\s+pagar",
                   r"total\s+en\s+(?:crc|usd)", r"total\s+documento", r"monto\s+total", r"total\s+general",
                   r"total\s+servicio", r"total\s+neto"]
    for pat in especificas + [r"(?<![a-z])total"]:
        hallados = [a_numero(m.group(1)) for m in re.finditer(pat + sep + NUM, plano)]
        hallados = [v for v in hallados if v]
        if hallados:
            total = hallados[-1]
            break
    if total is None:  # ultima opcion: "total" generico (no subtotal) mas cercano al final
        hallados = [a_numero(m.group(1)) for m in re.finditer(r"(?<!sub)(?<!sub )(?<!sub - )\btotal" + sep + NUM, plano)]
        hallados = [v for v in hallados if v]
        total = hallados[-1] if hallados else None
    r["total"] = total

    # Moneda: primero la declaracion explicita "Moneda: ...", luego simbolos
    m = re.search(r"moneda[^\n]{0,80}?\b(usd|crc|dolar\w*|colon\w*|us\$)", plano)
    if m:
        moneda = "USD" if m.group(1).startswith(("usd", "dolar", "us$")) else "CRC"
    elif re.search(r"total[^\n]{0,160}\$\s?\d", plano) and not re.search(r"total[^\n]{0,160}[¢₡]", texto.lower()):
        moneda = "USD"
    elif "¢" in texto or "₡" in texto or re.search(r"\bcrc\b|colones", plano):
        moneda = "CRC"
    else:
        moneda = "CRC"
    r["moneda"] = moneda

    # Medio de pago
    metodo = "Tarjeta"
    m = re.search(r"medio\s+de\s+pago[^\n:]*:?\s*([a-z /]+)", plano)
    if m:
        v = m.group(1)
        if "efectivo" in v:
            metodo = "Efectivo"
        elif "transfer" in v or "sinpe" in v or "deposito" in v:
            metodo = "Transferencia / SINPE"
        elif "tarjeta" in v:
            metodo = "Tarjeta"
    r["metodo"] = metodo

    # Proveedor y categoria
    prov, cat, desc = None, "Imprevistos y otros", "Factura electrónica (revisar)"
    for pat, p, c, d in PROVEEDORES:
        if re.search(pat, contexto_l):
            prov, cat, desc = p, c, d
            break
    if not prov:
        m = re.search(r"emisor:\s*([^\n]+?)(?:\s+notificacion|$)", sin_acentos(item.get("subject") or ""), re.I)
        if m:
            prov = m.group(1).strip(" .")
        else:
            prim = [l.strip() for l in texto.splitlines() if l.strip()]
            prov = prim[0][:60] if prim else (item.get("originalSender") or "Desconocido")
    r["proveedor"], r["categoria"], r["descripcion"] = prov, cat, desc
    r["revisar"] = total is None or clave is None
    return r


def cargar_manifest(paths):
    items = []
    for p in paths:
        if not os.path.exists(p):
            continue
        with open(p, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    d = json.loads(line)
                except json.JSONDecodeError:
                    continue
                if d.get("filename") and d.get("path") and not d.get("error"):
                    items.append(d)
    return items


def norm(s):
    return re.sub(r"[^a-z0-9]", "", sin_acentos(str(s or "")).lower())


def consecutivos_existentes(ws):
    """Pares (numero de factura, proveedor) ya anotados en Gastos Diarios."""
    vistos = set()
    for row in ws.iter_rows(min_row=6, values_only=True):
        g = row[6] if len(row) > 6 else None
        prov = norm(row[3] if len(row) > 3 else "")
        nums = [int(g)] if isinstance(g, (int, float)) else [int(n) for n in re.findall(r"\d+", g or "")] if isinstance(g, str) else []
        for n in nums:
            vistos.add((n, prov))
    return vistos


def ultima_fila(ws):
    ultima = 5
    for i, row in enumerate(ws.iter_rows(min_row=6, max_col=6, values_only=True), start=6):
        if any(v not in (None, "") for v in row[1:6]):
            ultima = i
    return ultima


def construir_excel(facturas, plantilla, salida_xlsx):
    wb = openpyxl.load_workbook(plantilla)
    ws = wb["Gastos Diarios"]
    existentes = consecutivos_existentes(ws)
    fila = ultima_fila(ws) + 1
    agregadas, ya = [], []
    for f in facturas:
        num = int(f["consecutivo"][-10:]) if f.get("consecutivo") else None
        if num is not None and (num, norm(f["proveedor"])) in existentes:
            f["estado_libro"] = "Ya registrada"
            ya.append(f)
            continue
        if f["total"] is None:
            f["estado_libro"] = "No agregada: monto no leído"
            continue
        if f["fecha"] < INICIO_PERIODO:
            f["estado_libro"] = "Fuera del periodo del libro"
            continue
        signo = -1 if f["tipo_doc"] == "Nota de crédito" else 1
        monto = f["total"] * signo
        ws.cell(fila, 1, dt.datetime.combine(f["fecha"], dt.time())).number_format = "dd/mm/yyyy"
        ws.cell(fila, 2, f["categoria"])
        ws.cell(fila, 3, f["descripcion"] + (" (nota de crédito)" if signo < 0 else ""))
        ws.cell(fila, 4, f["proveedor"])
        ws.cell(fila, 5, f["metodo"])
        if f["moneda"] == "CRC":
            ws.cell(fila, 6, f"={round(monto, 2)}/{TIPO_CAMBIO}")
        else:
            ws.cell(fila, 6, round(monto, 2))
        ws.cell(fila, 7, num)
        ws.cell(fila, 8, "Auto-correo")
        f["estado_libro"] = f"Agregada (fila {fila})"
        agregadas.append(f)
        if num is not None:
            existentes.add((num, norm(f["proveedor"])))
        fila += 1

    # Hojas por mes
    por_mes = defaultdict(list)
    for f in facturas:
        por_mes[(f["fecha"].year, f["fecha"].month)].append(f)
    azul = PatternFill("solid", fgColor="1F4E78")
    for (y, mth) in sorted(por_mes):
        nombre = f"Facturas {MESES[mth]} {y}"
        if nombre in wb.sheetnames:
            del wb[nombre]
        hs = wb.create_sheet(nombre)
        hs["A1"] = f"Facturas recibidas por correo – {MESES[mth]} {y}"
        hs["A1"].font = Font(bold=True, size=14)
        hs["A2"] = f"Montos en colones convertidos a USD con tipo de cambio {TIPO_CAMBIO}. Generado {dt.date.today():%d/%m/%Y}."
        enc = ["Fecha", "Categoría", "Descripción", "Proveedor", "Método de pago", "Monto (USD)",
               "Nº factura / Notas", "Moneda", "Monto original", "Tipo", "Clave", "Estado en libro", "Archivo PDF"]
        for c, h in enumerate(enc, 1):
            cell = hs.cell(4, c, h)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = azul
            cell.alignment = Alignment(wrap_text=True)
        r0 = 5
        for i, f in enumerate(sorted(por_mes[(y, mth)], key=lambda x: (x["fecha"], x["proveedor"])), start=r0):
            signo = -1 if f["tipo_doc"] == "Nota de crédito" else 1
            usd = None
            if f["total"] is not None:
                usd = f["total"] * signo / (TIPO_CAMBIO if f["moneda"] == "CRC" else 1)
            vals = [f["fecha"], f["categoria"], f["descripcion"], f["proveedor"], f["metodo"],
                    round(usd, 2) if usd is not None else None,
                    int(f["consecutivo"][-10:]) if f.get("consecutivo") else None,
                    f["moneda"], f["total"], f["tipo_doc"], f.get("clave"), f.get("estado_libro", ""),
                    os.path.basename(f["archivo"])]
            for c, v in enumerate(vals, 1):
                cell = hs.cell(i, c, v)
                if c == 1:
                    cell.number_format = "dd/mm/yyyy"
                if c in (6, 9):
                    cell.number_format = "#,##0.00"
        fin = r0 + len(por_mes[(y, mth)])
        hs.cell(fin + 1, 5, "TOTAL USD").font = Font(bold=True)
        hs.cell(fin + 1, 6, f"=SUM(F{r0}:F{fin - 1})").font = Font(bold=True)
        hs.cell(fin + 1, 6).number_format = "#,##0.00"
        for c, w in enumerate([12, 28, 34, 32, 22, 14, 16, 8, 16, 14, 54, 22, 40], 1):
            hs.column_dimensions[get_column_letter(c)].width = w
        hs.freeze_panes = "A5"
    wb.save(salida_xlsx)
    return agregadas, ya, por_mes


def construir_pdf(facturas, por_mes, agregadas, base_dir, salida_pdf, tmpdir):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import landscape, letter
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak

    st = getSampleStyleSheet()
    resumen = os.path.join(tmpdir, "_resumen.pdf")
    doc = SimpleDocTemplate(resumen, pagesize=landscape(letter), leftMargin=30, rightMargin=30,
                            topMargin=30, bottomMargin=30)
    el = [Paragraph("Canavida – Facturas recibidas por correo", st["Title"]),
          Paragraph(f"Actualización del {dt.date.today():%d/%m/%Y}. Facturas leídas: {len(facturas)}. "
                    f"Agregadas al libro hoy: {len(agregadas)}. Tipo de cambio: ₡{TIPO_CAMBIO} por $.",
                    st["Normal"]), Spacer(1, 10)]
    for (y, mth) in sorted(por_mes):
        lst = sorted(por_mes[(y, mth)], key=lambda x: (x["fecha"], x["proveedor"]))
        el.append(Paragraph(f"{MESES[mth]} {y} – {len(lst)} facturas", st["Heading2"]))
        filas = [["Fecha", "Proveedor", "Categoría", "Nº", "Moneda", "Monto", "USD", "Estado"]]
        tot = 0.0
        for f in lst:
            signo = -1 if f["tipo_doc"] == "Nota de crédito" else 1
            usd = (f["total"] * signo / (TIPO_CAMBIO if f["moneda"] == "CRC" else 1)) if f["total"] else None
            tot += usd or 0
            filas.append([f"{f['fecha']:%d/%m/%Y}", Paragraph(f["proveedor"][:45], st["BodyText"]),
                          f["categoria"][:28], (f.get("consecutivo") or "")[-10:].lstrip("0"), f["moneda"],
                          f"{f['total'] * signo:,.2f}" if f["total"] else "¿?",
                          f"{usd:,.2f}" if usd is not None else "¿?", f.get("estado_libro", "")[:24]])
        filas.append(["", "", "", "", "", "TOTAL", f"{tot:,.2f}", ""])
        t = Table(filas, colWidths=[60, 180, 140, 70, 45, 80, 70, 110], repeatRows=1)
        t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
                               ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                               ("FONTSIZE", (0, 0), (-1, -1), 8),
                               ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                               ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
                               ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
        el += [t, Spacer(1, 14)]
    doc.build(el)

    archivos = [resumen]
    for (y, mth) in sorted(por_mes):
        for f in sorted(por_mes[(y, mth)], key=lambda x: (x["fecha"], x["proveedor"])):
            archivos.append(os.path.join(base_dir, f["archivo"]))
    subprocess.run(["pdfunite"] + archivos + [salida_pdf], check=True)


def nombre_limpio(f):
    prov = re.sub(r"[^A-Za-z0-9]+", "_", sin_acentos(f["proveedor"]))[:40].strip("_")
    num = (f.get("consecutivo") or "")[-10:].lstrip("0") or "sn"
    return f"{f['fecha']:%Y-%m-%d}_{prov}_{num}.pdf"


def organizar(nuevos, base_nuevos, repo):
    """Copia los PDFs nuevos a facturas/AAAA-MM/ y los anota en facturas/registro.jsonl."""
    import shutil
    reg_path = os.path.join(repo, "facturas", "registro.jsonl")
    registro = cargar_manifest([reg_path])
    vistos = {(r.get("messageId"), r.get("filename")) for r in registro}
    claves = set()
    for r in registro:
        claves.add(r.get("clave") or r["path"])
    agregados = []
    for it in nuevos:
        if (it.get("messageId"), it.get("filename")) in vistos:
            continue
        f = analizar(it, base_nuevos)
        llave = f.get("clave") or it["path"]
        if llave in claves:
            continue
        claves.add(llave)
        carpeta = f"facturas/{f['fecha']:%Y-%m}"
        os.makedirs(os.path.join(repo, carpeta), exist_ok=True)
        destino = f"{carpeta}/{nombre_limpio(f)}"
        n = 2
        while os.path.exists(os.path.join(repo, destino)):
            destino = f"{carpeta}/{nombre_limpio(f)[:-4]}_{n}.pdf"
            n += 1
        shutil.copyfile(os.path.join(base_nuevos, it["path"]), os.path.join(repo, destino))
        nuevo = dict(it, path=destino, clave=f.get("clave"), agregado=dt.date.today().isoformat())
        registro.append(nuevo)
        agregados.append(nuevo)
        with open(reg_path, "a", encoding="utf-8") as fh:
            fh.write(json.dumps(nuevo, ensure_ascii=False) + "\n")
    return registro, agregados


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--nuevos", nargs="*", default=[], help="manifests de PDFs recien descargados")
    ap.add_argument("--base-nuevos", default=".", help="carpeta base de los paths de --nuevos")
    ap.add_argument("--repo", required=True, help="raiz del repositorio Gastos")
    ap.add_argument("--plantilla", required=True)
    ap.add_argument("--salida", required=True)
    a = ap.parse_args()

    registro, recien = organizar(cargar_manifest(a.nuevos), a.base_nuevos, a.repo)
    nuevos_ids = {(r.get("messageId"), r.get("filename")) for r in recien}
    facturas = []
    for it in registro:
        f = analizar(it, a.repo)
        f["nueva"] = (it.get("messageId"), it.get("filename")) in nuevos_ids
        facturas.append(f)

    os.makedirs(a.salida, exist_ok=True)
    hoy = dt.date.today()
    xlsx = os.path.join(a.salida, f"Canavida_Gastos_Facturas_{hoy:%Y-%m-%d}.xlsx")
    pdf = os.path.join(a.salida, f"Canavida_Facturas_por_mes_{hoy:%Y-%m-%d}.pdf")
    agregadas, ya, por_mes = construir_excel(facturas, a.plantilla, xlsx)
    construir_pdf(facturas, por_mes, agregadas, a.repo, pdf, a.salida)

    def usd(x):
        if x["total"] is None:
            return 0
        return x["total"] * (-1 if x["tipo_doc"] == "Nota de crédito" else 1) / (TIPO_CAMBIO if x["moneda"] == "CRC" else 1)

    resumen = {
        "xlsx": xlsx, "pdf": pdf, "total_facturas": len(facturas), "nuevas_hoy": len(recien),
        "agregadas_libro": len(agregadas), "ya_registradas": len(ya),
        "revisar": [{"archivo": f["archivo"], "proveedor": f["proveedor"], "total": f["total"]}
                    for f in facturas if f["revisar"]],
        "por_mes": {f"{MESES[m]} {y}": {"facturas": len(v), "usd": round(sum(usd(x) for x in v), 2)}
                    for (y, m), v in sorted(por_mes.items())},
        "detalle": [{k: (str(v) if isinstance(v, dt.date) else v) for k, v in f.items()
                     if k in ("fecha", "proveedor", "categoria", "moneda", "total", "tipo_doc", "nueva",
                              "consecutivo", "estado_libro", "archivo")} for f in facturas],
    }
    print(json.dumps(resumen, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
