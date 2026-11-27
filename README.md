# Gastos Canavida – facturas desde Gmail

Rutina diaria: busca en Gmail (admicanavida@gmail.com) los correos con facturas PDF, descarga PDF + XML de Hacienda
y genera, organizado por mes en el formato de `Copy of Canavida Gastos Ingresos Oficial (2).xlsx`:

- `reportes/Canavida_Gastos_Facturas_<fecha>.xlsx`: plantilla con los gastos nuevos agregados a *Gastos Diarios* (columna H = "Auto Gmail") y la hoja *Facturas por Mes*.
- `reportes/Canavida_Facturas_por_Mes_<fecha>.pdf`: resumen por mes + todas las facturas (≈8.5 MB).
- `reportes/Canavida_Resumen_Facturas_<fecha>.pdf`: solo el resumen (pequeño, apto para adjuntar).

Estado: `data/correos_procesados.txt` (IDs de Gmail ya procesados) y `data/facturas.json`.

```
python3 scripts/procesar_facturas.py <dir_descargas> "Copy of Canavida Gastos Ingresos Oficial (2).xlsx" <dir_salida> --tc 459
```
