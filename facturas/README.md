# Facturas por mes

Esta carpeta la llena automáticamente la rutina diaria:

1. Busca en el correo admicanavida@gmail.com los mensajes con facturas en PDF.
2. Guarda cada PDF en `facturas/AAAA-MM/` con el nombre `fecha_proveedor_numero.pdf`.
3. Anota cada factura en `registro.jsonl`, que sirve para no procesarla dos veces.
4. Genera una copia del libro oficial con las facturas nuevas agregadas a **Gastos Diarios**
   (mismas columnas: Fecha, Categoría, Descripción, Proveedor, Método de pago, Monto USD,
   Nº factura) y una hoja **Facturas <Mes> <Año>** por cada mes. Las filas agregadas
   llevan la nota `Auto-correo` en la columna H.
5. Envía a admicanavida@gmail.com un correo con el Excel y un PDF (resumen por mes más
   todas las facturas ordenadas por mes).

Notas:

- Los montos en colones se convierten con el tipo de cambio de `TIPO_CAMBIO` en
  `scripts/procesar_facturas.py` (₡459 por $, el mismo que usa el libro), escrito como
  fórmula `=monto/459` igual que en las filas manuales.
- La categoría se asigna según el proveedor (tabla `PROVEEDORES` del script). Los
  proveedores nuevos quedan en "Imprevistos y otros"; conviene revisarlos.
- Una factura no se agrega a Gastos Diarios si su número y su proveedor ya están en el
  libro, o si es anterior a octubre de 2026.
- Las notas de crédito se restan (monto negativo).
