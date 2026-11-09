# GITHUB_LENTO-MOVIMIENTO_CONECTA
Automatización de catálogos (PDF por tienda) para los productos que tienen lento movimiento.

## Cómo funciona
- **Workflow** `.github/workflows/flyer_automation.yml`: corre todos los días a las 3 PM (Perú),
  a mano (*Actions → Run workflow*) o al cambiar el script. Reparte las tiendas en 12 máquinas.
- **Datos** (Google Sheet `1NQdhnPx…`): `Origen Tdas` (inventario por mes), `TiendasxLista`
  (lista de precios de cada tienda), `Promo01/03/04` (precios), `listado_productos` (imagen de cada SKU).
- **Salida**: un PDF por tienda en Google Drive; los links quedan en la hoja `FLYER_TIENDA` y el
  detalle filtrado en `Detalle de Inventario`.
- **Mes**: usa las filas de `Origen Tdas` del mes actual. Si aún no se cargó el mes nuevo
  (inicio de mes), usa el último mes cargado y lo avisa en el log.
- **Imágenes**: lacuracao.pe bloquea las descargas desde GitHub. Las imágenes se descargan desde
  la laptop con la herramienta del repo OBSOLETOS
  (`GITHUB_OBSOLETOS_CONECTA/herramientas/actualizar_imagenes.bat`), que actualiza los dos procesos.

## Secrets necesarios
- `GOOGLE_SHEETS_JSON`: cuenta de servicio con permiso de editor en el Sheet.
- `GOOGLE_TOKEN_JSON`: token OAuth de usuario para subir los PDF a la carpeta de Drive.
