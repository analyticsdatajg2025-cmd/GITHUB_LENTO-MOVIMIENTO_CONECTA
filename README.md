# GITHUB_LENTO-MOVIMIENTO_CONECTA
Automatización de catálogos (PDF por tienda) para los productos que tienen lento movimiento.

## Cómo funciona
- **Workflow** `.github/workflows/generar_flyers_lento.yml`: corre todos los días a las 3 PM (Perú),
  a mano (*Actions → Run workflow*) o al cambiar el script. Reparte las tiendas en 12 máquinas.
- **Datos** (Google Sheet `1NQdhnPx…`): `Origen Tdas` (inventario por mes), `TiendasxLista`
  (lista de precios de cada tienda), `Promo01/03/04` (precios), `listado_productos` (imagen de cada SKU).
- **Salida**: un PDF por tienda en la carpeta de Google Drive `1UzcwDFT…`. Cada tienda conserva siempre
  el mismo archivo (se actualiza por nombre), así el link no cambia: lo usa la página `lacuracao.pe/lm`.
  Los links quedan en la hoja `FLYER_TIENDA` y el detalle filtrado en `Detalle de Inventario`.
- **Mes**: usa las filas de `Origen Tdas` del mes actual. Si aún no se cargó el mes nuevo
  (inicio de mes), usa el último mes cargado y lo avisa en el log.
- **Imágenes**: lacuracao.pe bloquea las descargas desde GitHub. Las imágenes se descargan desde
  la laptop con la herramienta del repo OBSOLETOS
  (`GITHUB_OBSOLETOS_CONECTA/herramientas/actualizar_imagenes.bat`), que actualiza los dos procesos.

## Configuración necesaria
- Secret `GOOGLE_SHEETS_JSON`: cuenta de servicio con permiso de editor en el Sheet **y en la carpeta de Drive**
  (compártela con el `client_email` de la cuenta de servicio, como Editor). Con eso actualiza los PDF existentes
  sin tokens que venzan.
- Secret `GOOGLE_TOKEN_JSON` (opcional): token OAuth personal. Solo hace falta para CREAR el PDF de una tienda
  nueva; si no está o venció, se usa la cuenta de servicio (y para una tienda nueva basta subir una vez a la
  carpeta un PDF cualquiera con el nombre `LENTO_<tienda>.pdf`).
