import sys
import pandas as pd
import requests
from PIL import Image, ImageDraw, ImageFont, ImageOps
from io import BytesIO
import gc
import os
import gspread
import json
import textwrap
import time
import hashlib
import re
from collections import Counter
from datetime import datetime, timedelta
from oauth2client.service_account import ServiceAccountCredentials
from concurrent.futures import ThreadPoolExecutor
import urllib.parse

# --- CONFIGURACIÓN VIP ---
ANCHO, ALTO = 2500, 3750
SHEET_ID = "1NQdhnPxgVe6N6LiVxh1ouzt5NHtqjR22EEqL6w1RpWQ"
# Los PDF se publican en GitHub Pages (antes iban a Google Drive, cuyo token OAuth vencía)
URL_BASE_PAGES = "https://analyticsdatajg2025-cmd.github.io/GITHUB_LENTO-MOVIMIENTO_CONECTA/flyers/"
PDF_ANCHO = 1000  # Ancho final de cada página (px): ~110 KB/pág para que todo LENTO quepa holgado en Pages (<1 GB)

output_dir = "docs/flyers"   # el workflow junta esta carpeta de las 12 máquinas y la publica
cache_dir = "temp_cache_img"
os.makedirs(output_dir, exist_ok=True)
os.makedirs(cache_dir, exist_ok=True)

ahora_global = datetime.utcnow() - timedelta(hours=5)
MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Setiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}
mes_actual = MESES_ES[ahora_global.month]
VERSION_LINK = ahora_global.strftime("%Y%m%d%H%M")

cache_memoria = {}

# --- PRE-CARGA DE FUENTES ---
FONT_BOLD_COND = "Mark Simonson - Proxima Nova Alt Condensed Bold.otf"
FONT_EXTRABOLD_COND = "Mark Simonson - Proxima Nova Alt Condensed Extrabold.otf"
FONT_REGULAR_COND = "Mark Simonson - Proxima Nova Alt Condensed Regular.otf"
FONT_EXTRABOLD = "Mark Simonson - Proxima Nova Extrabold.otf"
FONT_SEMIBOLD = "Mark Simonson - Proxima Nova Semibold.otf"

f_tienda = ImageFont.truetype(FONT_EXTRABOLD_COND, 90)
f_fecha = ImageFont.truetype(FONT_BOLD_COND, 45)
f_slogan = ImageFont.truetype(FONT_EXTRABOLD, 105)
f_marca = ImageFont.truetype(FONT_SEMIBOLD, 55)
f_nombre = ImageFont.truetype(FONT_REGULAR_COND, 65)
f_soles = ImageFont.truetype(FONT_EXTRABOLD, 60)
f_precio = ImageFont.truetype(FONT_EXTRABOLD, 110)
f_sku = ImageFont.truetype(FONT_BOLD_COND, 55)
f_stock_tag = ImageFont.truetype(FONT_BOLD_COND, 35)
f_stock_val = ImageFont.truetype(FONT_BOLD_COND, 55)

LC_AMARILLO, LC_AMARILLO_OSCURO = (255, 203, 5), (235, 180, 0)
EFE_AZUL, EFE_AZUL_OSCURO = (0, 107, 213), (0, 60, 150)
EFE_NARANJA, BLANCO, NEGRO, GRIS_MARCA = (255, 100, 0), (255, 255, 255), (0, 0, 0), (100, 100, 100)

def conectar_servicios():
    info_creds = json.loads(os.environ['GOOGLE_SHEETS_JSON'])
    creds_sheets = ServiceAccountCredentials.from_json_keyfile_dict(info_creds, ["https://spreadsheets.google.com/feeds"])
    return gspread.authorize(creds_sheets).open_by_key(SHEET_ID)

def lectura_segura(client, nombre_hoja):
    for intento in range(5):
        try:
            return pd.DataFrame(client.worksheet(nombre_hoja).get_all_records())
        except Exception as e:
            if "429" in str(e):
                espera = 60 * (intento + 1) 
                print(f"!!! CUOTA EXCEDIDA en {nombre_hoja}. Esperando {espera}s...")
                time.sleep(espera)
            else: raise e
    raise Exception(f"Fallo lectura tras 5 intentos en {nombre_hoja}")

def normalizar_nombre_tienda(nombre):
    # 1. Pasamos a mayúsculas y quitamos espacios invisibles a los bordes
    s = str(nombre).strip().upper()
    
    # 2. Quitamos tildes (Opcional pero recomendado para evitar fallos de teclado)
    # Si quieres ser ultra seguro con las tildes:
    s = s.replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U")
    
    # 3. Estandarizamos la marca por si acaso
    if "CURACAO" in s: s = s.replace("CURACAO", "LC")
    
    # 4. Quitamos guiones y pegamos todo
    # Esto hace que "Tumbes - EFE" y "Tumbes-EFE" sean ambos "TUMBESEFE"
    s = s.replace("-", "").replace(".", "").replace(",", "")
    
    return "".join(s.split())

# Caché de imágenes descargadas desde la red de la empresa con herramientas/actualizar_imagenes.bat
# (lacuracao.pe bloquea las descargas desde GitHub). El workflow la baja del release "imagenes-cache".
CARPETA_CACHE_LAPTOP = "imagenes_cache"   # archivos md5(url).jpg
LISTA_NECESARIAS = "imagenes_necesarias.txt"
BLOQUEOS = (401, 403, 404, 410, 426, 429)
estado_descargas = {"intentadas": 0, "ok": 0, "corte": False}
fallos_imagenes = Counter()
ejemplos_fallo = {}

def descargar_y_cachear(url):
    if not url or str(url).lower() in ['nan', ''] or url in cache_memoria: return
    clave = hashlib.md5(url.encode()).hexdigest()
    fname = os.path.join(cache_dir, clave + ".png")
    if os.path.exists(fname):
        cache_memoria[url] = fname
        return
    # 1) Imagen ya descargada desde la laptop
    laptop = os.path.join(CARPETA_CACHE_LAPTOP, clave + ".jpg")
    if os.path.exists(laptop):
        try:
            with Image.open(laptop) as img:
                img = img.convert("RGBA"); img.thumbnail((550, 550)); img.save(fname, "PNG")
            cache_memoria[url] = fname
            return
        except Exception:
            pass
    # 2) Intentar descargar (si el CDN está bloqueando, no insistir)
    if estado_descargas["corte"]:
        cache_memoria[url] = None
        fallos_imagenes["omitida (servidor bloqueando)"] += 1
        return
    motivo = None
    try:
        res = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=10)
        if res.status_code != 200:
            motivo = f"HTTP {res.status_code}"
        else:
            with Image.open(BytesIO(res.content)) as img:
                img = img.convert("RGBA")
                img.thumbnail((550, 550))
                img.save(fname, "PNG")
            cache_memoria[url] = fname
            estado_descargas["intentadas"] += 1; estado_descargas["ok"] += 1
            return
    except Exception as e:
        motivo = type(e).__name__
    cache_memoria[url] = None
    fallos_imagenes[motivo] += 1
    ejemplos_fallo.setdefault(motivo, url)
    estado_descargas["intentadas"] += 1
    if estado_descargas["intentadas"] >= 30 and estado_descargas["ok"] == 0 and motivo in ("HTTP 426", "HTTP 403", "HTTP 429"):
        estado_descargas["corte"] = True

def draw_custom_rounded(draw, xy, radius, fill, corners=(True, True, True, True)):
    x0, y0, x1, y1 = xy
    draw.rectangle([x0 + radius, y0, x1 - radius, y1], fill=fill)
    draw.rectangle([x0, y0 + radius, x1, y1 - radius], fill=fill)
    # Solo dibuja el redondeado si la esquina NO está pegada al borde
    if corners[0]: draw.pieslice([x0, y0, x0 + radius * 2, y0 + radius * 2], 180, 270, fill=fill)
    else: draw.rectangle([x0, y0, x0 + radius, y0 + radius], fill=fill) # Esquina recta
    
    if corners[1]: draw.pieslice([x1 - radius * 2, y0, x1, y0 + radius * 2], 270, 360, fill=fill)
    else: draw.rectangle([x1 - radius, y0, x1, y0 + radius], fill=fill)
    
    if corners[2]: draw.pieslice([x1 - radius * 2, y1 - radius * 2, x1, y1], 0, 90, fill=fill)
    else: draw.rectangle([x1 - radius, y1 - radius, x1, y1], fill=fill)
    
    if corners[3]: draw.pieslice([x0, y1 - radius * 2, x0 + radius * 2, y1], 90, 180, fill=fill)
    else: draw.rectangle([x0, y1 - radius, x0 + radius, y1], fill=fill)

def limpiar_valor_puro(valor, es_precio=True):
    s = str(valor).strip().replace(" ", "")
    if s in ["0", "0.0", "", "nan", "-", "SIN PRECIO"]: return "-"
    if s.endswith(".0"): s = s[:-2]
    return s

def crear_flyer(productos, tienda_nombre, num_pag):
    ahora = datetime.utcnow() - timedelta(hours=5)
    fecha_hoy = ahora.strftime("%d/%m/%Y %I:%M %p")
    
    es_efe = "EFE" in tienda_nombre.upper()
    color_fondo = EFE_AZUL_OSCURO if es_efe else LC_AMARILLO_OSCURO
    color_slogan_bg = EFE_AZUL if es_efe else LC_AMARILLO
    flyer = Image.new('RGB', (ANCHO, ALTO), color=color_fondo)
    draw = ImageDraw.Draw(flyer)
    
    try:
        bg_p = "efe tienda.jpg" if es_efe else "LC-MIRAFLORES-LOGO-3D[2].jpg"
        with Image.open(bg_p) as b:
            flyer.paste(ImageOps.fit(b.convert("RGBA"), (ANCHO, 1000)), (0, 0))
        logo_p = "logo-efe-sin-fondo.png" if es_efe else "logo-lc-sin-fondo.png"
        with Image.open(logo_p).convert("RGBA") as l:
            if es_efe:
                draw.ellipse([ANCHO-540, 40, ANCHO-80, 500], fill=BLANCO)
                logo = ImageOps.contain(l, (390, 390))
                flyer.paste(logo, (ANCHO-540+(460-logo.width)//2, 40+(460-logo.height)//2), logo)
            else:
                draw_custom_rounded(draw, [ANCHO-580, 0, ANCHO-80, 380], 50, BLANCO, (False, False, True, True))
                logo = ImageOps.contain(l, (425, 300))
                flyer.paste(logo, (ANCHO-580+(500-logo.width)//2, 50), logo)
    except: pass

    txt_tienda = tienda_nombre.upper()
    tw_t = draw.textlength(txt_tienda, font=f_tienda)
    if es_efe:
        draw_custom_rounded(draw, [ANCHO - tw_t - 150, 620, ANCHO, 800], 50, EFE_NARANJA, (True, False, False, True))
        draw.text((ANCHO - tw_t - 80, 655), txt_tienda, font=f_tienda, fill=BLANCO)
    else:
        p_x = ANCHO - tw_t - 250
        draw.polygon([(p_x, 720), (p_x + 100, 520), (ANCHO, 520), (ANCHO, 720)], fill=NEGRO)
        draw.text((ANCHO - tw_t - 100, 570), txt_tienda, font=f_tienda, fill=LC_AMARILLO)

    txt_gen = f"Generado: {fecha_hoy} - PÁG {num_pag}"
    draw_custom_rounded(draw, [0, 850, 850, 960], 40, BLANCO, (False, True, True, False))
    draw.text((40, 880), txt_gen, font=f_fecha, fill=NEGRO)
    
    draw.rectangle([0, 1030, ANCHO, 1260], fill=color_slogan_bg)
    draw.text((ANCHO//2, 1145), "¡APROVECHA ESTAS INCREÍBLES OFERTAS!", font=f_slogan, fill=BLANCO if es_efe else NEGRO, anchor="mm")

    anchos, altos = [110, 1300], [1350, 2150, 2950]
    for i, prod in enumerate(productos):
        x, y = anchos[i%2], altos[i//2]
        draw.rounded_rectangle([x, y, x+1090, y+760], radius=70, fill=BLANCO)
        
        path_img = cache_memoria.get(prod.get('image_link'))
        if path_img and os.path.exists(path_img):
            with Image.open(path_img) as img:
                flyer.paste(img.convert("RGBA"), (x+50, y+200), img.convert("RGBA"))

        stock_txt = limpiar_valor_puro(prod.get('Stock Und LM', '0'), es_precio=False)
        color_st = GRIS_MARCA if stock_txt == "-" else (EFE_AZUL if es_efe else LC_AMARILLO)
        draw.rounded_rectangle([x+30, y+30, x+300, y+160], radius=20, fill=color_st)
        draw.text((x+165, y+65), "STOCK", font=f_stock_tag, fill=BLANCO if es_efe else NEGRO, anchor="mm")
        draw.text((x+165, y+115), stock_txt, font=f_stock_val, fill=BLANCO if es_efe else NEGRO, anchor="mm")

        tx, area_w = x + 600, 450
        draw.text((tx, y+80), str(prod.get('Marca', '')).upper(), font=f_marca, fill=GRIS_MARCA)
        lines = textwrap.wrap(str(prod.get('Nombre Articulo', '')), width=16)
        ty = y + 160
        for line in lines[:3]:
            draw.text((tx, ty), line, font=f_nombre, fill=NEGRO)
            ty += 75

        # --- AJUSTE MAESTRO DE PRECIO Y COLOR ---
        p_raw = str(prod.get('Precio Vigente', '0'))
        p_final = limpiar_valor_puro(p_raw, es_precio=True)
        ty_p = y + 420
        
        # Validamos si no hay precio de forma estricta
        es_vacio = p_final == "-" or "SIN PRECIO" in p_raw.upper()
        
        if es_vacio:
            p_final = "-"
            color_p_bg = GRIS_MARCA
        else:
            color_p_bg = color_slogan_bg
            
        # Dibujamos contenedor
        draw_custom_rounded(draw, [tx, ty_p, tx + area_w, ty_p + 180], 25, color_p_bg, (True, True, False, False))
        
        if es_vacio:
            # Dibujamos SOLO el guion centrado, sin "S/ "
            draw.text((tx + area_w//2, ty_p + 90), "-", font=f_precio, fill=BLANCO if es_efe else NEGRO, anchor="mm")
        else:
            # Dibujamos precio normal con S/ 
            t_sol = "S/ "
            w_total = draw.textlength(t_sol, font=f_soles) + draw.textlength(p_final, font=f_precio)
            curr_x = tx + (area_w - w_total) // 2
            draw.text((curr_x, ty_p + 105), t_sol, font=f_soles, fill=BLANCO if es_efe else NEGRO, anchor="ls")
            draw.text((curr_x + draw.textlength(t_sol, font=f_soles), ty_p + 115), p_final, font=f_precio, fill=BLANCO if es_efe else NEGRO, anchor="ls")
        
        sku_val = str(prod['SKU'])
        sku_c = NEGRO if not es_efe else EFE_NARANJA
        draw_custom_rounded(draw, [tx, ty_p+180, tx+area_w, ty_p+280], 25, sku_c, (False, False, True, True))
        draw.text((tx+area_w//2, ty_p+230), sku_val, font=f_sku, fill=BLANCO, anchor="mm")
    return flyer

def procesar_tienda_batch(data):
    nombre, grupo = data
    try:
        prods = grupo.to_dict('records')
        clean = "".join(c for c in str(nombre) if c.isalnum() or c in " _").strip().replace(" ", "_")
        fn = f"LENTO_{clean}.pdf"
        local_path = os.path.join(output_dir, fn)
        
        # --- CONFIGURACIÓN DINÁMICA ---
        num_total_prods = len(prods)
        num_pags = (num_total_prods // 6) + (1 if num_total_prods % 6 != 0 else 0)
        
        # Compresión JPEG de cada página (calidad 40, croma 4:2:0) para que el sitio de Pages no pase de 1 GB
        calidad_final = 40
        muestreo = 2
        
        print(f">> Procesando {nombre}: {num_pags} páginas totales.")

        paginas_totales = []
        
        for i in range(0, len(prods), 6):
            pag_actual = (i//6) + 1
            img_full = crear_flyer(prods[i:i+6], str(nombre), pag_actual)
            img_flyer = img_full.convert("RGB").resize((PDF_ANCHO, int(ALTO * PDF_ANCHO / ANCHO)), Image.Resampling.LANCZOS)
            img_full.close()
            
            # [!] OPTIMIZACIÓN: Reducimos la paleta de colores para que cada objeto 
            # en la lista ocupe 4 veces menos RAM antes del guardado final.
            paginas_totales.append(img_flyer)
            
            if pag_actual % 20 == 0:
                gc.collect()

        if paginas_totales:
            print(f">> Generando PDF final con {len(paginas_totales)} páginas...")
            
            # [!] EL CAMBIO CLAVE: Un solo .save() con TODA la lista.
            # Al usar 'optimize=True', PIL comprimirá todo al escribir en disco.
            paginas_totales[0].save(
                local_path, 
                save_all=True, 
                append_images=paginas_totales[1:], 
                quality=calidad_final, 
                optimize=True,
                progressive=True,
                subsampling=muestreo
            )
            
            # --- LIMPIEZA TOTAL POST-GUARDADO ---
            for p in paginas_totales: 
                p.close()
            del paginas_totales
            gc.collect()
            
            # El PDF queda en docs/flyers y el workflow lo publica en GitHub Pages.
            # ?v=fecha obliga al navegador a bajar la versión nueva y no la guardada en caché
            link = f"{URL_BASE_PAGES}{urllib.parse.quote(fn)}?v={VERSION_LINK}"
            return [nombre, link]
            
    except Exception as e: 
        print(f"!!! ERROR CRÍTICO EN {nombre}: {e}")
        if 'paginas_totales' in locals():
            for p in paginas_totales: p.close()
            del paginas_totales
        gc.collect()
            
    return None

# --- FLUJO PRINCIPAL REFORMADO ---
ss_client = conectar_servicios()

# [!] REGLA DE ORO: Escalonamiento inicial
inicio = int(sys.argv[1]) if len(sys.argv) > 1 else 0
fin = int(sys.argv[2]) if len(sys.argv) > 2 else 999
time.sleep((inicio // 25) * 45) 

# 1. CARGA DE DATOS ORIGEN
df_raw = lectura_segura(ss_client, "Origen Tdas")
# Columnas por nombre (si alguien mueve una columna, no se leen datos equivocados).
# Si un nombre no existe, se usa la posición histórica (B, D, G, H, I, L).
COLUMNAS_ORIGEN = {'Mes': 1, 'Tienda': 3, 'Marca': 6, 'SKU': 7, 'Nombre Articulo': 8, 'Stock Und LM': 11}
df_origen = pd.DataFrame({
    col: (df_raw[col] if col in df_raw.columns else df_raw.iloc[:, pos])
    for col, pos in COLUMNAS_ORIGEN.items()
})

# 2. CARGA DE AUXILIARES (Solo para Mapeo de Precios)
df_lookup = lectura_segura(ss_client, "listado_productos")
img_dict = df_lookup.set_index('sku')['base_image_path'].to_dict()
df_origen['image_link'] = df_origen['SKU'].astype(str).str.replace('-EX', '', case=False).map(img_dict).fillna('')

# Cargamos TiendasxLista SOLO para saber qué LISTA de precios le toca a cada nombre
df_txl = lectura_segura(ss_client, "TiendasxLista")
txl_map = {normalizar_nombre_tienda(r['TIENDA']): str(r['LISTA']).replace(".0","") for r in df_txl.to_dict('records') if 'TIENDA' in r}

# 3. [!] NUEVO FILTRO DE TIENDAS (BASADO 100% EN NOMBRE DE ORIGEN TDAS)
def es_tienda_objetivo(nombre):
    # Añadimos replace de guion para que el filtro "in" sea más limpio
    n_clean = str(nombre).upper().replace("-", " ").strip()
    return "EFE" in n_clean or "LC" in n_clean or "CURACAO" in n_clean

df_origen = df_origen[df_origen['Tienda'].apply(es_tienda_objetivo)]

# 4. ASIGNACIÓN DE PRECIOS Y STOCK
promos = {}
for p in ["Promo01", "Promo03", "Promo04"]:
    df_p = lectura_segura(ss_client, p)
    df_p['K'] = df_p['Lista Precios'].astype(str).str.replace(".0","") + "_" + df_p['SKU'].astype(str)
    promos.update(df_p.set_index('K')['Precio Vigente'].to_dict())

# Asignamos la lista buscando el nombre normalizado (ordenado alfabéticamente)
df_origen['LISTA'] = df_origen['Tienda'].apply(normalizar_nombre_tienda).map(txl_map).fillna("")
df_origen['Precio Vigente'] = (df_origen['LISTA'] + "_" + df_origen['SKU'].astype(str)).map(promos).fillna("SIN PRECIO")

# 5. FILTRADO FINAL DE CALIDAD (Mes + Existencia real)
def validar_existencia(row):
    st = str(row['Stock Und LM']).strip()
    pr = str(row['Precio Vigente']).strip()
    # Si no tiene stock o no tiene precio, lo sacamos del catálogo
    if st in ["0", "0.0", "", "-", "nan"] or pr in ["SIN PRECIO", "0", "0.0", "", "-", "nan"]:
        return False
    return True

mes_col = df_origen['Mes'].astype(str).str.strip().str.capitalize().replace({"Septiembre": "Setiembre"})
meses_cargados = [m for m in MESES_ES.values() if (mes_col == m).any()]
print(f">> Meses en 'Origen Tdas': {meses_cargados}. Mes actual: {mes_actual}")
mes_usado = mes_actual
if mes_actual not in meses_cargados:
    # A inicio de mes los datos nuevos aún no están: se usa el último mes cargado para no dejar sin catálogo
    orden = list(MESES_ES.values())
    idx = orden.index(mes_actual)
    anteriores = [orden[(idx - k) % 12] for k in range(1, 12)]
    mes_usado = next((m for m in anteriores if m in meses_cargados), None)
    print(f"!!! AVISO: no hay datos de {mes_actual} en 'Origen Tdas'. Se usa el último mes cargado: {mes_usado}.")
df_final = df_origen[mes_col == mes_usado].copy()
if not df_final.empty:
    df_final = df_final[df_final.apply(validar_existencia, axis=1)]

print(f">> Tiendas detectadas con marca EFE/LC: {df_origen['Tienda'].nunique()}")
print(f">> Filtro aplicado: {len(df_final)} productos con stock y precio real (mes {mes_usado}).")
if df_final.empty:
    # Antes el script seguía, vaciaba FLYER_TIENDA y caía con KeyError: 'Tienda'
    print("!!! ERROR: 0 productos tras el filtro. Revisa 'Origen Tdas' (columna Mes), TiendasxLista y Promo01/03/04.")
    print("!!! No se toca el Sheet ni FLYER_TIENDA para conservar los catálogos anteriores.")
    sys.exit(1)

# Lista de imágenes para herramientas/actualizar_imagenes.bat (la sube la máquina 0 al release)
if inicio == 0:
    urls_necesarias = sorted(set(u for u in df_final['image_link'] if u))
    with open(LISTA_NECESARIAS, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(urls_necesarias) + "\n")
    print(f">> {len(urls_necesarias)} imágenes necesarias escritas en {LISTA_NECESARIAS}")

# [!] ACTUALIZACIÓN DE HOJA MAESTRA (Solo la primera máquina)
if inicio == 0:
    try:
        print(">> Iniciando mantenimiento de Sheets y actualización maestra...")
        
        # 1. Identificar hojas basura PRIMERO
        hojas_actuales = ss_client.worksheets()
        # Solo hojas creadas por error ("Hoja 2", "Hoja2 (1)", "Detalle de Inventario (2)"...);
        # nunca las hojas que usa el proceso.
        protegidas = {"Origen Tdas", "listado_productos", "TiendasxLista", "Promo01", "Promo03",
                      "Promo04", "Detalle de Inventario", "FLYER_TIENDA"}
        patron = re.compile(r"^(Hoja\s*\d+(\s*\(\d+\))?|.+\s\(\d+\))$")
        hojas_a_borrar = [h for h in hojas_actuales if h.title not in protegidas and patron.match(h.title.strip())]
        
        for hoja in hojas_a_borrar:
            try:
                print(f"!!! Eliminando residuo: {hoja.title}")
                ss_client.del_worksheet(hoja)
                time.sleep(1.5) # Pausa vital para no saturar la cuota de Google
            except: pass # Si ya se borró por otro hilo, ignorar

        # 2. Actualizar 'Detalle de Inventario' de forma atómica
        ws_inv = ss_client.worksheet("Detalle de Inventario")
        columnas = ['Mes', 'Tienda', 'Marca', 'SKU', 'Nombre Articulo', 'Stock Und LM', 'LISTA', 'Precio Vigente', 'image_link']
        data_inv = [columnas] + df_final[columnas].fillna("-").values.tolist()
        
        # Usar update evita los duplicados que genera clear() en conexiones inestables
        ws_inv.update('A1', data_inv)
        
        # 3. Resetear FLYER_TIENDA
        print(">> Reseteando FLYER_TIENDA...")
        ws_links = ss_client.worksheet("FLYER_TIENDA")
        ws_links.update('A1', [["TIENDA", "LINK DRIVE"]])
        
        # Borrar el resto de filas para que la hoja quede limpia
        # Esto es más seguro que clear()
        ws_links.resize(rows=1) 
        ws_links.resize(rows=1000)

        print(">> Mantenimiento completado. Sheets limpio.")

    except Exception as e: 
        print(f"Error en actualización maestra: {e}")

tiendas_procesadas = list(df_final.groupby('Tienda'))
total_reales = len(tiendas_procesadas)

if len(sys.argv) > 2:
    inicio, fin = int(sys.argv[1]), int(sys.argv[2])
    # Validación Senior: Si el inicio es mayor al total, cerramos con éxito
    if inicio >= total_reales:
        print(f">>> [!] AVISO: El inicio {inicio} supera el total de tiendas ({total_reales}). Nada que hacer.")
        sys.exit(0)
    
    fin = min(fin, total_reales)
    tiendas_a_procesar = tiendas_procesadas[inicio:fin]
else:
    tiendas_a_procesar = tiendas_procesadas

# [!] SEGURO: Si no hay tiendas en este rango específico, salimos limpiamente
if not tiendas_a_procesar:
    print(">>> [!] Rango vacío para esta máquina. Finalizando...")
    sys.exit(0)

# Descarga previa (Blindada contra lista vacía)
try:
    urls = [u for u in pd.concat([g for n, g in tiendas_a_procesar])['image_link'].unique() if u]
    with ThreadPoolExecutor(max_workers=5) as exe:
        list(exe.map(descargar_y_cachear, urls))
    ok = sum(1 for u in urls if cache_memoria.get(u))
    print(f">> Imágenes: {ok}/{len(urls)} disponibles ({estado_descargas['ok']} descargadas ahora, resto desde caché).")
    for motivo, n in fallos_imagenes.most_common():
        print(f"   Aviso: {n} imágenes fallaron por {motivo}. Ejemplo: {ejemplos_fallo.get(motivo, '-')}")
    if fallos_imagenes:
        print("   Aviso: para completar las imágenes, ejecuta herramientas/actualizar_imagenes.bat en tu laptop.")
except Exception as e:
    print(f">>> Error preparando imágenes: {e}")

tienda_links = []
batch_size = 2 

for idx, data in enumerate(tiendas_a_procesar):
    try:
        res = procesar_tienda_batch(data)
        if res:
            tienda_links.append(res)
            if len(tienda_links) % batch_size == 0:
                ss_client.worksheet("FLYER_TIENDA").append_rows(tienda_links[-batch_size:])
    except Exception as e:
        print(f"Error crítico en tienda {data[0]}: {e}")
    gc.collect() 

# --- SECCIÓN FINAL CORREGIDA PARA ESCRITURA TOTAL ---
if tienda_links:
    try:
        ws_output = ss_client.worksheet("FLYER_TIENDA")
        # Calculamos cuántos faltan por escribir que no entraron en los lotes de 2
        ya_escritos = (len(tienda_links) // batch_size) * batch_size
        pendientes = tienda_links[ya_escritos:]
        
        if pendientes:
            print(f">> Escribiendo últimos {len(pendientes)} links pendientes...")
            ws_output.append_rows(pendientes)
    except Exception as e: 
        print(f"!!! Error en escritura final: {e}")

fallidas = len(tiendas_a_procesar) - len(tienda_links)
if fallidas:
    print(f"!!! {fallidas} de {len(tiendas_a_procesar)} tiendas de este bloque NO se generaron. Revisa los errores arriba.")
    sys.exit(1)
print(f">> PROCESO COMPLETADO CON ÉXITO ({len(tienda_links)} tiendas).")
