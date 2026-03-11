import pandas as pd
import requests
from PIL import Image, ImageDraw, ImageFont, ImageOps
from io import BytesIO
import os
import gspread
import json
import textwrap
import urllib.parse
import gc
import hashlib
import time
from datetime import datetime, timedelta
from oauth2client.service_account import ServiceAccountCredentials
from concurrent.futures import ThreadPoolExecutor

# --- CONFIGURACIÓN OPTIMIZADA ---
ANCHO, ALTO = 1500, 2250
SHEET_ID = "1NQdhnPxgVe6N6LiVxh1ouzt5NHtqjR22EEqL6w1RpWQ"
USUARIO_GITHUB = "analyticsdatajg2025-cmd" 
REPO_NOMBRE = "GITHUB_LENTO-MOVIMIENTO_CONECTA"
URL_BASE_PAGES = f"https://{USUARIO_GITHUB}.github.io/{REPO_NOMBRE}/"

output_dir = "docs/flyers"
cache_dir = "temp_cache_img"
os.makedirs(output_dir, exist_ok=True)
os.makedirs(cache_dir, exist_ok=True)

ahora_peru = datetime.utcnow() - timedelta(hours=5)
fecha_peru = ahora_peru.strftime("%d/%m/%Y %I:%M %p")
semana_actual = f"Sem{ahora_peru.isocalendar()[1]}"

cache_memoria = {}

# Fuentes y Colores
FONT_BOLD_COND = "Mark Simonson - Proxima Nova Alt Condensed Bold.otf"
FONT_EXTRABOLD_COND = "Mark Simonson - Proxima Nova Alt Condensed Extrabold.otf"
FONT_REGULAR_COND = "Mark Simonson - Proxima Nova Alt Condensed Regular.otf"
FONT_EXTRABOLD = "Mark Simonson - Proxima Nova Extrabold.otf"
FONT_SEMIBOLD = "Mark Simonson - Proxima Nova Semibold.otf"

LC_AMARILLO, LC_AMARILLO_OSCURO = (255, 203, 5), (235, 180, 0)
EFE_AZUL, EFE_AZUL_OSCURO = (0, 107, 213), (0, 60, 150)
EFE_NARANJA, BLANCO, NEGRO, GRIS_MARCA = (255, 100, 0), (255, 255, 255), (0, 0, 0), (100, 100, 100)

def conectar_sheets():
    info_creds = json.loads(os.environ['GOOGLE_SHEETS_JSON'])
    creds = ServiceAccountCredentials.from_json_keyfile_dict(info_creds, ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"])
    return gspread.authorize(creds).open_by_key(SHEET_ID)

def normalizar_nombre_tienda(nombre):
    s = str(nombre).upper().replace(" ", "").replace("-", "")
    if s.endswith("EFE"): s = "EFE" + s[:-3]
    if s.endswith("LC"): s = "LC" + s[:-2]
    return s

def descargar_y_cachear(url):
    if not url or str(url).lower() in ['nan', ''] or url in cache_memoria: return
    fname = os.path.join(cache_dir, hashlib.md5(url.encode()).hexdigest() + ".png")
    if os.path.exists(fname):
        cache_memoria[url] = fname
        return
    try:
        res = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=8)
        img = Image.open(BytesIO(res.content)).convert("RGBA")
        img.thumbnail((450, 450)) 
        img.save(fname, "PNG")
        cache_memoria[url] = fname
    except: cache_memoria[url] = None

def crear_flyer(productos, tienda_nombre, num_pag):
    es_efe = "EFE" in tienda_nombre.upper()
    color_fondo = EFE_AZUL_OSCURO if es_efe else LC_AMARILLO_OSCURO
    color_slogan_bg = EFE_AZUL if es_efe else LC_AMARILLO
    
    flyer = Image.new('RGB', (ANCHO, ALTO), color=color_fondo)
    draw = ImageDraw.Draw(flyer)
    
    try:
        bg_p = "efe tienda.jpg" if es_efe else "LC-MIRAFLORES-LOGO-3D[2].jpg"
        with Image.open(bg_p) as b:
            flyer.paste(ImageOps.fit(b.convert("RGBA"), (ANCHO, 600)), (0, 0))
        logo_p = "logo-efe-sin-fondo.png" if es_efe else "logo-lc-sin-fondo.png"
        with Image.open(logo_p) as l:
            logo = ImageOps.contain(l.convert("RGBA"), (300, 250))
            if es_efe:
                draw.ellipse([ANCHO-350, 30, ANCHO-50, 330], fill=BLANCO)
                flyer.paste(logo, (ANCHO-350+(300-logo.width)//2, 30+(300-logo.height)//2), logo)
            else:
                draw.rounded_rectangle([ANCHO-380, 0, ANCHO-50, 250], radius=30, fill=BLANCO)
                flyer.paste(logo, (ANCHO-380+(330-logo.width)//2, 10), logo)
    except: pass

    f_t = ImageFont.truetype(FONT_EXTRABOLD_COND, 65)
    txt_t = f"{tienda_nombre.upper()} - PÁG {num_pag}"
    draw.text((60, 400), txt_t, font=f_t, fill=BLANCO)

    draw.rectangle([0, 620, ANCHO, 760], fill=color_slogan_bg)
    draw.text((ANCHO//2, 690), "¡APROVECHA ESTAS INCREÍBLES OFERTAS!", font=ImageFont.truetype(FONT_EXTRABOLD, 65), fill=BLANCO if es_efe else NEGRO, anchor="mm")

    anchos, altos = [50, 775], [800, 1280, 1760]
    for i, prod in enumerate(productos):
        x, y = anchos[i%2], altos[i//2]
        draw.rounded_rectangle([x, y, x+675, y+450], radius=40, fill=BLANCO)
        
        # STOCK LITERAL (Ej: 6.180)
        stock_val = str(prod.get('Stock LM', '0'))
        draw.rounded_rectangle([x+15, y+15, x+220, y+85], radius=10, fill=EFE_AZUL if es_efe else LC_AMARILLO)
        draw.text((x+117, y+32), "STOCK", font=ImageFont.truetype(FONT_BOLD_COND, 22), fill=BLANCO if es_efe else NEGRO, anchor="mm")
        draw.text((x+117, y+60), stock_val, font=ImageFont.truetype(FONT_EXTRABOLD, 32), fill=BLANCO if es_efe else NEGRO, anchor="mm")

        path_img = cache_memoria.get(prod.get('image_link'))
        if path_img:
            with Image.open(path_img) as img:
                flyer.paste(img, (x+20, y+110), img)

        tx, area_w = x + 350, 310
        draw.text((tx, y+30), str(prod.get('Marca', '')).upper(), font=ImageFont.truetype(FONT_SEMIBOLD, 32), fill=GRIS_MARCA)
        lines = textwrap.wrap(str(prod.get('Nombre Articulo', '')), width=16)
        ty = y + 70
        for line in lines[:3]:
            draw.text((tx, ty), line, font=ImageFont.truetype(FONT_REGULAR_COND, 36), fill=NEGRO)
            ty += 40

        # PRECIO (Fix .00 y 0.0)
        ty_p = y + 250
        draw.rounded_rectangle([tx, ty_p, tx + area_w, ty_p + 110], radius=15, fill=color_slogan_bg)
        p_raw = str(prod.get('Precio Vigente', '0')).replace(".00", "").strip()
        
        if p_raw in ["0.0", "0", "", "nan", "SIN PRECIO"]:
            draw.text((tx + area_w//2, ty_p + 55), "SIN PRECIO", font=ImageFont.truetype(FONT_EXTRABOLD, 45), fill=BLANCO if es_efe else NEGRO, anchor="mm")
        else:
            draw.text((tx + area_w//2, ty_p + 55), f"S/ {p_raw}", font=ImageFont.truetype(FONT_EXTRABOLD, 65), fill=BLANCO if es_efe else NEGRO, anchor="mm")

        draw.rounded_rectangle([tx, ty_p+110, tx+area_w, ty_p+165], radius=12, fill=NEGRO if not es_efe else EFE_NARANJA)
        draw.text((tx+area_w//2, ty_p+137), str(prod['SKU']), font=ImageFont.truetype(FONT_BOLD_COND, 32), fill=BLANCO, anchor="mm")

    return flyer

def procesar_tienda_batch(data):
    nombre, grupo = data
    try:
        prods = grupo.to_dict('records')
        paginas = []
        for i in range(0, len(prods), 6):
            paginas.append(crear_flyer(prods[i:i+6], str(nombre), (i//6)+1).convert("RGB"))
        
        if paginas:
            clean = "".join(c for c in str(nombre) if c.isalnum() or c in " _").strip().replace(" ", "_")
            if not clean: clean = "TIENDA_SIN_NOMBRE"
            fn = f"LENTO_{clean}.pdf"
            paginas[0].save(os.path.join(output_dir, fn), save_all=True, append_images=paginas[1:], quality=35, optimize=True)
            for p in paginas: p.close()
            return [nombre, f"{URL_BASE_PAGES}view.html?file={urllib.parse.quote(fn)}"]
    except Exception as e:
        print(f"Error procesando {nombre}: {e}")
    return None

# --- FLUJO ---
print(">> [SISTEMA] Conectando y cargando datos...")
ss = conectar_sheets()

# 1. Carga de Origen
df_raw = pd.DataFrame(ss.worksheet("Origen Tdas").get_all_records())
df_origen = pd.DataFrame({
    'Semana': df_raw.iloc[:, 1], 
    'Tienda': df_raw.iloc[:, 3], 
    'Marca': df_raw.iloc[:, 6], 
    'SKU': df_raw.iloc[:, 7], 
    'Nombre Articulo': df_raw.iloc[:, 8], 
    'Stock LM': df_raw.iloc[:, 11]
})

# 2. Carga de Imágenes
df_lookup = pd.DataFrame(ss.worksheet("listado_productos").get_all_records())
img_dict = df_lookup.set_index('sku')['base_image_path'].to_dict()
df_origen['image_link'] = df_origen['SKU'].astype(str).str.replace('-EX', '', case=False).map(img_dict).fillna('')

# 3. Carga de Precios
promos = {}
for p in ["Promo01", "Promo03", "Promo04"]:
    df_p = pd.DataFrame(ss.worksheet(p).get_all_records())
    df_p.columns = df_p.columns.str.strip()
    df_p['K'] = df_p['Lista Precios'].astype(str).str.replace(".0","") + "_" + df_p['SKU'].astype(str)
    promos.update(df_p.set_index('K')['Precio Vigente'].to_dict())

# 4. Carga de Tiendas x Lista
df_txl = pd.DataFrame(ss.worksheet("TiendasxLista").get_all_records())
txl_map = {normalizar_nombre_tienda(r['TIENDA']): str(r['LISTA']).replace(".0","") for r in df_txl.to_dict('records') if 'TIENDA' in r}
df_origen['LISTA'] = df_origen['Tienda'].apply(normalizar_nombre_tienda).map(txl_map).fillna("")

# 5. Cruce Final de Precios
df_origen['Precio Vigente'] = (df_origen['LISTA'] + "_" + df_origen['SKU'].astype(str)).map(promos).fillna("SIN PRECIO")

# 6. Filtrado de Semana Actual
df_final = df_origen[df_origen['Semana'].astype(str) == semana_actual].copy()

# --- PASO CRÍTICO: GUARDAR DETALLE DE INVENTARIO ANTES DE RENDERIZAR ---
print(">> [SISTEMA] Actualizando hoja Detalle de Inventario...")
try:
    ws_det = ss.worksheet("Detalle de Inventario")
    ws_det.clear()
    # Asegurar orden de columnas solicitado por el usuario
    cols_detalle = ['Semana', 'Tienda', 'Marca', 'SKU', 'Nombre Articulo', 'Stock LM', 'LISTA', 'Precio Vigente', 'image_link']
    df_detalle = df_final[cols_detalle].astype(str)
    ws_det.update([df_detalle.columns.tolist()] + df_detalle.values.tolist(), range_name='A1')
    print(">> [OK] Detalle de Inventario actualizado.")
except Exception as e:
    print(f">> [ERROR] Falló actualización de Detalle: {e}")

# --- PASO 7: RENDERIZADO ---
print(">> [SISTEMA] Pre-descarga de imágenes...")
with ThreadPoolExecutor(max_workers=30) as exe:
    exe.map(descargar_y_cachear, df_final['image_link'].unique())

print(f">> [SISTEMA] Renderizando {len(df_final.groupby('Tienda'))} tiendas...")
tienda_links = []
with ThreadPoolExecutor(max_workers=4) as exe:
    resultados = list(exe.map(procesar_tienda_batch, df_final.groupby('Tienda')))
    tienda_links = [r for r in resultados if r]

# Actualizar hoja final de links
ss.worksheet("FLYER_TIENDA").clear()
if tienda_links:
    ss.worksheet("FLYER_TIENDA").update([["TIENDA", "LINK"]] + tienda_links, range_name='A1')
print(">> PROCESO FINALIZADO.")