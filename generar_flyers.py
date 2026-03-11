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

# --- CONFIGURACIÓN ESCALADA (1000x1500) ---
# Reducción drástica de peso para evitar Error 500 en GitHub
ANCHO, ALTO = 1000, 1500
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

# Fuentes (Tamaños ajustados a la nueva escala)
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
        img.thumbnail((300, 300)) # Imagen más pequeña para ahorro de memoria
        img.save(fname, "PNG")
        cache_memoria[url] = fname
    except: cache_memoria[url] = None

def crear_flyer(productos, tienda_nombre, num_pag):
    es_efe = "EFE" in tienda_nombre.upper()
    color_fondo = EFE_AZUL_OSCURO if es_efe else LC_AMARILLO_OSCURO
    color_slogan_bg = EFE_AZUL if es_efe else LC_AMARILLO
    
    flyer = Image.new('RGB', (ANCHO, ALTO), color=color_fondo)
    draw = ImageDraw.Draw(flyer)
    
    # Header (400px altura proporcional)
    try:
        bg_p = "efe tienda.jpg" if es_efe else "LC-MIRAFLORES-LOGO-3D[2].jpg"
        with Image.open(bg_p) as b:
            flyer.paste(ImageOps.fit(b.convert("RGBA"), (ANCHO, 400)), (0, 0))
        logo_p = "logo-efe-sin-fondo.png" if es_efe else "logo-lc-sin-fondo.png"
        with Image.open(logo_p) as l:
            logo = ImageOps.contain(l.convert("RGBA"), (200, 170))
            if es_efe:
                draw.ellipse([ANCHO-250, 20, ANCHO-30, 240], fill=BLANCO)
                flyer.paste(logo, (ANCHO-250+(220-logo.width)//2, 20+(220-logo.height)//2), logo)
            else:
                draw.rounded_rectangle([ANCHO-280, 0, ANCHO-30, 180], radius=25, fill=BLANCO)
                flyer.paste(logo, (ANCHO-280+(250-logo.width)//2, 5), logo)
    except: pass

    # Título Tienda
    f_t = ImageFont.truetype(FONT_EXTRABOLD_COND, 45)
    txt_t = f"{tienda_nombre.upper()} - PÁG {num_pag}"
    draw.text((40, 280), txt_t, font=f_t, fill=BLANCO)

    # Franja Slogan
    draw.rectangle([0, 410, ANCHO, 510], fill=color_slogan_bg)
    draw.text((ANCHO//2, 460), "¡APROVECHA ESTAS INCREÍBLES OFERTAS!", font=ImageFont.truetype(FONT_EXTRABOLD, 45), fill=BLANCO if es_efe else NEGRO, anchor="mm")

    # Grilla de Productos (3 filas x 2 columnas = 6 prods)
    # x: [Izquierda, Derecha], y: [Fila 1, Fila 2, Fila 3]
    anchos, altos = [30, 515], [540, 860, 1180]
    for i, prod in enumerate(productos):
        x, y = anchos[i%2], altos[i//2]
        draw.rounded_rectangle([x, y, x+455, y+300], radius=30, fill=BLANCO)
        
        # STOCK LITERAL
        stock_val = str(prod.get('Stock LM', '0'))
        draw.rounded_rectangle([x+10, y+10, x+150, y+60], radius=8, fill=EFE_AZUL if es_efe else LC_AMARILLO)
        draw.text((x+80, y+22), "STOCK", font=ImageFont.truetype(FONT_BOLD_COND, 16), fill=BLANCO if es_efe else NEGRO, anchor="mm")
        draw.text((x+80, y+42), stock_val, font=ImageFont.truetype(FONT_EXTRABOLD, 22), fill=BLANCO if es_efe else NEGRO, anchor="mm")

        path_img = cache_memoria.get(prod.get('image_link'))
        if path_img:
            with Image.open(path_img) as img:
                flyer.paste(img, (x+10, y+70), img)

        tx, area_w = x + 230, 210
        draw.text((tx, y+20), str(prod.get('Marca', '')).upper(), font=ImageFont.truetype(FONT_SEMIBOLD, 22), fill=GRIS_MARCA)
        lines = textwrap.wrap(str(prod.get('Nombre Articulo', '')), width=16)
        ty = y + 45
        for line in lines[:3]:
            draw.text((tx, ty), line, font=ImageFont.truetype(FONT_REGULAR_COND, 26), fill=NEGRO)
            ty += 30

        # PRECIO / SKU
        ty_p = y + 160
        draw.rounded_rectangle([tx, ty_p, tx + area_w, ty_p + 80], radius=10, fill=color_slogan_bg)
        p_raw = str(prod.get('Precio Vigente', '0')).replace(".00", "").strip()
        
        if p_raw in ["0.0", "0", "", "nan", "SIN PRECIO"]:
            draw.text((tx + area_w//2, ty_p + 40), "SIN PRECIO", font=ImageFont.truetype(FONT_EXTRABOLD, 32), fill=BLANCO if es_efe else NEGRO, anchor="mm")
        else:
            draw.text((tx + area_w//2, ty_p + 40), f"S/ {p_raw}", font=ImageFont.truetype(FONT_EXTRABOLD, 45), fill=BLANCO if es_efe else NEGRO, anchor="mm")

        draw.rounded_rectangle([tx, ty_p+80, tx+area_w, ty_p+120], radius=8, fill=NEGRO if not es_efe else EFE_NARANJA)
        draw.text((tx+area_w//2, ty_p+100), str(prod['SKU']), font=ImageFont.truetype(FONT_BOLD_COND, 22), fill=BLANCO, anchor="mm")

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
            if not clean: clean = "TIENDA"
            fn = f"LENTO_{clean}.pdf"
            # Calidad 20 para máxima ligereza y evitar Error 500
            paginas[0].save(os.path.join(output_dir, fn), save_all=True, append_images=paginas[1:], quality=20, optimize=True)
            for p in paginas: p.close()
            return [nombre, f"{URL_BASE_PAGES}view.html?file={urllib.parse.quote(fn)}"]
    except Exception as e:
        print(f"Error en {nombre}: {e}")
    return None

# --- FLUJO ---
print(">> [PASO 1] Cargando datos de Google Sheets...")
ss = conectar_sheets()

df_raw = pd.DataFrame(ss.worksheet("Origen Tdas").get_all_records())
df_origen = pd.DataFrame({
    'Semana': df_raw.iloc[:, 1], 
    'Tienda': df_raw.iloc[:, 3], 
    'Marca': df_raw.iloc[:, 6], 
    'SKU': df_raw.iloc[:, 7], 
    'Nombre Articulo': df_raw.iloc[:, 8], 
    'Stock LM': df_raw.iloc[:, 11]
})

df_lookup = pd.DataFrame(ss.worksheet("listado_productos").get_all_records())
img_dict = df_lookup.set_index('sku')['base_image_path'].to_dict()
df_origen['image_link'] = df_origen['SKU'].astype(str).str.replace('-EX', '', case=False).map(img_dict).fillna('')

promos = {}
for p in ["Promo01", "Promo03", "Promo04"]:
    df_p = pd.DataFrame(ss.worksheet(p).get_all_records())
    df_p.columns = df_p.columns.str.strip()
    df_p['K'] = df_p['Lista Precios'].astype(str).str.replace(".0","") + "_" + df_p['SKU'].astype(str)
    promos.update(df_p.set_index('K')['Precio Vigente'].to_dict())

df_txl = pd.DataFrame(ss.worksheet("TiendasxLista").get_all_records())
txl_map = {normalizar_nombre_tienda(r['TIENDA']): str(r['LISTA']).replace(".0","") for r in df_txl.to_dict('records') if 'TIENDA' in r}
df_origen['LISTA'] = df_origen['Tienda'].apply(normalizar_nombre_tienda).map(txl_map).fillna("")
df_origen['Precio Vigente'] = (df_origen['LISTA'] + "_" + df_origen['SKU'].astype(str)).map(promos).fillna("SIN PRECIO")

df_final = df_origen[df_origen['Semana'].astype(str) == semana_actual].copy()

print(">> [PASO 2] Actualizando Detalle de Inventario...")
try:
    ws_det = ss.worksheet("Detalle de Inventario")
    ws_det.clear()
    cols = ['Semana', 'Tienda', 'Marca', 'SKU', 'Nombre Articulo', 'Stock LM', 'LISTA', 'Precio Vigente', 'image_link']
    df_det = df_final[cols].astype(str)
    ws_det.update([df_det.columns.tolist()] + df_det.values.tolist(), range_name='A1')
except: pass

print(">> [PASO 3] Pre-descarga de imágenes...")
with ThreadPoolExecutor(max_workers=30) as exe:
    exe.map(descargar_y_cachear, df_final['image_link'].unique())

print(">> [PASO 4] Generando Flyers Multipágina...")
tienda_links = []
with ThreadPoolExecutor(max_workers=4) as exe:
    resultados = list(exe.map(procesar_tienda_batch, df_final.groupby('Tienda')))
    tienda_links = [r for r in resultados if r]

ss.worksheet("FLYER_TIENDA").clear()
if tienda_links:
    ss.worksheet("FLYER_TIENDA").update([["TIENDA", "LINK"]] + tienda_links, range_name='A1')

print(">> PROCESO COMPLETADO EXITOSAMENTE.")