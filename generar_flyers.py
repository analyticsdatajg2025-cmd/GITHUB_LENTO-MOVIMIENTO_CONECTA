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

# --- CONFIGURACIÓN BALANCEADA (Calidad vs Peso) ---
ANCHO, ALTO = 1600, 2400
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

# Fuentes
FONT_BOLD_COND = "Mark Simonson - Proxima Nova Alt Condensed Bold.otf"
FONT_EXTRABOLD_COND = "Mark Simonson - Proxima Nova Alt Condensed Extrabold.otf"
FONT_REGULAR_COND = "Mark Simonson - Proxima Nova Alt Condensed Regular.otf"
FONT_EXTRABOLD = "Mark Simonson - Proxima Nova Extrabold.otf"
FONT_SEMIBOLD = "Mark Simonson - Proxima Nova Semibold.otf"

# Colores
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
        img.thumbnail((400, 400)) 
        img.save(fname, "PNG")
        cache_memoria[url] = fname
    except: cache_memoria[url] = None

def crear_flyer(productos, tienda_nombre, num_pag):
    es_efe = "EFE" in tienda_nombre.upper()
    color_fondo = EFE_AZUL_OSCURO if es_efe else LC_AMARILLO_OSCURO
    color_slogan_bg = EFE_AZUL if es_efe else LC_AMARILLO
    
    flyer = Image.new('RGB', (ANCHO, ALTO), color=color_fondo)
    draw = ImageDraw.Draw(flyer)
    
    # --- HEADER DESIGN ---
    try:
        bg_p = "efe tienda.jpg" if es_efe else "LC-MIRAFLORES-LOGO-3D[2].jpg"
        with Image.open(bg_p) as b:
            flyer.paste(ImageOps.fit(b.convert("RGBA"), (ANCHO, 650)), (0, 0))
        
        logo_p = "logo-efe-sin-fondo.png" if es_efe else "logo-lc-sin-fondo.png"
        with Image.open(logo_p).convert("RGBA") as l:
            if es_efe:
                draw.ellipse([ANCHO-380, 40, ANCHO-60, 360], fill=BLANCO)
                logo = ImageOps.contain(l, (260, 260))
                flyer.paste(logo, (ANCHO-380+(320-logo.width)//2, 40+(320-logo.height)//2), logo)
            else:
                draw.rounded_rectangle([ANCHO-420, 20, ANCHO-40, 300], radius=40, fill=BLANCO)
                logo = ImageOps.contain(l, (320, 220))
                flyer.paste(logo, (ANCHO-420+(380-logo.width)//2, 40), logo)
    except: pass

    # --- NOMBRE TIENDA (RECUPERADO) ---
    f_tienda = ImageFont.truetype(FONT_EXTRABOLD_COND, 70)
    txt_tienda = tienda_nombre.upper()
    tw_t = draw.textlength(txt_tienda, font=f_tienda)
    
    if es_efe:
        draw.rounded_rectangle([ANCHO - tw_t - 120, 420, ANCHO, 550], radius=30, fill=EFE_NARANJA)
        draw.text((ANCHO - tw_t - 60, 440), txt_tienda, font=f_tienda, fill=BLANCO)
    else:
        p_x = ANCHO - tw_t - 200
        draw.polygon([(p_x, 520), (p_x + 80, 400), (ANCHO, 400), (ANCHO, 520)], fill=NEGRO)
        draw.text((ANCHO - tw_t - 80, 420), txt_tienda, font=f_tienda, fill=LC_AMARILLO)

    # --- FECHA GENERACIÓN ---
    f_fecha = ImageFont.truetype(FONT_BOLD_COND, 35)
    txt_gen = f"Generado: {fecha_peru} - PÁG {num_pag}"
    draw.rounded_rectangle([20, 570, 650, 630], radius=20, fill=BLANCO)
    draw.text((40, 582), txt_gen, font=f_fecha, fill=NEGRO)

    # --- SLOGAN ---
    draw.rectangle([0, 650, ANCHO, 800], fill=color_slogan_bg)
    draw.text((ANCHO//2, 725), "¡APROVECHA ESTAS INCREÍBLES OFERTAS!", font=ImageFont.truetype(FONT_EXTRABOLD, 68), fill=BLANCO if es_efe else NEGRO, anchor="mm")

    # --- GRILLA DE PRODUCTOS ---
    anchos, altos = [60, 840], [840, 1340, 1840]
    for i, prod in enumerate(productos):
        x, y = anchos[i%2], altos[i//2]
        draw.rounded_rectangle([x, y, x+700, y+470], radius=50, fill=BLANCO)
        
        # STOCK (Lógica punto decimal + Sin Stock)
        stock_raw = str(prod.get('Stock LM', '0')).replace(".", "").strip()
        if stock_raw in ["0", "", "nan"]: stock_txt, color_st = "SIN STOCK", GRIS_MARCA
        else: stock_txt, color_st = f"STOCK: {stock_raw}", (EFE_AZUL if es_efe else LC_AMARILLO)
        
        draw.rounded_rectangle([x+20, y+20, x+250, y+80], radius=15, fill=color_st)
        draw.text((x+135, y+50), stock_txt, font=ImageFont.truetype(FONT_BOLD_COND, 24), fill=BLANCO if es_efe else NEGRO, anchor="mm")

        # Imagen Producto (Ajustada para no chocar)
        path_img = cache_memoria.get(prod.get('image_link'))
        if path_img:
            with Image.open(path_img) as img:
                flyer.paste(img, (x+20, y+100), img)

        tx, area_w = x + 340, 330
        draw.text((tx, y+30), str(prod.get('Marca', '')).upper(), font=ImageFont.truetype(FONT_SEMIBOLD, 30), fill=GRIS_MARCA)
        
        # Nombre Artículo
        lines = textwrap.wrap(str(prod.get('Nombre Articulo', '')), width=18)
        ty = y + 75
        for line in lines[:3]:
            draw.text((tx, ty), line, font=ImageFont.truetype(FONT_REGULAR_COND, 34), fill=NEGRO)
            ty += 38

        # --- BLOQUE UNIFICADO PRECIO + SKU ---
        ty_p = y + 250
        # Botón Precio (Arriba)
        draw.rounded_rectangle([tx, ty_p, tx + area_w, ty_p + 100], radius=15, fill=color_slogan_bg)
        p_raw = str(prod.get('Precio Vigente', '0')).replace(".00", "").replace(",00", "").strip()
        
        if p_raw in ["0.0", "0", "", "nan", "SIN PRECIO"]:
            draw.text((tx + area_w//2, ty_p + 50), "SIN PRECIO", font=ImageFont.truetype(FONT_EXTRABOLD, 40), fill=BLANCO if es_efe else NEGRO, anchor="mm")
        else:
            draw.text((tx + area_w//2, ty_p + 50), f"S/ {p_raw}", font=ImageFont.truetype(FONT_EXTRABOLD, 60), fill=BLANCO if es_efe else NEGRO, anchor="mm")

        # Botón SKU (Abajo - Recto arriba para unir)
        sku_color = NEGRO if not es_efe else EFE_NARANJA
        draw.rounded_rectangle([tx, ty_p+100, tx+area_w, ty_p+160], radius=15, fill=sku_color)
        # Rectángulo tapando redondeo superior del SKU para unirlo al precio
        draw.rectangle([tx, ty_p+100, tx+area_w, ty_p+110], fill=sku_color) 
        draw.text((tx+area_w//2, ty_p+130), f"SKU: {prod['SKU']}", font=ImageFont.truetype(FONT_BOLD_COND, 30), fill=BLANCO, anchor="mm")

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
            # Calidad 50: Texto nítido, imagen comprimida
            paginas[0].save(os.path.join(output_dir, fn), save_all=True, append_images=paginas[1:], quality=50, optimize=True)
            for p in paginas: p.close()
            return [nombre, f"{URL_BASE_PAGES}view.html?file={urllib.parse.quote(fn)}"]
    except Exception as e:
        print(f"Error: {e}")
    return None

# --- FLUJO PRINCIPAL ---
ss = conectar_sheets()
df_raw = pd.DataFrame(ss.worksheet("Origen Tdas").get_all_records())
df_origen = pd.DataFrame({'Semana': df_raw.iloc[:, 1], 'Tienda': df_raw.iloc[:, 3], 'Marca': df_raw.iloc[:, 6], 'SKU': df_raw.iloc[:, 7], 'Nombre Articulo': df_raw.iloc[:, 8], 'Stock LM': df_raw.iloc[:, 11]})

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

# Guardar Detalle
try:
    ws_det = ss.worksheet("Detalle de Inventario")
    ws_det.clear()
    cols = ['Semana', 'Tienda', 'Marca', 'SKU', 'Nombre Articulo', 'Stock LM', 'LISTA', 'Precio Vigente', 'image_link']
    ws_det.update([cols] + df_final[cols].astype(str).values.tolist(), range_name='A1')
except: pass

with ThreadPoolExecutor(max_workers=30) as exe:
    exe.map(descargar_y_cachear, df_final['image_link'].unique())

tienda_links = []
with ThreadPoolExecutor(max_workers=4) as exe:
    resultados = list(exe.map(procesar_tienda_batch, df_final.groupby('Tienda')))
    tienda_links = [r for r in resultados if r]

ss.worksheet("FLYER_TIENDA").clear()
if tienda_links:
    ss.worksheet("FLYER_TIENDA").update([["TIENDA", "LINK"]] + tienda_links, range_name='A1')

print(">> PROCESO COMPLETADO.")