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
from datetime import datetime, timedelta
from oauth2client.service_account import ServiceAccountCredentials
from concurrent.futures import ThreadPoolExecutor

# --- CONFIGURACIÓN ---
ANCHO, ALTO = 2500, 3750
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

# Cache global en memoria para esta ejecución
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
    print(">> Conectando a Google Sheets...")
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
        res = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=10)
        img = Image.open(BytesIO(res.content)).convert("RGBA")
        img.thumbnail((600, 600))
        img.save(fname)
        cache_memoria[url] = fname
    except: cache_memoria[url] = None

def crear_flyer(productos, tienda_nombre, num_pag):
    es_efe = "EFE" in tienda_nombre.upper()
    color_fondo = EFE_AZUL_OSCURO if es_efe else LC_AMARILLO_OSCURO
    color_slogan_bg = EFE_AZUL if es_efe else LC_AMARILLO
    
    flyer = Image.new('RGB', (ANCHO, ALTO), color=color_fondo)
    draw = ImageDraw.Draw(flyer)
    
    # Recursos estáticos
    try:
        bg_p = "efe tienda.jpg" if es_efe else "LC-MIRAFLORES-LOGO-3D[2].jpg"
        with Image.open(bg_p) as b:
            flyer.paste(ImageOps.fit(b.convert("RGBA"), (ANCHO, 1000)), (0, 0))
        logo_p = "logo-efe-sin-fondo.png" if es_efe else "logo-lc-sin-fondo.png"
        with Image.open(logo_p) as l:
            logo = l.convert("RGBA")
            if es_efe:
                draw.ellipse([ANCHO-540, 40, ANCHO-80, 500], fill=BLANCO)
                flyer.paste(ImageOps.contain(logo, (390, 390)), (ANCHO-540+(460-logo.width)//2, 40+(460-logo.height)//2), logo)
            else:
                draw.rounded_rectangle([ANCHO-580, 0, ANCHO-80, 380], radius=50, fill=BLANCO)
                flyer.paste(ImageOps.contain(logo, (425, 300)), (ANCHO-580+(500-logo.width)//2, 50), logo)
    except: pass

    # Tienda y Fecha
    f_t = ImageFont.truetype(FONT_EXTRABOLD_COND, 90)
    txt_t = f"{tienda_nombre.upper()} - PÁG {num_pag}"
    draw.text((100, 650), txt_t, font=f_t, fill=BLANCO)
    draw.text((40, 880), f"Generado: {fecha_peru}", font=ImageFont.truetype(FONT_BOLD_COND, 40), fill=BLANCO)

    # Slogan
    draw.rectangle([0, 1030, ANCHO, 1260], fill=color_slogan_bg)
    draw.text((ANCHO//2, 1145), "¡APROVECHA ESTAS INCREÍBLES OFERTAS!", font=ImageFont.truetype(FONT_EXTRABOLD, 105), fill=BLANCO if es_efe else NEGRO, anchor="mm")

    # Grilla
    anchos, altos = [110, 1300], [1350, 2150, 2950]
    for i, prod in enumerate(productos):
        x, y = anchos[i%2], altos[i//2]
        draw.rounded_rectangle([x, y, x+1090, y+760], radius=70, fill=BLANCO)
        
        # STOCK (Mantiene formato original 6.180)
        stock_val = str(prod.get('Stock LM', '0'))
        draw.rounded_rectangle([x+30, y+30, x+320, y+140], radius=15, fill=EFE_AZUL if es_efe else LC_AMARILLO)
        draw.text((x+175, y+55), "STOCK", font=ImageFont.truetype(FONT_BOLD_COND, 30), fill=BLANCO if es_efe else NEGRO, anchor="mm")
        draw.text((x+175, y+100), stock_val, font=ImageFont.truetype(FONT_EXTRABOLD, 50), fill=BLANCO if es_efe else NEGRO, anchor="mm")

        # Imagen desde Cache
        path_img = cache_memoria.get(prod.get('image_link'))
        if path_img and os.path.exists(path_img):
            with Image.open(path_img) as img:
                flyer.paste(img, (x+40, y+180), img)

        tx, area_w = x + 570, 480
        draw.text((tx, y+50), str(prod.get('Marca', '')).upper(), font=ImageFont.truetype(FONT_SEMIBOLD, 50), fill=GRIS_MARCA)
        
        lines = textwrap.wrap(str(prod.get('Nombre Articulo', '')), width=18)
        ty = y + 110
        for line in lines[:3]:
            draw.text((tx, ty), line, font=ImageFont.truetype(FONT_REGULAR_COND, 60), fill=NEGRO)
            ty += 65

        # PRECIO (Sin decimales)
        ty_p = y + 420
        draw.rounded_rectangle([tx, ty_p, tx + area_w, ty_p + 180], radius=25, fill=color_slogan_bg)
        
        p_raw = str(prod.get('Precio Vigente', ''))
        p_clean = p_raw.replace(".00", "").strip()
        
        if p_clean in ["", "nan", "0", "0.0", "SIN PRECIO"]:
            draw.text((tx + area_w//2, ty_p + 90), "SIN PRECIO", font=ImageFont.truetype(FONT_EXTRABOLD, 80), fill=BLANCO if es_efe else NEGRO, anchor="mm")
        else:
            draw.text((tx + area_w//2, ty_p + 90), f"S/ {p_clean}", font=ImageFont.truetype(FONT_EXTRABOLD, 110), fill=BLANCO if es_efe else NEGRO, anchor="mm")

        # SKU
        sku_c = NEGRO if not es_efe else EFE_NARANJA
        draw.rounded_rectangle([tx, ty_p+180, tx+area_w, ty_p+265], radius=20, fill=sku_c)
        draw.text((tx+area_w//2, ty_p+222), str(prod['SKU']), font=ImageFont.truetype(FONT_BOLD_COND, 55), fill=BLANCO, anchor="mm")

    return flyer

# --- FLUJO ---
ss = conectar_sheets()

print(">> PARTE 1: Sincronización...")
# (Aquí va tu lógica de promos_dict, txl_dict y df_origen igual que antes)
# ... [Omitido para ahorrar espacio, pero mantén tu lógica de cruce de datos aquí] ...
# Carga rápida con Diccionarios
promos_dict = {}
for p in ["Promo01", "Promo03", "Promo04"]:
    df_p = pd.DataFrame(ss.worksheet(p).get_all_records())
    df_p.columns = df_p.columns.str.strip()
    if 'Lista Precios' in df_p.columns:
        df_p['KEY'] = df_p['Lista Precios'].astype(str).str.replace(".0","", regex=False) + "_" + df_p['SKU'].astype(str)
        promos_dict.update(df_p.set_index('KEY')['Precio Vigente'].to_dict())

df_txl = pd.DataFrame(ss.worksheet("TiendasxLista").get_all_records())
txl_dict = {normalizar_nombre_tienda(r['TIENDA']): str(r['LISTA']).replace(".0","") for r in df_txl.to_dict('records') if 'TIENDA' in r}

df_raw = pd.DataFrame(ss.worksheet("Origen Tdas").get_all_records())
df_origen = pd.DataFrame({'Semana': df_raw.iloc[:, 1], 'Tienda': df_raw.iloc[:, 3], 'Marca': df_raw.iloc[:, 6], 'SKU': df_raw.iloc[:, 7], 'Nombre Articulo': df_raw.iloc[:, 8], 'Stock LM': df_raw.iloc[:, 11]})
df_origen['LISTA'] = df_origen['Tienda'].apply(normalizar_nombre_tienda).map(txl_dict).fillna("")
df_origen['Precio Vigente'] = (df_origen['LISTA'] + "_" + df_origen['SKU'].astype(str)).map(promos_dict).fillna("SIN PRECIO")

df_lookup = pd.DataFrame(ss.worksheet("listado_productos").get_all_records())
img_dict = df_lookup.set_index('sku')['base_image_path'].to_dict()
df_origen['image_link'] = df_origen['SKU'].astype(str).str.replace('-EX', '', case=False).map(img_dict).fillna('')

df_final = df_origen[df_origen['Semana'].astype(str) == semana_actual].copy()
ws_det = ss.worksheet("Detalle de Inventario")
ws_det.clear()
ws_det.update([df_final.columns.tolist()] + df_final.astype(str).values.tolist(), range_name='A1')

print(">> PARTE 2: Pre-descarga de imágenes...")
urls = df_final['image_link'].unique()
with ThreadPoolExecutor(max_workers=20) as exe:
    exe.map(descargar_y_cachear, urls)

print(">> PARTE 3: Renderizado...")
tienda_links = []
df_render = pd.DataFrame(ws_det.get_all_records())

for nombre, grupo in df_render.groupby('Tienda'):
    prods = grupo.to_dict('records')
    paginas = []
    for i in range(0, len(prods), 6):
        paginas.append(crear_flyer(prods[i:i+6], str(nombre), (i//6)+1).convert("RGB"))
    
    if paginas:
        clean = "".join(c for c in str(nombre) if c.isalnum() or c in " -_").strip().replace(" ", "_")
        fn = f"LENTO_{clean}.pdf"
        # CALIDAD 30 y OPTIMIZE para evitar que Git se cuelgue al subir
        paginas[0].save(os.path.join(output_dir, fn), save_all=True, append_images=paginas[1:], quality=30, optimize=True)
        tienda_links.append([nombre, f"{URL_BASE_PAGES}view.html?file={urllib.parse.quote(fn)}"])
        for p in paginas: p.close()
        gc.collect()

ss.worksheet("FLYER_TIENDA").clear()
ss.worksheet("FLYER_TIENDA").update([["TIENDA RETAIL", "LINK PDF LENTO MOVIMIENTO"]] + tienda_links, range_name='A1')
print(">> PROCESO COMPLETADO.")