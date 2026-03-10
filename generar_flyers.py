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
os.makedirs(output_dir, exist_ok=True)

ahora_peru = datetime.utcnow() - timedelta(hours=5)
fecha_peru = ahora_peru.strftime("%d/%m/%Y %I:%M %p")
semana_actual = f"Sem{ahora_peru.isocalendar()[1]}"

cache_imagenes = {}

# Carga de fuentes una sola vez para ahorrar tiempo
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

def pre_descargar_una_imagen(url):
    if not url or str(url).lower() in ['nan', ''] or url in cache_imagenes:
        return
    try:
        res = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=5)
        img = Image.open(BytesIO(res.content)).convert("RGBA")
        img.thumbnail((600, 600)) # Tamaño óptimo para PDF y RAM
        cache_imagenes[url] = img
    except:
        cache_imagenes[url] = None

def normalizar_nombre_tienda(nombre):
    s = str(nombre).upper().replace(" ", "").replace("-", "")
    if s.endswith("EFE"): s = "EFE" + s[:-3]
    if s.endswith("LC"): s = "LC" + s[:-2]
    return s

def crear_pagina_flyer(productos, tienda_nombre, num_pag):
    es_efe = "EFE" in tienda_nombre.upper()
    color_fondo = EFE_AZUL_OSCURO if es_efe else LC_AMARILLO_OSCURO
    color_slogan_bg = EFE_AZUL if es_efe else LC_AMARILLO
    
    flyer = Image.new('RGB', (ANCHO, ALTO), color=color_fondo)
    draw = ImageDraw.Draw(flyer)
    
    # Optimizamos carga de recursos estáticos
    try:
        bg_path = "efe tienda.jpg" if es_efe else "LC-MIRAFLORES-LOGO-3D[2].jpg"
        with Image.open(bg_path) as bg_raw:
            bg = ImageOps.fit(bg_raw.convert("RGBA"), (ANCHO, 1000))
            flyer.paste(bg, (0, 0))
        
        logo_path = "logo-efe-sin-fondo.png" if es_efe else "logo-lc-sin-fondo.png"
        with Image.open(logo_path) as logo_raw:
            logo = logo_raw.convert("RGBA")
            if es_efe:
                draw.ellipse([ANCHO-540, 40, ANCHO-80, 500], fill=BLANCO)
                logo = ImageOps.contain(logo, (390, 390))
                flyer.paste(logo, (ANCHO-540+(460-logo.width)//2, 40+(460-logo.height)//2), logo)
            else:
                draw.rounded_rectangle([ANCHO-580, 0, ANCHO-80, 380], radius=50, fill=BLANCO)
                logo = ImageOps.contain(logo, (425, 300))
                flyer.paste(logo, (ANCHO-580+(500-logo.width)//2, (380-logo.height)//2 + 10), logo)
    except: pass

    # Textos de cabecera
    f_t = ImageFont.truetype(FONT_EXTRABOLD_COND, 90)
    txt_t = f"{tienda_nombre.upper()} - PÁG {num_pag}"
    if es_efe:
        draw.rounded_rectangle([ANCHO-draw.textlength(txt_t, f_t)-150, 620, ANCHO, 800], radius=50, fill=EFE_NARANJA)
        draw.text((ANCHO-draw.textlength(txt_t, f_t)-80, 655), txt_t, font=f_t, fill=BLANCO)
    else:
        draw.polygon([(ANCHO-draw.textlength(txt_t, f_t)-250, 720), (ANCHO-draw.textlength(txt_t, f_t)-150, 520), (ANCHO, 520), (ANCHO, 720)], fill=NEGRO)
        draw.text((ANCHO-draw.textlength(txt_t, f_t)-100, 570), txt_t, font=f_t, fill=LC_AMARILLO)

    draw.text((40, 880), f"Generado: {fecha_peru}", font=ImageFont.truetype(FONT_BOLD_COND, 40), fill=BLANCO)

    # Slogan
    draw.rectangle([0, 1030, ANCHO, 1260], fill=color_slogan_bg)
    slogan = "¡APROVECHA ESTAS INCREÍBLES OFERTAS!"
    f_s = ImageFont.truetype(FONT_EXTRABOLD, 105)
    draw.text(((ANCHO-draw.textlength(slogan, f_s))//2, 1085), slogan, font=f_s, fill=BLANCO if es_efe else NEGRO)

    # Grilla de 6 productos
    anchos, altos = [110, 1300], [1350, 2150, 2950]
    for i, prod in enumerate(productos):
        x, y = anchos[i%2], altos[i//2]
        draw.rounded_rectangle([x, y, x+1090, y+760], radius=70, fill=BLANCO)
        
        # Stock
        stock_val = str(prod.get('Stock LM', '0'))
        f_sl, f_sn = ImageFont.truetype(FONT_BOLD_COND, 30), ImageFont.truetype(FONT_EXTRABOLD, 50)
        draw.rounded_rectangle([x+30, y+30, x+250, y+140], radius=15, fill=EFE_AZUL if es_efe else LC_AMARILLO)
        draw.text((x+50, y+40), "STOCK", font=f_sl, fill=BLANCO if es_efe else NEGRO)
        draw.text((x+50, y+75), stock_val, font=f_sn, fill=BLANCO if es_efe else NEGRO)

        # Imagen desde caché
        img = cache_imagenes.get(prod.get('image_link'))
        if img:
            flyer.paste(img, (x+40, y+150), img)

        tx, area_w = x + 570, 480
        draw.text((tx, y+50), str(prod['Marca']).upper(), font=ImageFont.truetype(FONT_SEMIBOLD, 50), fill=GRIS_MARCA)
        
        lines = textwrap.wrap(str(prod['Nombre Articulo']), width=18)
        ty = y + 110
        for line in lines[:3]:
            draw.text((tx, ty), line, font=ImageFont.truetype(FONT_REGULAR_COND, 60), fill=NEGRO)
            ty += 65

        # Precio
        ty_p = y + 420
        draw.rounded_rectangle([tx, ty_p, tx + area_w, ty_p + 180], radius=25, fill=color_slogan_bg)
        precio = str(prod.get('Precio Vigente', '0')).replace(".00", "").replace("nan", "SIN PRECIO")
        f_p = ImageFont.truetype(FONT_EXTRABOLD, 110)
        draw.text((tx + 20, ty_p+55), f"S/ {precio}" if "SIN" not in precio else precio, font=f_p, fill=BLANCO if es_efe else NEGRO)

        # SKU
        draw.rounded_rectangle([tx, ty_p+180, tx+area_w, ty_p+265], radius=20, fill=NEGRO if not es_efe else EFE_NARANJA)
        draw.text((tx+40, ty_p+195), str(prod['SKU']), font=ImageFont.truetype(FONT_BOLD_COND, 55), fill=BLANCO)

    return flyer

def procesar_tienda_multipagina(nombre, df_tienda):
    try:
        paginas = []
        lista_prods = df_tienda.to_dict('records')
        for i in range(0, len(lista_prods), 6):
            paginas.append(crear_pagina_flyer(lista_prods[i:i+6], str(nombre), (i//6)+1).convert("RGB"))
        
        if paginas:
            clean_name = "".join(c for c in nombre if c.isalnum() or c in " -_").strip().replace(" ", "_")
            fn = f"LENTO_{clean_name}.pdf"
            path = os.path.join(output_dir, fn)
            paginas[0].save(path, save_all=True, append_images=paginas[1:], quality=60, optimize=True)
            # Liberar RAM de las imágenes
            for p in paginas: p.close()
            return [nombre, f"{URL_BASE_PAGES}view.html?file={urllib.parse.quote(fn)}"]
    except Exception as e:
        print(f"Error en {nombre}: {e}")
    return None

# --- FLUJO ---
ss = conectar_sheets()
print(">> Cargando datos...")
# ... (Bloque de carga de datos consolidado para velocidad)
df_raw = pd.DataFrame(ss.worksheet("Origen Tdas").get_all_records())
df_txl = pd.DataFrame(ss.worksheet("TiendasxLista").get_all_records())
df_final = df_raw[df_raw.iloc[:, 1].astype(str) == semana_actual].copy() # Filtro directo

print(">> Pre-descarga de imágenes...")
urls_unicas = df_final.iloc[:, 11].unique() # Asumiendo columna de link
with ThreadPoolExecutor(max_workers=20) as executor:
    executor.map(pre_descargar_una_imagen, urls_unicas)

print(f">> Generando PDFs para {len(df_final.groupby(df_final.columns[3]))} tiendas...")
tienda_links = []
# PARALELISMO CLAVE: Usamos workers para procesar varias tiendas a la vez
with ThreadPoolExecutor(max_workers=4) as executor:
    tareas = {executor.submit(procesar_tienda_multipagina, n, g): n for n, g in df_final.groupby(df_final.columns[3])}
    for future in tareas:
        res = future.result()
        if res: tienda_links.append(res)
        gc.collect()

ss.worksheet("FLYER_TIENDA").clear()
ss.worksheet("FLYER_TIENDA").update([["TIENDA RETAIL", "LINK PDF LENTO MOVIMIENTO"]] + tienda_links, range_name='A1')
print(">> COMPLETO EN TIEMPO RÉCORD.")
