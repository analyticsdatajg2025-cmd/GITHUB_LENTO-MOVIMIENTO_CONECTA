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

def obtener_imagen_con_cache(url):
    if not url or str(url).lower() in ['nan', '']: return None
    
    # Crear nombre de archivo basado en el hash de la URL para el cache en disco
    filename = os.path.join(cache_dir, hashlib.md5(url.encode()).hexdigest() + ".png")
    
    if os.path.exists(filename):
        return Image.open(filename).convert("RGBA")
    
    try:
        res = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=10)
        img = Image.open(BytesIO(res.content)).convert("RGBA")
        img.thumbnail((600, 600))
        img.save(filename) # Guardar en disco para la siguiente tienda
        return img
    except: return None

import hashlib # Requerido para el nombre del cache

def crear_flyer(productos, tienda_nombre, num_pag):
    es_efe = "EFE" in tienda_nombre.upper()
    color_fondo = EFE_AZUL_OSCURO if es_efe else LC_AMARILLO_OSCURO
    color_slogan_bg = EFE_AZUL if es_efe else LC_AMARILLO
    
    flyer = Image.new('RGB', (ANCHO, ALTO), color=color_fondo)
    draw = ImageDraw.Draw(flyer)
    
    # Recursos Gráficos
    try:
        bg_p = "efe tienda.jpg" if es_efe else "LC-MIRAFLORES-LOGO-3D[2].jpg"
        with Image.open(bg_p) as b:
            flyer.paste(ImageOps.fit(b.convert("RGBA"), (ANCHO, 1000)), (0, 0))
        
        logo_p = "logo-efe-sin-fondo.png" if es_efe else "logo-lc-sin-fondo.png"
        with Image.open(logo_p) as l:
            logo = l.convert("RGBA")
            if es_efe:
                draw.ellipse([ANCHO-540, 40, ANCHO-80, 500], fill=BLANCO)
                logo = ImageOps.contain(logo, (390, 390))
                flyer.paste(logo, (ANCHO-540+(460-logo.width)//2, 40+(460-logo.height)//2), logo)
            else:
                draw.rounded_rectangle([ANCHO-580, 0, ANCHO-80, 380], radius=50, fill=BLANCO)
                logo = ImageOps.contain(logo, (425, 300))
                flyer.paste(logo, (ANCHO-580+(500-logo.width)//2, (380-logo.height)//2 + 10), logo)
    except: pass

    f_t = ImageFont.truetype(FONT_EXTRABOLD_COND, 90)
    txt_t = f"{tienda_nombre.upper()} - PÁG {num_pag}"
    if es_efe:
        draw.rounded_rectangle([ANCHO-draw.textlength(txt_t, f_t)-150, 620, ANCHO, 800], radius=50, fill=EFE_NARANJA)
        draw.text((ANCHO-draw.textlength(txt_t, f_t)-80, 655), txt_t, font=f_t, fill=BLANCO)
    else:
        draw.polygon([(ANCHO-draw.textlength(txt_t, f_t)-250, 720), (ANCHO-draw.textlength(txt_t, f_t)-150, 520), (ANCHO, 520), (ANCHO, 720)], fill=NEGRO)
        draw.text((ANCHO-draw.textlength(txt_t, f_t)-100, 570), txt_t, font=f_t, fill=LC_AMARILLO)

    draw.text((40, 880), f"Generado: {fecha_peru}", font=ImageFont.truetype(FONT_BOLD_COND, 40), fill=BLANCO)
    draw.rectangle([0, 1030, ANCHO, 1260], fill=color_slogan_bg)
    slogan = "¡APROVECHA ESTAS INCREÍBLES OFERTAS!"
    draw.text(((ANCHO-draw.textlength(slogan, ImageFont.truetype(FONT_EXTRABOLD, 105)))//2, 1085), slogan, font=ImageFont.truetype(FONT_EXTRABOLD, 105), fill=BLANCO if es_efe else NEGRO)

    anchos, altos = [110, 1300], [1350, 2150, 2950]
    for i, prod in enumerate(productos):
        x, y = anchos[i%2], altos[i//2]
        draw.rounded_rectangle([x, y, x+1090, y+760], radius=70, fill=BLANCO)
        
        # Stock
        stock_val = str(prod.get('Stock LM', '0'))
        draw.rounded_rectangle([x+30, y+30, x+280, y+140], radius=15, fill=EFE_AZUL if es_efe else LC_AMARILLO)
        draw.text((x+50, y+40), "STOCK", font=ImageFont.truetype(FONT_BOLD_COND, 30), fill=BLANCO if es_efe else NEGRO)
        draw.text((x+50, y+75), stock_val, font=ImageFont.truetype(FONT_EXTRABOLD, 50), fill=BLANCO if es_efe else NEGRO)

        # Imagen (Usando Cache en Disco)
        img = obtener_imagen_con_cache(prod.get('image_link'))
        if img:
            flyer.paste(img, (x+40, y+180), img)
            img.close()

        tx, area_w = x + 570, 480
        draw.text((tx, y+50), str(prod.get('Marca', '')).upper(), font=ImageFont.truetype(FONT_SEMIBOLD, 50), fill=GRIS_MARCA)
        lines = textwrap.wrap(str(prod.get('Nombre Articulo', '')), width=18)
        ty = y + 110
        for line in lines[:3]:
            draw.text((tx, ty), line, font=ImageFont.truetype(FONT_REGULAR_COND, 60), fill=NEGRO)
            ty += 65

        ty_p, h_p = y + 420, 180
        draw.rounded_rectangle([tx, ty_p, tx + area_w, ty_p + h_p], radius=25, fill=color_slogan_bg)
        p_val = str(prod.get('Precio Vigente', 'SIN PRECIO')).replace(".00", "").strip()
        draw.text((tx+20, ty_p+50), f"S/ {p_val}" if "SIN" not in p_val else p_val, font=ImageFont.truetype(FONT_EXTRABOLD, 110 if "SIN" not in p_val else 80), fill=BLANCO if es_efe else NEGRO)

        sku_c = NEGRO if not es_efe else EFE_NARANJA
        draw.rounded_rectangle([tx, ty_p+180, tx+area_w, ty_p+265], radius=20, fill=sku_c)
        draw.text((tx+40, ty_p+195), str(prod['SKU']), font=ImageFont.truetype(FONT_BOLD_COND, 55), fill=BLANCO)

    return flyer

# --- FLUJO DIVIDIDO ---
ss = conectar_sheets()

# PARTE 1: Sincronización de Datos (Rápida)
print(">> PARTE 1: Actualizando Base de Datos de Detalle...")
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

# LIBERAR MEMORIA DE DATOS PESADOS
del df_raw, df_origen, df_lookup, promos_dict, txl_dict, img_dict
gc.collect()

# PARTE 2: Diseño de PDFs usando la hoja recién llenada
print(">> PARTE 2: Iniciando Renderizado de PDFs por Tienda...")
tienda_links = []
# Leemos directamente de la hoja para trabajar con datos frescos y ligeros
df_render = pd.DataFrame(ws_det.get_all_records())

for nombre, grupo in df_render.groupby('Tienda'):
    try:
        prods = grupo.to_dict('records')
        paginas = []
        for i in range(0, len(prods), 6):
            paginas.append(crear_flyer(prods[i:i+6], str(nombre), (i//6)+1).convert("RGB"))
        
        if paginas:
            clean = "".join(c for c in str(nombre) if c.isalnum() or c in " -_").strip().replace(" ", "_")
            fn = f"LENTO_{clean}.pdf"
            paginas[0].save(os.path.join(output_dir, fn), save_all=True, append_images=paginas[1:], quality=40, optimize=True)
            print(f"   [OK] {nombre}")
            tienda_links.append([nombre, f"{URL_BASE_PAGES}view.html?file={urllib.parse.quote(fn)}"])
            
            for p in paginas: p.close()
            del paginas
            gc.collect() # Limpieza forzada tras cada PDF
    except Exception as e:
        print(f"Error en {nombre}: {e}")

# Actualizar Links en FLYER_TIENDA
ss.worksheet("FLYER_TIENDA").clear()
ss.worksheet("FLYER_TIENDA").update([["TIENDA RETAIL", "LINK PDF LENTO MOVIMIENTO"]] + tienda_links, range_name='A1')
print(">> PROCESO COMPLETADO EXITOSAMENTE.")