import pandas as pd
import requests
from PIL import Image, ImageDraw, ImageFont, ImageOps
from io import BytesIO
import os
import gspread
import json
import textwrap
import urllib.parse
import time
from datetime import datetime, timedelta
from oauth2client.service_account import ServiceAccountCredentials
from concurrent.futures import ThreadPoolExecutor

# --- CONFIGURACIÓN DE LIENZO ---
ANCHO, ALTO = 2500, 3750
SHEET_ID = "1NQdhnPxgVe6N6LiVxh1ouzt5NHtqjR22EEqL6w1RpWQ"
USUARIO_GITHUB = "analyticsdatajg2025-cmd" 
REPO_NOMBRE = "GITHUB_LENTO-MOVIMIENTO_CONECTA"
URL_BASE_PAGES = f"https://{USUARIO_GITHUB}.github.io/{REPO_NOMBRE}/flyers/"

# --- RUTAS DE FUENTES ---
FONT_BOLD_COND = "Mark Simonson - Proxima Nova Alt Condensed Bold.otf"
FONT_EXTRABOLD_COND = "Mark Simonson - Proxima Nova Alt Condensed Extrabold.otf"
FONT_REGULAR_COND = "Mark Simonson - Proxima Nova Alt Condensed Regular.otf"
FONT_EXTRABOLD = "Mark Simonson - Proxima Nova Extrabold.otf"
FONT_SEMIBOLD = "Mark Simonson - Proxima Nova Semibold.otf"

# --- COLORES ---
LC_AMARILLO = (255, 203, 5)
LC_AMARILLO_OSCURO = (235, 180, 0)
EFE_AZUL = (0, 107, 213) 
EFE_AZUL_OSCURO = (0, 60, 150)
EFE_NARANJA = (255, 100, 0)
BLANCO = (255, 255, 255)
NEGRO = (0, 0, 0)
GRIS_MARCA = (100, 100, 100)

output_dir = "docs/flyers"
os.makedirs(output_dir, exist_ok=True)

ahora_peru = datetime.utcnow() - timedelta(hours=5)
fecha_peru = ahora_peru.strftime("%d/%m/%Y %I:%M %p")

def conectar_sheets():
    info_creds = json.loads(os.environ['GOOGLE_SHEETS_JSON'])
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(info_creds, scope)
    client = gspread.authorize(creds)
    return client.open_by_key(SHEET_ID)

def descargar_imagen(url):
    if not url or str(url).lower() == 'nan' or str(url).strip() == "": return None
    try:
        headers = {"User-Agent": "Mozilla/5.0"}
        res = requests.get(url, headers=headers, timeout=10)
        return Image.open(BytesIO(res.content)).convert("RGBA")
    except: return None

def crear_flyer(productos, tienda_nombre, flyer_count):
    es_efe = "EFE" in tienda_nombre.upper()
    color_fondo = EFE_AZUL_OSCURO if es_efe else LC_AMARILLO_OSCURO
    color_slogan_bg = EFE_AZUL if es_efe else LC_AMARILLO
    logo_path = "logo-efe-sin-fondo.png" if es_efe else "logo-lc-sin-fondo.png"
    tienda_bg_path = "efe tienda.jpg" if es_efe else "LC-MIRAFLORES-LOGO-3D[2].jpg"
    
    flyer = Image.new('RGB', (ANCHO, ALTO), color=color_fondo)
    draw = ImageDraw.Draw(flyer)
    
    # Header Background
    header_h = 1000
    try:
        bg = Image.open(tienda_bg_path).convert("RGBA")
        bg = ImageOps.fit(bg, (ANCHO, header_h), method=Image.Resampling.LANCZOS)
        overlay = Image.new('RGBA', (ANCHO, header_h), (0, 0, 0, 60))
        bg.paste(overlay, (0, 0), overlay)
        flyer.paste(bg, (0, 0))
    except: pass

    # Logo Logic
    try:
        logo = Image.open(logo_path).convert("RGBA")
        if es_efe:
            diametro = 460
            c_x, c_y = ANCHO - diametro - 80, 40
            draw.ellipse([c_x, c_y, c_x + diametro, c_y + diametro], fill=BLANCO)
            logo_w = int(diametro * 0.85)
            logo = ImageOps.contain(logo, (logo_w, logo_w), method=Image.Resampling.LANCZOS)
            flyer.paste(logo, (c_x + (diametro - logo.width) // 2, c_y + (diametro - logo.height) // 2), logo)
        else:
            c_ancho, c_alto = 500, 380
            c_x, c_y = ANCHO - c_ancho - 80, 0
            draw.rounded_rectangle([c_x, c_y, c_x + c_ancho, c_y + c_alto], radius=50, fill=BLANCO)
            logo = ImageOps.contain(logo, (int(c_ancho*0.85), int(c_alto*0.80)), method=Image.Resampling.LANCZOS)
            flyer.paste(logo, (c_x + (c_ancho - logo.width) // 2, c_y + (c_alto - logo.height) // 2 + 10), logo)
    except: pass

    # Tienda Name
    f_tienda = ImageFont.truetype(FONT_EXTRABOLD_COND, 90)
    txt_tienda = tienda_nombre.upper()
    tw_t = draw.textlength(txt_tienda, font=f_tienda)
    if es_efe:
        draw.rounded_rectangle([ANCHO - tw_t - 150, 620, ANCHO, 800], radius=50, fill=EFE_NARANJA)
        draw.text((ANCHO - tw_t - 80, 655), txt_tienda, font=f_tienda, fill=BLANCO)
    else:
        p_x = ANCHO - tw_t - 250
        draw.polygon([(p_x, 720), (p_x + 100, 520), (ANCHO, 520), (ANCHO, 720)], fill=NEGRO)
        draw.text((ANCHO - tw_t - 100, 570), txt_tienda, font=f_tienda, fill=LC_AMARILLO)

    # Fecha Generación
    f_fecha = ImageFont.truetype(FONT_BOLD_COND, 45)
    txt_gen = f"Generado: {fecha_peru}"
    tw_g = draw.textlength(txt_gen, font=f_fecha)
    draw.rounded_rectangle([0, 850, tw_g + 80, 960], radius=40, fill=BLANCO)
    draw.text((40, 880), txt_gen, font=f_fecha, fill=NEGRO)

    # Slogan
    f_slogan = ImageFont.truetype(FONT_EXTRABOLD, 105)
    slogan_txt = "¡APROVECHA ESTAS INCREÍBLES OFERTAS!"
    sw = draw.textlength(slogan_txt, font=f_slogan)
    draw.rectangle([0, 1030, ANCHO, 1260], fill=color_slogan_bg)
    draw.text(((ANCHO-sw)//2, 1085), slogan_txt, font=f_slogan, fill=BLANCO if es_efe else NEGRO)

    # Grilla de Productos (3 filas x 2 columnas)
    anchos = [110, 1300]
    altos = [1350, 2150, 2950] 
    f_marca_prod = ImageFont.truetype(FONT_SEMIBOLD, 55)
    f_sku_prod = ImageFont.truetype(FONT_BOLD_COND, 60)

    for i, prod in enumerate(productos):
        if i >= 6: break
        x, y = anchos[i%2], altos[i//2]
        draw.rounded_rectangle([x, y, x+1090, y+760], radius=70, fill=BLANCO)
        
        img_p = descargar_imagen(prod.get('image_link'))
        if img_p:
            img_p.thumbnail((550, 550))
            flyer.paste(img_p, (x+30, y + (760-img_p.height)//2), img_p)

        tx = x + 600
        area_texto_w = 450
        
        # Marca
        marca = str(prod['Marca']).upper()
        draw.text((tx, y+80), marca, font=f_marca_prod, fill=GRIS_MARCA)
        
        # Título Artículo
        titulo = str(prod['Articulo'])
        f_size = 65
        f_art_prod = ImageFont.truetype(FONT_REGULAR_COND, f_size)
        lines = textwrap.wrap(titulo, width=15)
        
        ty = y + 160
        for line in lines[:4]:
            draw.text((tx, ty), line, font=f_art_prod, fill=NEGRO)
            ty += f_size + 5
            
        # SKU (Ubicado en la parte inferior del recuadro blanco)
        sku_val = str(prod['SKU'])
        rec_color_sku = EFE_NARANJA if es_efe else NEGRO
        draw.rounded_rectangle([tx - 20, y + 620, tx + area_texto_w, y + 710], radius=20, fill=rec_color_sku)
        tw_sku = draw.textlength(sku_val, font=f_sku_prod)
        draw.text((tx - 20 + (area_texto_w + 20 - tw_sku)//2, y + 635), sku_val, font=f_sku_prod, fill=BLANCO)

    return flyer

def procesar_tienda(nombre_tienda, grupo):
    print(f"Generando PDF: {nombre_tienda}")
    paginas = []
    
    indices = grupo.index.tolist()
    for i in range(0, len(indices), 6):
        bloque = grupo.iloc[i:i+6].to_dict('records')
        img_f = crear_flyer(bloque, str(nombre_tienda), (i//6)+1)
        paginas.append(img_f.convert("RGB"))
    
    if paginas:
        t_clean = "".join(x for x in str(nombre_tienda) if x.isalnum() or x in " -_")
        pdf_fn = f"LENTO_{t_clean}.pdf"
        pdf_path = os.path.join(output_dir, pdf_fn)
        paginas[0].save(pdf_path, save_all=True, append_images=paginas[1:])
        return [nombre_tienda, f"{URL_BASE_PAGES}{urllib.parse.quote(pdf_fn)}"]
    return None

# --- FLUJO PRINCIPAL ---
ss = conectar_sheets()

print("Cargando datos de Lento Movimiento...")
# Cargamos la hoja específica y el listado de productos para las imágenes
ws_lento = ss.worksheet("Lento_Movimiento")
df_source = pd.DataFrame(ws_lento.get_all_records())

# Mapeo de columnas según tu especificación (G=Tienda, C=Marca, D=SKU, E=Articulo)
# Nota: get_all_records usa los encabezados, asegúrate que coincidan con estos nombres:
df_source = df_source.rename(columns={
    df_source.columns[6]: 'Tienda',
    df_source.columns[2]: 'Marca',
    df_source.columns[3]: 'SKU',
    df_source.columns[4]: 'Articulo'
})

# Limpieza de SKU para cruce de imágenes
df_source['SKU_CLEAN'] = df_source['SKU'].astype(str).str.replace('-EX', '', case=False).str.strip()

print("Cruzando con listado_productos para obtener imágenes...")
df_lookup = pd.DataFrame(ss.worksheet("listado_productos").get_all_records())
lookup_dict = df_lookup.set_index('sku')['base_image_path'].to_dict()
df_source['image_link'] = df_source['SKU_CLEAN'].map(lookup_dict).fillna('')

# Generar PDFs por Tienda
grupos = df_source.groupby('Tienda')
tienda_links_pdf = []

with ThreadPoolExecutor(max_workers=4) as executor:
    futuros = [executor.submit(procesar_tienda, n, g) for n, g in grupos if str(n).strip()]
    for f in futuros:
        res = f.result()
        if res: tienda_links_pdf.append(res)

# Actualizar Tabla Maestra FLYER_TIENDA
print("Actualizando hoja FLYER_TIENDA...")
try:
    hoja_pdf = ss.worksheet("FLYER_TIENDA")
except:
    hoja_pdf = ss.add_worksheet(title="FLYER_TIENDA", rows="100", cols="2")

hoja_pdf.clear()
hoja_pdf.update(values=[["TIENDA RETAIL", "LINK PDF LENTO MOVIMIENTO"]] + tienda_links_pdf, range_name='A1')

print("¡Proceso de Lento Movimiento completado!")
