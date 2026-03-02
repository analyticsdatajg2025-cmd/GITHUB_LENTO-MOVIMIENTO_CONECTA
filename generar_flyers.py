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

# --- CONFIGURACIÓN ---
ANCHO, ALTO = 2500, 3750
SHEET_ID = "1NQdhnPxgVe6N6LiVxh1ouzt5NHtqjR22EEqL6w1RpWQ"
USUARIO_GITHUB = "analyticsdatajg2025-cmd" 
REPO_NOMBRE = "GITHUB_LENTO-MOVIMIENTO_CONECTA"
URL_BASE_PAGES = f"https://{USUARIO_GITHUB}.github.io/{REPO_NOMBRE}/flyers/"

# --- FUENTES ---
FONT_BOLD_COND = "Mark Simonson - Proxima Nova Alt Condensed Bold.otf"
FONT_EXTRABOLD_COND = "Mark Simonson - Proxima Nova Alt Condensed Extrabold.otf"
FONT_REGULAR_COND = "Mark Simonson - Proxima Nova Alt Condensed Regular.otf"
FONT_EXTRABOLD = "Mark Simonson - Proxima Nova Extrabold.otf"
FONT_SEMIBOLD = "Mark Simonson - Proxima Nova Semibold.otf"

# --- COLORES ---
LC_AMARILLO, LC_AMARILLO_OSCURO = (255, 203, 5), (235, 180, 0)
EFE_AZUL, EFE_AZUL_OSCURO = (0, 107, 213), (0, 60, 150)
EFE_NARANJA, BLANCO, NEGRO, GRIS_MARCA = (255, 100, 0), (255, 255, 255), (0, 0, 0), (100, 100, 100)

output_dir = "docs/flyers"
os.makedirs(output_dir, exist_ok=True)

ahora_peru = datetime.utcnow() - timedelta(hours=5)
fecha_peru = ahora_peru.strftime("%d/%m/%Y %I:%M %p")
semana_actual = f"Sem{ahora_peru.isocalendar()[1]}"

def conectar_sheets():
    info_creds = json.loads(os.environ['GOOGLE_SHEETS_JSON'])
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(info_creds, scope)
    client = gspread.authorize(creds)
    return client.open_by_key(SHEET_ID)

def descargar_imagen(url):
    if not url or str(url).lower() in ['nan', ''] : return None
    try:
        res = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=10)
        return Image.open(BytesIO(res.content)).convert("RGBA")
    except: return None

def crear_flyer(productos, tienda_nombre):
    es_efe = "EFE" in tienda_nombre.upper()
    color_fondo = EFE_AZUL_OSCURO if es_efe else LC_AMARILLO_OSCURO
    color_slogan_bg = EFE_AZUL if es_efe else LC_AMARILLO
    logo_path = "logo-efe-sin-fondo.png" if es_efe else "logo-lc-sin-fondo.png"
    tienda_bg_path = "efe tienda.jpg" if es_efe else "LC-MIRAFLORES-LOGO-3D[2].jpg"
    
    flyer = Image.new('RGB', (ANCHO, ALTO), color=color_fondo)
    draw = ImageDraw.Draw(flyer)
    
    # 1. Header Background
    try:
        bg = ImageOps.fit(Image.open(tienda_bg_path).convert("RGBA"), (ANCHO, 1000), method=Image.Resampling.LANCZOS)
        flyer.paste(bg, (0, 0))
        overlay = Image.new('RGBA', (ANCHO, 1000), (0, 0, 0, 60))
        flyer.paste(overlay, (0, 0), overlay)
    except: pass

    # 2. Logo Logic (Corrección: Recto arriba para LC)
    try:
        logo = Image.open(logo_path).convert("RGBA")
        if es_efe:
            diametro = 460
            c_x, c_y = ANCHO - diametro - 80, 40
            draw.ellipse([c_x, c_y, c_x + diametro, c_y + diametro], fill=BLANCO)
            logo = ImageOps.contain(logo, (int(diametro*0.85), int(diametro*0.85)))
            flyer.paste(logo, (c_x + (diametro - logo.width)//2, c_y + (diametro - logo.height)//2), logo)
        else:
            c_ancho, c_alto = 500, 380
            c_x = ANCHO - c_ancho - 80
            # Rectángulo recto arriba (y=0), redondeado solo abajo
            draw.rectangle([c_x, 0, c_x + c_ancho, 40], fill=BLANCO) # Tapa superior recta
            draw.rounded_rectangle([c_x, 0, c_x + c_ancho, c_alto], radius=50, fill=BLANCO)
            logo = ImageOps.contain(logo, (int(c_ancho*0.85), int(c_alto*0.80)))
            flyer.paste(logo, (c_x + (c_ancho - logo.width)//2, (c_alto - logo.height)//2 + 10), logo)
    except: pass

    # 3. Nombre Tienda (Corrección: EFE recto a la derecha)
    f_tienda = ImageFont.truetype(FONT_EXTRABOLD_COND, 90)
    txt_tienda = tienda_nombre.upper()
    tw_t = draw.textlength(txt_tienda, font=f_tienda)
    if es_efe:
        # Recto a la derecha (x=ANCHO)
        draw.rounded_rectangle([ANCHO - tw_t - 150, 620, ANCHO, 800], radius=50, fill=EFE_NARANJA)
        draw.rectangle([ANCHO - 60, 620, ANCHO, 800], fill=EFE_NARANJA) # Endereza borde derecho
        draw.text((ANCHO - tw_t - 80, 655), txt_tienda, font=f_tienda, fill=BLANCO)
    else:
        p_x = ANCHO - tw_t - 250
        draw.polygon([(p_x, 720), (p_x + 100, 520), (ANCHO, 520), (ANCHO, 720)], fill=NEGRO)
        draw.text((ANCHO - tw_t - 100, 570), txt_tienda, font=f_tienda, fill=LC_AMARILLO)

    # 4. Fecha Generación (Corrección: Recto a la izquierda)
    f_fecha = ImageFont.truetype(FONT_BOLD_COND, 45)
    txt_gen = f"Generado: {fecha_peru}"
    tw_g = draw.textlength(txt_gen, font=f_fecha)
    # x=0 para pegar al borde, redondeado solo derecha
    draw.rounded_rectangle([0, 850, tw_g + 80, 960], radius=40, fill=BLANCO)
    draw.rectangle([0, 850, 50, 960], fill=BLANCO) # Endereza borde izquierdo
    draw.text((40, 880), txt_gen, font=f_fecha, fill=NEGRO)

    # 5. Slogan
    f_slogan = ImageFont.truetype(FONT_EXTRABOLD, 105)
    slogan_txt = "¡APROVECHA ESTAS INCREÍBLES OFERTAS!"
    draw.rectangle([0, 1030, ANCHO, 1260], fill=color_slogan_bg)
    draw.text(((ANCHO-draw.textlength(slogan_txt, f_slogan))//2, 1085), slogan_txt, font=f_slogan, fill=BLANCO if es_efe else NEGRO)

    # 6. Productos
    anchos, altos = [110, 1300], [1350, 2150, 2950]
    f_marca = ImageFont.truetype(FONT_SEMIBOLD, 50)
    f_sku = ImageFont.truetype(FONT_BOLD_COND, 55)
    f_stock_label = ImageFont.truetype(FONT_BOLD_COND, 40)
    f_stock_val = ImageFont.truetype(FONT_EXTRABOLD, 100)

    for i, prod in enumerate(productos):
        if i >= 6: break
        x, y = anchos[i%2], altos[i//2]
        draw.rounded_rectangle([x, y, x+1090, y+760], radius=70, fill=BLANCO)
        
        img_p = descargar_imagen(prod.get('image_link'))
        if img_p:
            img_p.thumbnail((520, 520))
            flyer.paste(img_p, (x+30, y + (760-img_p.height)//2), img_p)

        tx, area_w = x + 570, 480
        draw.text((tx, y+50), str(prod['Marca']).upper(), font=f_marca, fill=GRIS_MARCA)
        
        # Articulo
        lines = textwrap.wrap(str(prod['Nombre Articulo']), width=18)
        ty = y + 110
        for line in lines[:3]:
            draw.text((tx, ty), line, font=ImageFont.truetype(FONT_REGULAR_COND, 60), fill=NEGRO)
            ty += 65

        # Stock Box
        ty_stock = y + 420
        draw.rounded_rectangle([tx, ty_stock, tx + 200, ty_stock + 180], radius=20, fill=color_slogan_bg)
        draw.text((tx + 45, ty_stock + 20), "STOCK", font=f_stock_label, fill=BLANCO if es_efe else NEGRO)
        stock_n = str(prod['Stock LM'])
        draw.text((tx + (200-draw.textlength(stock_n, f_stock_val))//2, ty_stock + 60), stock_n, font=f_stock_val, fill=BLANCO if es_efe else NEGRO)

        # SKU
        sku_val = str(prod['SKU'])
        draw.rounded_rectangle([tx, y + 630, tx + area_w, y + 710], radius=20, fill=NEGRO if not es_efe else EFE_NARANJA)
        draw.text((tx + (area_w - draw.textlength(sku_val, f_sku))//2, y + 645), sku_val, font=f_sku, fill=BLANCO)

    return flyer

# --- FLUJO ---
ss = conectar_sheets()
print(f"Filtrando datos para: {semana_actual}")

df_origen = pd.DataFrame(ss.worksheet("Origen Tdas").get_all_records())
df_lookup = pd.DataFrame(ss.worksheet("listado_productos").get_all_records())

df_origen['Stock LM'] = pd.to_numeric(df_origen['Stock LM'], errors='coerce').fillna(0)
df_filtered = df_origen[
    (df_origen['Semana'].astype(str) == semana_actual) & 
    (df_origen['Stock LM'] > 0)
].copy()

df_filtered['SKU_CLEAN'] = df_filtered['SKU'].astype(str).str.replace('-EX', '', case=False).str.strip()
img_map = df_lookup.set_index('sku')['base_image_path'].to_dict()
df_filtered['image_link'] = df_filtered['SKU_CLEAN'].map(img_map).fillna('')

ws_det = ss.worksheet("Detalle de Inventario")
ws_det.clear()
ws_det.update([df_filtered.columns.values.tolist()] + df_filtered.astype(str).values.tolist(), range_name='A1')

tienda_links = []
with ThreadPoolExecutor(max_workers=4) as exe:
    grupos = df_filtered.groupby('Tienda')
    futuros = {exe.submit(crear_flyer, g.to_dict('records'), str(n)): n for n, g in grupos if str(n).strip()}
    for f in futuros:
        tienda = futuros[f]
        try:
            img_f = f.result()
            t_clean = "".join(c for c in tienda if c.isalnum() or c in " -_").strip().replace(" ", "_")
            pdf_fn = f"LENTO_{t_clean}.pdf"
            img_f.convert("RGB").save(os.path.join(output_dir, pdf_fn))
            tienda_links.append([tienda, f"{URL_BASE_PAGES}view.html?file={urllib.parse.quote(pdf_fn)}"])
        except Exception as e: print(f"Error en {tienda}: {e}")

ws_flyer = ss.worksheet("FLYER_TIENDA")
ws_flyer.clear()
ws_flyer.update([["TIENDA RETAIL", "LINK PDF LENTO MOVIMIENTO"]] + tienda_links, range_name='A1')

print("¡Proceso finalizado!")