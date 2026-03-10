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
URL_BASE_PAGES = f"https://{USUARIO_GITHUB}.github.io/{REPO_NOMBRE}/"

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

def formatear_precio(valor):
    try:
        if not valor or str(valor).strip() in ["", "-"]: return "0.00"
        s = str(valor).replace("S/.", "").replace("S/", "").replace(",", "").strip()
        f = float(s)
        return "{:,.2f}".format(f)
    except: return "0.00"

def normalizar_nombre_tienda(nombre):
    """ Estandariza nombres: 'Abancay - EFE' -> 'EFEABANCAY' """
    s = str(nombre).upper().replace(" ", "").replace("-", "")
    # Si el formato es CIUDADEFE, lo movemos a EFECIUDAD para el match
    if s.endswith("EFE"): s = "EFE" + s[:-3]
    if s.endswith("LC"): s = "LC" + s[:-2]
    return s

def crear_flyer(productos, tienda_nombre):
    es_efe = "EFE" in tienda_nombre.upper()
    color_fondo = EFE_AZUL_OSCURO if es_efe else LC_AMARILLO_OSCURO
    color_slogan_bg = EFE_AZUL if es_efe else LC_AMARILLO
    logo_path = "logo-efe-sin-fondo.png" if es_efe else "logo-lc-sin-fondo.png"
    tienda_bg_path = "efe tienda.jpg" if es_efe else "LC-MIRAFLORES-LOGO-3D[2].jpg"
    
    flyer = Image.new('RGB', (ANCHO, ALTO), color=color_fondo)
    draw = ImageDraw.Draw(flyer)
    
    try:
        bg = ImageOps.fit(Image.open(tienda_bg_path).convert("RGBA"), (ANCHO, 1000), method=Image.Resampling.LANCZOS)
        flyer.paste(bg, (0, 0))
        overlay = Image.new('RGBA', (ANCHO, 1000), (0, 0, 0, 60))
        flyer.paste(overlay, (0, 0), overlay)
    except: pass

    # Logo, Tienda, Fecha y Slogan (se mantienen igual)...
    # [Omitido por brevedad para enfocar en los cambios solicitados]
    try:
        logo = Image.open(logo_path).convert("RGBA")
        if es_efe:
            diametro = 460
            draw.ellipse([ANCHO-diametro-80, 40, ANCHO-80, 40+diametro], fill=BLANCO)
            logo = ImageOps.contain(logo, (int(diametro*0.85), int(diametro*0.85)))
            flyer.paste(logo, (ANCHO-diametro-80 + (diametro-logo.width)//2, 40 + (diametro-logo.height)//2), logo)
        else:
            c_w, c_h = 500, 380
            draw.rectangle([ANCHO-c_w-80, 0, ANCHO-80, 40], fill=BLANCO)
            draw.rounded_rectangle([ANCHO-c_w-80, 0, ANCHO-80, c_h], radius=50, fill=BLANCO)
            logo = ImageOps.contain(logo, (int(c_w*0.85), int(c_h*0.80)))
            flyer.paste(logo, (ANCHO-c_w-80 + (c_w-logo.width)//2, (c_h-logo.height)//2 + 10), logo)
    except: pass

    f_tienda = ImageFont.truetype(FONT_EXTRABOLD_COND, 90)
    txt_tienda = tienda_nombre.upper()
    tw_t = draw.textlength(txt_tienda, font=f_tienda)
    if es_efe:
        draw.rounded_rectangle([ANCHO-tw_t-150, 620, ANCHO, 800], radius=50, fill=EFE_NARANJA)
        draw.rectangle([ANCHO-60, 620, ANCHO, 800], fill=EFE_NARANJA)
        draw.text((ANCHO-tw_t-80, 655), txt_tienda, font=f_tienda, fill=BLANCO)
    else:
        draw.polygon([(ANCHO-tw_t-250, 720), (ANCHO-tw_t-150, 520), (ANCHO, 520), (ANCHO, 720)], fill=NEGRO)
        draw.text((ANCHO-tw_t-100, 570), txt_tienda, font=f_tienda, fill=LC_AMARILLO)

    f_fecha = ImageFont.truetype(FONT_BOLD_COND, 45)
    txt_gen = f"Generado: {fecha_peru}"
    tw_g = draw.textlength(txt_gen, font=f_fecha)
    draw.rounded_rectangle([0, 850, tw_g+80, 960], radius=40, fill=BLANCO)
    draw.rectangle([0, 850, 50, 960], fill=BLANCO)
    draw.text((40, 880), txt_gen, font=f_fecha, fill=NEGRO)

    f_slogan = ImageFont.truetype(FONT_EXTRABOLD, 105)
    slogan_txt = "¡APROVECHA ESTAS INCREÍBLES OFERTAS!"
    draw.rectangle([0, 1030, ANCHO, 1260], fill=color_slogan_bg)
    draw.text(((ANCHO-draw.textlength(slogan_txt, f_slogan))//2, 1085), slogan_txt, font=f_slogan, fill=BLANCO if es_efe else NEGRO)

    anchos, altos = [110, 1300], [1350, 2150, 2950]
    f_marca = ImageFont.truetype(FONT_SEMIBOLD, 50)
    f_sku = ImageFont.truetype(FONT_BOLD_COND, 55)
    f_simbolo = ImageFont.truetype(FONT_BOLD_COND, 65)
    
    for i, prod in enumerate(productos):
        if i >= 6: break
        x, y = anchos[i%2], altos[i//2]
        draw.rounded_rectangle([x, y, x+1090, y+760], radius=70, fill=BLANCO)
        
        # --- TAG STOCK CORREGIDO ---
        try:
            stock_clean = str(int(float(prod.get('Stock LM', 0))))
        except:
            stock_clean = str(prod.get('Stock LM', '0'))

        f_stock_label = ImageFont.truetype(FONT_BOLD_COND, 30)
        f_stock_num = ImageFont.truetype(FONT_EXTRABOLD, 50)
        
        tw_label = draw.textlength("STOCK", f_stock_label)
        tw_num = draw.textlength(stock_clean, f_stock_num)
        max_tw = max(tw_label, tw_num) + 40
        
        box_color = EFE_AZUL if es_efe else LC_AMARILLO
        txt_color = BLANCO if es_efe else NEGRO
        draw.rounded_rectangle([x+30, y+30, x+30+max_tw, y+140], radius=15, fill=box_color)
        # Palabra Arriba
        draw.text((x+30 + (max_tw-tw_label)//2, y+40), "STOCK", font=f_stock_label, fill=txt_color)
        # Numero Debajo
        draw.text((x+30 + (max_tw-tw_num)//2, y+75), stock_clean, font=f_stock_num, fill=txt_color)

        img_p = descargar_imagen(prod.get('image_link'))
        if img_p:
            img_p.thumbnail((500, 500))
            flyer.paste(img_p, (x+40, y + (760-img_p.height)//2 + 20), img_p)

        tx, area_w = x + 570, 480
        draw.text((tx, y+50), str(prod['Marca']).upper(), font=f_marca, fill=GRIS_MARCA)
        
        lines = textwrap.wrap(str(prod['Nombre Articulo']), width=18)
        ty = y + 110
        for line in lines[:3]:
            draw.text((tx, ty), line, font=ImageFont.truetype(FONT_REGULAR_COND, 60), fill=NEGRO)
            ty += 65

        # --- BLOQUES FUSIONADOS (PRECIO + SKU) ---
        ty_p = y + 420
        h_p = 180
        draw.rounded_rectangle([tx, ty_p, tx + area_w, ty_p + h_p], radius=25, fill=color_slogan_bg)
        draw.rectangle([tx, ty_p + h_p - 30, tx + area_w, ty_p + h_p], fill=color_slogan_bg) 
        
        precio_val = formatear_precio(prod.get('Precio Vigente', '0.00'))
        f_p_size = 110
        f_p = ImageFont.truetype(FONT_EXTRABOLD, f_p_size)
        while draw.textlength(f"S/ {precio_val}", f_p) > area_w - 40:
            f_p_size -= 5
            f_p = ImageFont.truetype(FONT_EXTRABOLD, f_p_size)
            
        full_p_w = draw.textlength("S/ ", f_simbolo) + draw.textlength(precio_val, f_p)
        start_px = tx + (area_w - full_p_w)//2
        draw.text((start_px, ty_p + 55), "S/ ", font=f_simbolo, fill=BLANCO if es_efe else NEGRO)
        draw.text((start_px + draw.textlength("S/ ", f_simbolo), ty_p + 35), precio_val, font=f_p, fill=BLANCO if es_efe else NEGRO)

        ty_sku = ty_p + h_p
        sku_color = NEGRO if not es_efe else EFE_NARANJA
        draw.rounded_rectangle([tx, ty_sku, tx + area_w, ty_sku + 85], radius=20, fill=sku_color)
        draw.rectangle([tx, ty_sku, tx + area_w, ty_sku + 30], fill=sku_color) 
        sku_txt = str(prod['SKU'])
        draw.text((tx + (area_w - draw.textlength(sku_txt, f_sku))//2, ty_sku + 15), sku_txt, font=f_sku, fill=BLANCO)

    return flyer

# --- FLUJO DE DATOS (TABLA MAESTRA) ---
ss = conectar_sheets()
print("Generando Tabla Maestra...")

# 1. Origen Tdas
df_raw = pd.DataFrame(ss.worksheet("Origen Tdas").get_all_records())
df_origen = pd.DataFrame({
    'Semana': df_raw.iloc[:, 1], 'Tienda': df_raw.iloc[:, 3],
    'Marca': df_raw.iloc[:, 6], 'SKU': df_raw.iloc[:, 7],
    'Nombre Articulo': df_raw.iloc[:, 8], 'Stock LM': df_raw.iloc[:, 11]
})
df_origen['TIENDA_KEY'] = df_origen['Tienda'].apply(normalizar_nombre_tienda)

# 2. TiendasxLista (Limpieza de .0)
df_txl = pd.DataFrame(ss.worksheet("TiendasxLista").get_all_records())
df_txl.columns = df_txl.columns.str.strip().str.upper()
df_txl['LISTA'] = pd.to_numeric(df_txl['LISTA'], errors='coerce').fillna(0).astype(int).astype(str)
df_txl['TIENDA_KEY'] = df_txl['TIENDA'].apply(normalizar_nombre_tienda)

# Merge 1: Obtener LISTA
df_origen = df_origen.merge(df_txl[['TIENDA_KEY', 'LISTA']], on='TIENDA_KEY', how='left')

# 3. Promos
promos = []
for p_sheet in ["Promo01", "Promo03", "Promo04"]:
    temp = pd.DataFrame(ss.worksheet(p_sheet).get_all_records())
    temp.columns = temp.columns.str.strip()
    # Estandarizar Lista Precios a texto (ej: "4")
    temp['Lista Precios'] = pd.to_numeric(temp['Lista Precios'], errors='coerce').fillna(0).astype(int).astype(str)
    promos.append(temp[['Lista Precios', 'SKU', 'Precio Vigente']])
df_promos = pd.concat(promos).drop_duplicates(subset=['Lista Precios', 'SKU'])

# Merge 2: Obtener Precio usando LISTA limpia
df_master = df_origen.merge(df_promos, left_on=['LISTA', 'SKU'], right_on=['Lista Precios', 'SKU'], how='left')

# 4. Imágenes y Limpieza
df_lookup = pd.DataFrame(ss.worksheet("listado_productos").get_all_records())
df_master['SKU_CLEAN'] = df_master['SKU'].astype(str).str.replace('-EX', '', case=False).str.strip()
img_map = df_lookup.set_index('sku')['base_image_path'].to_dict()
df_master['image_link'] = df_master['SKU_CLEAN'].map(img_map).fillna('')

df_master['Stock LM_NUM'] = pd.to_numeric(df_master['Stock LM'], errors='coerce').fillna(0)
df_final = df_master[
    (df_master['Semana'].astype(str) == semana_actual) & (df_master['Stock LM_NUM'] > 0)
].copy()

# Guardar Tabla Maestra
df_final = df_final.fillna("")
ws_det = ss.worksheet("Detalle de Inventario")
ws_det.clear()
col_order = ['Semana', 'Tienda', 'Marca', 'SKU', 'Nombre Articulo', 'Stock LM', 'LISTA', 'Precio Vigente', 'SKU_CLEAN', 'image_link']
ws_det.update([col_order] + df_final[col_order].astype(str).values.tolist(), range_name='A1')

# --- GENERACIÓN DE PDFS ---
tienda_links = []
with ThreadPoolExecutor(max_workers=4) as exe:
    grupos = df_final.groupby('Tienda')
    futuros = {exe.submit(crear_flyer, g.to_dict('records'), str(n)): n for n, g in grupos if str(n).strip()}
    for f in futuros:
        tienda = futuros[f]
        try:
            pdf_name = f"LENTO_{tienda.replace(' ', '_')}.pdf"
            f.result().convert("RGB").save(os.path.join(output_dir, pdf_name))
            tienda_links.append([tienda, f"{URL_BASE_PAGES}view.html?file={urllib.parse.quote(pdf_name)}"])
        except Exception as e: print(f"Error en {tienda}: {e}")

ss.worksheet("FLYER_TIENDA").clear()
ss.worksheet("FLYER_TIENDA").update([["TIENDA RETAIL", "LINK PDF LENTO MOVIMIENTO"]] + tienda_links, range_name='A1')
print("¡Proceso finalizado!")