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
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# --- CONFIGURACIÓN VIP (Resolución Máxima) ---
ANCHO, ALTO = 2500, 3750
SHEET_ID = "1NQdhnPxgVe6N6LiVxh1ouzt5NHtqjR22EEqL6w1RpWQ"
DRIVE_FOLDER_ID = "1UzcwDFTNDE3961roIbzDgOi88Xi_n66Y"

output_dir = "temp_flyers"
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

def conectar_servicios():
    info_creds = json.loads(os.environ['GOOGLE_SHEETS_JSON'])
    creds = ServiceAccountCredentials.from_json_keyfile_dict(info_creds, ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"])
    client_sheets = gspread.authorize(creds).open_by_key(SHEET_ID)
    service_drive = build('drive', 'v3', credentials=creds)
    return client_sheets, service_drive

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
        img.thumbnail((650, 650))
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
            flyer.paste(ImageOps.fit(b.convert("RGBA"), (ANCHO, 1000)), (0, 0))
        
        logo_p = "logo-efe-sin-fondo.png" if es_efe else "logo-lc-sin-fondo.png"
        with Image.open(logo_p).convert("RGBA") as l:
            if es_efe:
                draw.ellipse([ANCHO-540, 40, ANCHO-80, 500], fill=BLANCO)
                logo = ImageOps.contain(l, (390, 390))
                flyer.paste(logo, (ANCHO-540+(460-logo.width)//2, 40+(460-logo.height)//2), logo)
            else:
                draw.rounded_rectangle([ANCHO-580, 0, ANCHO-80, 380], radius=50, fill=BLANCO)
                logo = ImageOps.contain(l, (425, 300))
                flyer.paste(logo, (ANCHO-580+(500-logo.width)//2, 50), logo)
    except: pass

    # Nombre Tienda con Diseño Recuperado
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

    # Fecha y Paginación
    f_fecha = ImageFont.truetype(FONT_BOLD_COND, 45)
    txt_gen = f"Generado: {fecha_peru} - PÁG {num_pag}"
    draw.rounded_rectangle([0, 850, 850, 960], radius=40, fill=BLANCO)
    draw.text((40, 880), txt_gen, font=f_fecha, fill=NEGRO)

    # Slogan
    draw.rectangle([0, 1030, ANCHO, 1260], fill=color_slogan_bg)
    draw.text((ANCHO//2, 1145), "¡APROVECHA ESTAS INCREÍBLES OFERTAS!", font=ImageFont.truetype(FONT_EXTRABOLD, 105), fill=BLANCO if es_efe else NEGRO, anchor="mm")

    # Grilla de Productos
    anchos, altos = [110, 1300], [1350, 2150, 2950]
    for i, prod in enumerate(productos):
        x, y = anchos[i%2], altos[i//2]
        draw.rounded_rectangle([x, y, x+1090, y+760], radius=70, fill=BLANCO)
        
        stock_raw = str(prod.get('Stock LM', '0')).replace(".", "").strip()
        if stock_raw in ["0", "", "nan"]: stock_txt, color_st = "SIN STOCK", GRIS_MARCA
        else: stock_txt, color_st = f"STOCK: {stock_raw}", (EFE_AZUL if es_efe else LC_AMARILLO)
        
        draw.rounded_rectangle([x+30, y+30, x+340, y+140], radius=15, fill=color_st)
        draw.text((x+185, y+85), stock_txt, font=ImageFont.truetype(FONT_BOLD_COND, 40), fill=BLANCO if es_efe else NEGRO, anchor="mm")

        path_img = cache_memoria.get(prod.get('image_link'))
        if path_img:
            with Image.open(path_img) as img:
                flyer.paste(img, (x+40, y+150), img)

        tx, area_w = x + 600, 450
        draw.text((tx, y+80), str(prod.get('Marca', '')).upper(), font=ImageFont.truetype(FONT_SEMIBOLD, 55), fill=GRIS_MARCA)
        lines = textwrap.wrap(str(prod.get('Nombre Articulo', '')), width=16)
        ty = y + 160
        for line in lines[:3]:
            draw.text((tx, ty), line, font=ImageFont.truetype(FONT_REGULAR_COND, 65), fill=NEGRO)
            ty += 75

        # Bloque Precio + SKU Unificado
        ty_p = y + 420
        p_raw = str(prod.get('Precio Vigente', '0')).replace(".00", "").replace(",00", "").strip()
        draw.rounded_rectangle([tx, ty_p, tx + area_w, ty_p + 180], radius=25, fill=color_slogan_bg)
        
        if p_raw in ["0.0", "0", "", "nan", "SIN PRECIO"]:
            draw.text((tx + area_w//2, ty_p + 90), "SIN PRECIO", font=ImageFont.truetype(FONT_EXTRABOLD, 80), fill=BLANCO if es_efe else NEGRO, anchor="mm")
        else:
            draw.text((tx + area_w//2, ty_p + 90), f"S/ {p_raw}", font=ImageFont.truetype(FONT_EXTRABOLD, 110), fill=BLANCO if es_efe else NEGRO, anchor="mm")

        sku_c = NEGRO if not es_efe else EFE_NARANJA
        draw.rounded_rectangle([tx, ty_p+180, tx+area_w, ty_p+280], radius=25, fill=sku_c)
        draw.rectangle([tx, ty_p+180, tx+area_w, ty_p+200], fill=sku_c) # Unificador
        draw.text((tx+area_w//2, ty_p+230), f"SKU: {prod['SKU']}", font=ImageFont.truetype(FONT_BOLD_COND, 55), fill=BLANCO, anchor="mm")

    return flyer

def gestionar_archivo_drive(service, file_path, file_name):
    media = MediaFileUpload(file_path, mimetype='application/pdf', resumable=True)
    
    # Buscamos si existe con soporte para todas las unidades
    query = f"name = '{file_name}' and '{DRIVE_FOLDER_ID}' in parents and trashed = false"
    results = service.files().list(
        q=query, 
        fields="files(id)",
        supportsAllDrives=True,
        includeItemsFromTrash=False
    ).execute()
    files = results.get('files', [])

    if files:
        file_id = files[0]['id']
        # Actualizar: Al actualizar un archivo que ya está en tu carpeta, usa TU cuota
        service.files().update(
            fileId=file_id, 
            media_body=media,
            supportsAllDrives=True
        ).execute()
    else:
        # Crear: Forzamos que se cree directamente bajo tu jerarquía
        file_metadata = {
            'name': file_name,
            'parents': [DRIVE_FOLDER_ID]
        }
        
        file = service.files().create(
            body=file_metadata, 
            media_body=media, 
            fields='id',
            supportsAllDrives=True # CLAVE: Ignora la cuota del bot y usa la de la carpeta
        ).execute()
        file_id = file.get('id')
        
    return f"https://drive.google.com/uc?export=download&id={file_id}"

def procesar_tienda_batch(data, service_drive):
    nombre, grupo = data
    try:
        prods = grupo.to_dict('records')
        paginas = []
        for i in range(0, len(prods), 6):
            paginas.append(crear_flyer(prods[i:i+6], str(nombre), (i//6)+1).convert("RGB"))
        
        if paginas:
            clean = "".join(c for c in str(nombre) if c.isalnum() or c in " _").strip().replace(" ", "_")
            fn = f"LENTO_{clean}.pdf"
            local_path = os.path.join(output_dir, fn)
            paginas[0].save(local_path, save_all=True, append_images=paginas[1:], quality=85, optimize=True)
            for p in paginas: p.close()
            
            # SUBIDA A DRIVE
            link_drive = gestionar_archivo_drive(service_drive, local_path, fn)
            return [nombre, link_drive]
    except Exception as e:
        print(f"Error en {nombre}: {e}")
    return None

# --- FLUJO ---
ss_client, drive_service = conectar_servicios()

# Carga de datos
df_raw = pd.DataFrame(ss_client.worksheet("Origen Tdas").get_all_records())
df_origen = pd.DataFrame({'Semana': df_raw.iloc[:, 1], 'Tienda': df_raw.iloc[:, 3], 'Marca': df_raw.iloc[:, 6], 'SKU': df_raw.iloc[:, 7], 'Nombre Articulo': df_raw.iloc[:, 8], 'Stock LM': df_raw.iloc[:, 11]})

df_lookup = pd.DataFrame(ss_client.worksheet("listado_productos").get_all_records())
img_dict = df_lookup.set_index('sku')['base_image_path'].to_dict()
df_origen['image_link'] = df_origen['SKU'].astype(str).str.replace('-EX', '', case=False).map(img_dict).fillna('')

promos = {}
for p in ["Promo01", "Promo03", "Promo04"]:
    df_p = pd.DataFrame(ss_client.worksheet(p).get_all_records())
    df_p.columns = df_p.columns.str.strip()
    df_p['K'] = df_p['Lista Precios'].astype(str).str.replace(".0","") + "_" + df_p['SKU'].astype(str)
    promos.update(df_p.set_index('K')['Precio Vigente'].to_dict())

df_txl = pd.DataFrame(ss_client.worksheet("TiendasxLista").get_all_records())
txl_map = {normalizar_nombre_tienda(r['TIENDA']): str(r['LISTA']).replace(".0","") for r in df_txl.to_dict('records') if 'TIENDA' in r}
df_origen['LISTA'] = df_origen['Tienda'].apply(normalizar_nombre_tienda).map(txl_map).fillna("")
df_origen['Precio Vigente'] = (df_origen['LISTA'] + "_" + df_origen['SKU'].astype(str)).map(promos).fillna("SIN PRECIO")

df_final = df_origen[df_origen['Semana'].astype(str) == semana_actual].copy()

# Guardar Detalle de Inventario
try:
    ws_det = ss_client.worksheet("Detalle de Inventario")
    ws_det.clear()
    cols = ['Semana', 'Tienda', 'Marca', 'SKU', 'Nombre Articulo', 'Stock LM', 'LISTA', 'Precio Vigente', 'image_link']
    ws_det.update([cols] + df_final[cols].astype(str).values.tolist(), range_name='A1')
except: pass

with ThreadPoolExecutor(max_workers=30) as exe:
    exe.map(descargar_y_cachear, df_final['image_link'].unique())

print(">> Generando y subiendo a Drive...")
tienda_links = []
for data in df_final.groupby('Tienda'):
    res = procesar_tienda_batch(data, drive_service)
    if res: 
        tienda_links.append(res)
        time.sleep(1) # Un segundo de respiro para que la API no se bloquee

ss_client.worksheet("FLYER_TIENDA").clear()
if tienda_links:
    ss_client.worksheet("FLYER_TIENDA").update([["TIENDA", "LINK DRIVE"]] + tienda_links, range_name='A1')

print(">> PROCESO COMPLETADO EXITOSAMENTE.")