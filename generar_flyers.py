import pandas as pd
import requests
from PIL import Image, ImageDraw, ImageFont, ImageOps
from io import BytesIO
import gc
import os
import gspread
import json
import textwrap
import time
import hashlib
from datetime import datetime, timedelta
from oauth2client.service_account import ServiceAccountCredentials
from concurrent.futures import ThreadPoolExecutor
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2.credentials import Credentials

# --- CONFIGURACIÓN VIP ---
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

def conectar_servicios():
    info_creds = json.loads(os.environ['GOOGLE_SHEETS_JSON'])
    creds_sheets = ServiceAccountCredentials.from_json_keyfile_dict(info_creds, ["https://spreadsheets.google.com/feeds"])
    client_sheets = gspread.authorize(creds_sheets).open_by_key(SHEET_ID)
    info_token = json.loads(os.environ['GOOGLE_TOKEN_JSON'])
    creds_drive = Credentials.from_authorized_user_info(info_token)
    service_drive = build('drive', 'v3', credentials=creds_drive)
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
        img.thumbnail((500, 500)) 
        img.save(fname, "PNG")
        cache_memoria[url] = fname
    except: cache_memoria[url] = None

def draw_custom_rounded(draw, xy, radius, fill, corners=(True, True, True, True)):
    """Dibuja un rectángulo con esquinas redondeadas selectivas sin errores de píxeles"""
    x0, y0, x1, y1 = xy
    # Cuerpo central
    draw.rectangle([x0 + radius, y0, x1 - radius, y1], fill=fill)
    draw.rectangle([x0, y0 + radius, x1, y1 - radius], fill=fill)
    
    # Esquinas: (TL, TR, BR, BL)
    if corners[0]: draw.pieslice([x0, y0, x0 + radius * 2, y0 + radius * 2], 180, 270, fill=fill)
    else: draw.rectangle([x0, y0, x0 + radius, y0 + radius], fill=fill)
    
    if corners[1]: draw.pieslice([x1 - radius * 2, y0, x1, y0 + radius * 2], 270, 360, fill=fill)
    else: draw.rectangle([x1 - radius, y0, x1, y0 + radius], fill=fill)
    
    if corners[2]: draw.pieslice([x1 - radius * 2, y1 - radius * 2, x1, y1], 0, 90, fill=fill)
    else: draw.rectangle([x1 - radius, y1 - radius, x1, y1], fill=fill)
    
    if corners[3]: draw.pieslice([x0, y1 - radius * 2, x0 + radius * 2, y1], 90, 180, fill=fill)
    else: draw.rectangle([x0, y1 - radius, x0 + radius, y1], fill=fill)

def limpiar_valor_puro(valor, es_precio=True):
    s = str(valor).strip().replace(" ", "")
    if s in ["0", "0.0", "", "nan", "-", "SIN PRECIO", "SINPRECIO"]:
        return "-"
    if s.endswith(".0"): s = s[:-2]
    return s

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
                # Logo LC: Solo redondeado abajo
                draw_custom_rounded(draw, [ANCHO-580, 0, ANCHO-80, 380], 50, BLANCO, (False, False, True, True))
                logo = ImageOps.contain(l, (425, 300))
                flyer.paste(logo, (ANCHO-580+(500-logo.width)//2, 50), logo)
    except: pass

    f_tienda = ImageFont.truetype(FONT_EXTRABOLD_COND, 90)
    txt_tienda = tienda_nombre.upper()
    tw_t = draw.textlength(txt_tienda, font=f_tienda)
    if es_efe:
        draw_custom_rounded(draw, [ANCHO - tw_t - 150, 620, ANCHO, 800], 50, EFE_NARANJA, (True, False, False, True))
        draw.text((ANCHO - tw_t - 80, 655), txt_tienda, font=f_tienda, fill=BLANCO)
    else:
        p_x = ANCHO - tw_t - 250
        draw.polygon([(p_x, 720), (p_x + 100, 520), (ANCHO, 520), (ANCHO, 720)], fill=NEGRO)
        draw.text((ANCHO - tw_t - 100, 570), txt_tienda, font=f_tienda, fill=LC_AMARILLO)

    f_fecha = ImageFont.truetype(FONT_BOLD_COND, 45)
    txt_gen = f"Generado: {fecha_peru} - PÁG {num_pag}"
    draw_custom_rounded(draw, [0, 850, 850, 960], 40, BLANCO, (False, True, True, False))
    draw.text((40, 880), txt_gen, font=f_fecha, fill=NEGRO)

    draw.rectangle([0, 1030, ANCHO, 1260], fill=color_slogan_bg)
    draw.text((ANCHO//2, 1145), "¡APROVECHA ESTAS INCREÍBLES OFERTAS!", font=ImageFont.truetype(FONT_EXTRABOLD, 105), fill=BLANCO if es_efe else NEGRO, anchor="mm")

    anchos, altos = [110, 1300], [1350, 2150, 2950]
    for i, prod in enumerate(productos):
        x, y = anchos[i%2], altos[i//2]
        draw.rounded_rectangle([x, y, x+1090, y+760], radius=70, fill=BLANCO)
        
        # --- STOCK ---
        stock_txt = limpiar_valor_puro(prod.get('Stock LM', '0'), es_precio=False)
        color_st = GRIS_MARCA if stock_txt == "-" else (EFE_AZUL if es_efe else LC_AMARILLO)
        draw.rounded_rectangle([x+30, y+30, x+300, y+160], radius=20, fill=color_st)
        draw.text((x+165, y+65), "STOCK", font=ImageFont.truetype(FONT_BOLD_COND, 35), fill=BLANCO if es_efe else NEGRO, anchor="mm")
        draw.text((x+165, y+115), stock_txt, font=ImageFont.truetype(FONT_BOLD_COND, 55), fill=BLANCO if es_efe else NEGRO, anchor="mm")

        path_img = cache_memoria.get(prod.get('image_link'))
        if path_img:
            with Image.open(path_img) as img: flyer.paste(img, (x+50, y+200), img)

        tx, area_w = x + 600, 450
        draw.text((tx, y+80), str(prod.get('Marca', '')).upper(), font=ImageFont.truetype(FONT_SEMIBOLD, 55), fill=GRIS_MARCA)
        lines = textwrap.wrap(str(prod.get('Nombre Articulo', '')), width=16)
        ty = y + 160
        for line in lines[:3]:
            draw.text((tx, ty), line, font=ImageFont.truetype(FONT_REGULAR_COND, 65), fill=NEGRO)
            ty += 75

        # --- PRECIO ---
        p_final = limpiar_valor_puro(prod.get('Precio Vigente', '0'), es_precio=True)
        ty_p = y + 420
        color_p_bg = color_slogan_bg if p_final != "-" else GRIS_MARCA
        # Precio: Solo arriba redondeado
        draw_custom_rounded(draw, [tx, ty_p, tx + area_w, ty_p + 180], 25, color_p_bg, (True, True, False, False))
        
        if p_final == "-":
            draw.text((tx + area_w//2, ty_p + 90), "-", font=ImageFont.truetype(FONT_EXTRABOLD, 110), fill=BLANCO if es_efe else NEGRO, anchor="mm")
        else:
            f_sol, f_num = ImageFont.truetype(FONT_EXTRABOLD, 60), ImageFont.truetype(FONT_EXTRABOLD, 110)
            t_sol = "S/ "
            w_total = draw.textlength(t_sol, font=f_sol) + draw.textlength(p_final, font=f_num)
            curr_x = tx + (area_w - w_total) // 2
            draw.text((curr_x, ty_p + 105), t_sol, font=f_sol, fill=BLANCO if es_efe else NEGRO, anchor="ls")
            draw.text((curr_x + draw.textlength(t_sol, font=f_sol), ty_p + 115), p_final, font=f_num, fill=BLANCO if es_efe else NEGRO, anchor="ls")

        # --- SKU (CORREGIDO ESQUINAS Y CUADRADITO) ---
        sku_val = str(prod['SKU'])
        sku_c = NEGRO if not es_efe else EFE_NARANJA
        # SKU: Solo abajo redondeado (BR, BL). TL y TR quedan rectos (False, False)
        draw_custom_rounded(draw, [tx, ty_p+180, tx+area_w, ty_p+280], 25, sku_c, (False, False, True, True))
        draw.text((tx+area_w//2, ty_p+230), sku_val, font=ImageFont.truetype(FONT_BOLD_COND, 55), fill=BLANCO, anchor="mm")

    return flyer

def gestionar_archivo_drive(service, file_path, file_name):
    media = MediaFileUpload(file_path, mimetype='application/pdf', resumable=True)
    query = f"name = '{file_name}' and '{DRIVE_FOLDER_ID}' in parents and trashed = false"
    results = service.files().list(q=query, fields="files(id)").execute()
    files = results.get('files', [])
    if files:
        file_id = files[0]['id']
        service.files().update(fileId=file_id, media_body=media).execute()
    else:
        file_metadata = {'name': file_name, 'parents': [DRIVE_FOLDER_ID]}
        file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        file_id = file.get('id')
    service.permissions().create(fileId=file_id, body={'type': 'anyone', 'role': 'reader'}).execute()
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
            
            # 1. Cerramos las imágenes de las páginas
            for p in paginas: p.close()
            
            link_drive = gestionar_archivo_drive(service_drive, local_path, fn)
            
            # 2. EL CAMBIO CLAVE: Borramos la lista y liberamos RAM
            del paginas
            import gc
            gc.collect() 
            
            return [nombre, link_drive]
    except Exception as e: print(f"Error en {nombre}: {e}")
    return None

ss_client, drive_service = conectar_servicios()
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

with ThreadPoolExecutor(max_workers=20) as exe:
    exe.map(descargar_y_cachear, df_final['image_link'].unique())

tienda_links = []
for data in df_final.groupby('Tienda'):
    res = procesar_tienda_batch(data, drive_service)
    if res:
        tienda_links.append(res)
        time.sleep(0.5)

ss_client.worksheet("FLYER_TIENDA").clear()
if tienda_links:
    ss_client.worksheet("FLYER_TIENDA").update([["TIENDA", "LINK DRIVE"]] + tienda_links, range_name='A1')
print(">> PROCESO COMPLETADO EXITOSAMENTE.")