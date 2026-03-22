import sys
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

ahora_global = datetime.utcnow() - timedelta(hours=5)
semana_actual = f"Sem{ahora_global.isocalendar()[1]}"

cache_memoria = {}

# --- PRE-CARGA DE FUENTES ---
FONT_BOLD_COND = "Mark Simonson - Proxima Nova Alt Condensed Bold.otf"
FONT_EXTRABOLD_COND = "Mark Simonson - Proxima Nova Alt Condensed Extrabold.otf"
FONT_REGULAR_COND = "Mark Simonson - Proxima Nova Alt Condensed Regular.otf"
FONT_EXTRABOLD = "Mark Simonson - Proxima Nova Extrabold.otf"
FONT_SEMIBOLD = "Mark Simonson - Proxima Nova Semibold.otf"

f_tienda = ImageFont.truetype(FONT_EXTRABOLD_COND, 90)
f_fecha = ImageFont.truetype(FONT_BOLD_COND, 45)
f_slogan = ImageFont.truetype(FONT_EXTRABOLD, 105)
f_marca = ImageFont.truetype(FONT_SEMIBOLD, 55)
f_nombre = ImageFont.truetype(FONT_REGULAR_COND, 65)
f_soles = ImageFont.truetype(FONT_EXTRABOLD, 60)
f_precio = ImageFont.truetype(FONT_EXTRABOLD, 110)
f_sku = ImageFont.truetype(FONT_BOLD_COND, 55)
f_stock_tag = ImageFont.truetype(FONT_BOLD_COND, 35)
f_stock_val = ImageFont.truetype(FONT_BOLD_COND, 55)

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

def lectura_segura(client, nombre_hoja):
    for intento in range(5):
        try:
            return pd.DataFrame(client.worksheet(nombre_hoja).get_all_records())
        except Exception as e:
            if "429" in str(e):
                espera = 60 * (intento + 1) 
                print(f"!!! CUOTA EXCEDIDA en {nombre_hoja}. Esperando {espera}s...")
                time.sleep(espera)
            else: raise e
    raise Exception(f"Fallo lectura tras 5 intentos en {nombre_hoja}")

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
        with Image.open(BytesIO(res.content)) as img:
            img = img.convert("RGBA")
            img.thumbnail((550, 550))
            img.save(fname, "PNG")
        cache_memoria[url] = fname
    except: cache_memoria[url] = None

def draw_custom_rounded(draw, xy, radius, fill, corners=(True, True, True, True)):
    x0, y0, x1, y1 = xy
    draw.rectangle([x0 + radius, y0, x1 - radius, y1], fill=fill)
    draw.rectangle([x0, y0 + radius, x1, y1 - radius], fill=fill)
    # Solo dibuja el redondeado si la esquina NO está pegada al borde
    if corners[0]: draw.pieslice([x0, y0, x0 + radius * 2, y0 + radius * 2], 180, 270, fill=fill)
    else: draw.rectangle([x0, y0, x0 + radius, y0 + radius], fill=fill) # Esquina recta
    
    if corners[1]: draw.pieslice([x1 - radius * 2, y0, x1, y0 + radius * 2], 270, 360, fill=fill)
    else: draw.rectangle([x1 - radius, y0, x1, y0 + radius], fill=fill)
    
    if corners[2]: draw.pieslice([x1 - radius * 2, y1 - radius * 2, x1, y1], 0, 90, fill=fill)
    else: draw.rectangle([x1 - radius, y1 - radius, x1, y1], fill=fill)
    
    if corners[3]: draw.pieslice([x0, y1 - radius * 2, x0 + radius * 2, y1], 90, 180, fill=fill)
    else: draw.rectangle([x0, y1 - radius, x0 + radius, y1], fill=fill)

def limpiar_valor_puro(valor, es_precio=True):
    s = str(valor).strip().replace(" ", "")
    if s in ["0", "0.0", "", "nan", "-", "SIN PRECIO"]: return "-"
    if s.endswith(".0"): s = s[:-2]
    return s

def crear_flyer(productos, tienda_nombre, num_pag):
    ahora = datetime.utcnow() - timedelta(hours=5)
    fecha_hoy = ahora.strftime("%d/%m/%Y %I:%M %p")
    
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
                draw_custom_rounded(draw, [ANCHO-580, 0, ANCHO-80, 380], 50, BLANCO, (False, False, True, True))
                logo = ImageOps.contain(l, (425, 300))
                flyer.paste(logo, (ANCHO-580+(500-logo.width)//2, 50), logo)
    except: pass

    txt_tienda = tienda_nombre.upper()
    tw_t = draw.textlength(txt_tienda, font=f_tienda)
    if es_efe:
        draw_custom_rounded(draw, [ANCHO - tw_t - 150, 620, ANCHO, 800], 50, EFE_NARANJA, (True, False, False, True))
        draw.text((ANCHO - tw_t - 80, 655), txt_tienda, font=f_tienda, fill=BLANCO)
    else:
        p_x = ANCHO - tw_t - 250
        draw.polygon([(p_x, 720), (p_x + 100, 520), (ANCHO, 520), (ANCHO, 720)], fill=NEGRO)
        draw.text((ANCHO - tw_t - 100, 570), txt_tienda, font=f_tienda, fill=LC_AMARILLO)

    txt_gen = f"Generado: {fecha_hoy} - PÁG {num_pag}"
    draw_custom_rounded(draw, [0, 850, 850, 960], 40, BLANCO, (False, True, True, False))
    draw.text((40, 880), txt_gen, font=f_fecha, fill=NEGRO)
    
    draw.rectangle([0, 1030, ANCHO, 1260], fill=color_slogan_bg)
    draw.text((ANCHO//2, 1145), "¡APROVECHA ESTAS INCREÍBLES OFERTAS!", font=f_slogan, fill=BLANCO if es_efe else NEGRO, anchor="mm")

    anchos, altos = [110, 1300], [1350, 2150, 2950]
    for i, prod in enumerate(productos):
        x, y = anchos[i%2], altos[i//2]
        draw.rounded_rectangle([x, y, x+1090, y+760], radius=70, fill=BLANCO)
        
        path_img = cache_memoria.get(prod.get('image_link'))
        if path_img and os.path.exists(path_img):
            with Image.open(path_img) as img:
                flyer.paste(img.convert("RGBA"), (x+50, y+200), img.convert("RGBA"))

        stock_txt = limpiar_valor_puro(prod.get('Stock LM', '0'), es_precio=False)
        color_st = GRIS_MARCA if stock_txt == "-" else (EFE_AZUL if es_efe else LC_AMARILLO)
        draw.rounded_rectangle([x+30, y+30, x+300, y+160], radius=20, fill=color_st)
        draw.text((x+165, y+65), "STOCK", font=f_stock_tag, fill=BLANCO if es_efe else NEGRO, anchor="mm")
        draw.text((x+165, y+115), stock_txt, font=f_stock_val, fill=BLANCO if es_efe else NEGRO, anchor="mm")

        tx, area_w = x + 600, 450
        draw.text((tx, y+80), str(prod.get('Marca', '')).upper(), font=f_marca, fill=GRIS_MARCA)
        lines = textwrap.wrap(str(prod.get('Nombre Articulo', '')), width=16)
        ty = y + 160
        for line in lines[:3]:
            draw.text((tx, ty), line, font=f_nombre, fill=NEGRO)
            ty += 75

        # --- AJUSTE MAESTRO DE PRECIO Y COLOR ---
        p_raw = str(prod.get('Precio Vigente', '0'))
        p_final = limpiar_valor_puro(p_raw, es_precio=True)
        ty_p = y + 420
        
        # Validamos si no hay precio de forma estricta
        es_vacio = p_final == "-" or "SIN PRECIO" in p_raw.upper()
        
        if es_vacio:
            p_final = "-"
            color_p_bg = GRIS_MARCA
        else:
            color_p_bg = color_slogan_bg
            
        # Dibujamos contenedor
        draw_custom_rounded(draw, [tx, ty_p, tx + area_w, ty_p + 180], 25, color_p_bg, (True, True, False, False))
        
        if es_vacio:
            # Dibujamos SOLO el guion centrado, sin "S/ "
            draw.text((tx + area_w//2, ty_p + 90), "-", font=f_precio, fill=BLANCO if es_efe else NEGRO, anchor="mm")
        else:
            # Dibujamos precio normal con S/ 
            t_sol = "S/ "
            w_total = draw.textlength(t_sol, font=f_soles) + draw.textlength(p_final, font=f_precio)
            curr_x = tx + (area_w - w_total) // 2
            draw.text((curr_x, ty_p + 105), t_sol, font=f_soles, fill=BLANCO if es_efe else NEGRO, anchor="ls")
            draw.text((curr_x + draw.textlength(t_sol, font=f_soles), ty_p + 115), p_final, font=f_precio, fill=BLANCO if es_efe else NEGRO, anchor="ls")
        
        sku_val = str(prod['SKU'])
        sku_c = NEGRO if not es_efe else EFE_NARANJA
        draw_custom_rounded(draw, [tx, ty_p+180, tx+area_w, ty_p+280], 25, sku_c, (False, False, True, True))
        draw.text((tx+area_w//2, ty_p+230), sku_val, font=f_sku, fill=BLANCO, anchor="mm")
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
        clean = "".join(c for c in str(nombre) if c.isalnum() or c in " _").strip().replace(" ", "_")
        fn = f"LENTO_{clean}.pdf"
        local_path = os.path.join(output_dir, fn)
        
        paginas = []
        
        for i in range(0, len(prods), 6):
            pag_actual = (i//6) + 1
            img_flyer = crear_flyer(prods[i:i+6], str(nombre), pag_actual).convert("RGB")
            paginas.append(img_flyer)
            
            # Limpieza frecuente de RAM (Cada 15 páginas)
            if pag_actual % 15 == 0:
                gc.collect()

        if paginas:
            num_pags = len(paginas)
            
            # --- AJUSTE SOLICITADO: CALIDAD DINÁMICA ---
            # Si el PDF tiene más de 50 páginas (como la de 350), usamos 33.
            # Si tiene pocas, usamos 35 (que es tu estándar actual).
            calidad_final = 33 if num_pags > 50 else 35
            
            # Subsampling 2 ahorra mucho espacio en archivos grandes sin perder nitidez de texto.
            muestreo = 2 if num_pags > 50 else 1
            
            print(f">> Guardando {num_pags} págs | Calidad: {calidad_final} | Tienda: {nombre}")

            paginas[0].save(
                local_path, 
                save_all=True, 
                append_images=paginas[1:], 
                quality=calidad_final, 
                optimize=True, 
                progressive=True, 
                subsampling=muestreo
            )
            
            # --- LIBERACIÓN DE RAM ABSOLUTA (Crucial para el bloque 240) ---
            for p in paginas:
                p.close()
            
            del paginas
            gc.collect()
            
            link = gestionar_archivo_drive(service_drive, local_path, fn)
            
            if os.path.exists(local_path): 
                os.remove(local_path)
                
            return [nombre, link]
            
    except Exception as e: 
        print(f"!!! ERROR CRÍTICO EN {nombre}: {e}")
        if 'paginas' in locals():
            for p in paginas: p.close()
            del paginas
            gc.collect()
            
    return None

# --- FLUJO PRINCIPAL ---
ss_client, drive_service = conectar_servicios()

# [!] REGLA DE ORO: Escalonamiento inicial ANTES de cargar datos pesados
inicio = int(sys.argv[1]) if len(sys.argv) > 1 else 0
fin = int(sys.argv[2]) if len(sys.argv) > 2 else 999
time.sleep((inicio // 25) * 45) # Escalonamiento para no saturar Google

df_raw = lectura_segura(ss_client, "Origen Tdas")
df_origen = pd.DataFrame({'Semana': df_raw.iloc[:, 1], 'Tienda': df_raw.iloc[:, 3], 'Marca': df_raw.iloc[:, 6], 'SKU': df_raw.iloc[:, 7], 'Nombre Articulo': df_raw.iloc[:, 8], 'Stock LM': df_raw.iloc[:, 11]})
df_lookup = lectura_segura(ss_client, "listado_productos")
img_dict = df_lookup.set_index('sku')['base_image_path'].to_dict()
df_origen['image_link'] = df_origen['SKU'].astype(str).str.replace('-EX', '', case=False).map(img_dict).fillna('')

promos = {}
for p in ["Promo01", "Promo03", "Promo04"]:
    df_p = lectura_segura(ss_client, p)
    df_p['K'] = df_p['Lista Precios'].astype(str).str.replace(".0","") + "_" + df_p['SKU'].astype(str)
    promos.update(df_p.set_index('K')['Precio Vigente'].to_dict())

df_txl = lectura_segura(ss_client, "TiendasxLista")
txl_map = {normalizar_nombre_tienda(r['TIENDA']): str(r['LISTA']).replace(".0","") for r in df_txl.to_dict('records') if 'TIENDA' in r}
df_origen['LISTA'] = df_origen['Tienda'].apply(normalizar_nombre_tienda).map(txl_map).fillna("")
df_origen['Precio Vigente'] = (df_origen['LISTA'] + "_" + df_origen['SKU'].astype(str)).map(promos).fillna("SIN PRECIO")
df_final = df_origen[df_origen['Semana'].astype(str) == semana_actual].copy()

# [!] ACTUALIZACIÓN DE HOJA MAESTRA (Solo la primera máquina)
if inicio == 0:
    try:
        print(">> Actualizando hoja maestra 'Detalle de Inventario'...")
        ws_inv = ss_client.worksheet("Detalle de Inventario")
        ws_inv.clear()
        columnas = ['Semana', 'Tienda', 'Marca', 'SKU', 'Nombre Articulo', 'Stock LM', 'LISTA', 'Precio Vigente', 'image_link']
        data_inv = [columnas] + df_final[columnas].fillna("-").values.tolist()
        ws_inv.update('A1', data_inv)
        # Limpiar también la hoja de links
        ss_client.worksheet("FLYER_TIENDA").clear()
        ss_client.worksheet("FLYER_TIENDA").update('A1', [["TIENDA", "LINK DRIVE"]])
    except Exception as e: print(f"Error actualizando maestra: {e}")

tiendas_procesadas = list(df_final.groupby('Tienda'))
total_reales = len(tiendas_procesadas)

if len(sys.argv) > 2:
    inicio, fin = int(sys.argv[1]), int(sys.argv[2])
    # Validación Senior: Si el inicio es mayor al total, cerramos con éxito
    if inicio >= total_reales:
        print(f">>> [!] AVISO: El inicio {inicio} supera el total de tiendas ({total_reales}). Nada que hacer.")
        sys.exit(0)
    
    fin = min(fin, total_reales)
    tiendas_a_procesar = tiendas_procesadas[inicio:fin]
else:
    tiendas_a_procesar = tiendas_procesadas

# [!] SEGURO: Si no hay tiendas en este rango específico, salimos limpiamente
if not tiendas_a_procesar:
    print(">>> [!] Rango vacío para esta máquina. Finalizando...")
    sys.exit(0)

# Descarga previa (Blindada contra lista vacía)
try:
    urls = pd.concat([g for n, g in tiendas_a_procesar])['image_link'].unique()
    with ThreadPoolExecutor(max_workers=5) as exe:
        exe.map(descargar_y_cachear, urls)
except Exception as e:
    print(f">>> Error preparando imágenes: {e}")

tienda_links = []
batch_size = 2 

for idx, data in enumerate(tiendas_a_procesar):
    # Reintento de conexión por si el token parpadea (invalid_grant)
    for intento_auth in range(2):
        try:
            res = procesar_tienda_batch(data, drive_service)
            if res:
                tienda_links.append(res)
                if len(tienda_links) % batch_size == 0:
                    ss_client.worksheet("FLYER_TIENDA").append_rows(tienda_links[-batch_size:])
            break # Éxito en la tienda
        except Exception as e:
            if "invalid_grant" in str(e) and intento_auth == 0:
                print("!!! Re-conectando servicios por token expirado...")
                ss_client, drive_service = conectar_servicios() # Intenta re-conectar
                continue
            print(f"Error crítico en tienda {data[0]}: {e}")
            break
    gc.collect() 

# --- SECCIÓN FINAL CORREGIDA PARA ESCRITURA TOTAL ---
if tienda_links:
    try:
        ws_output = ss_client.worksheet("FLYER_TIENDA")
        # Calculamos cuántos faltan por escribir que no entraron en los lotes de 2
        ya_escritos = (len(tienda_links) // batch_size) * batch_size
        pendientes = tienda_links[ya_escritos:]
        
        if pendientes:
            print(f">> Escribiendo últimos {len(pendientes)} links pendientes...")
            ws_output.append_rows(pendientes)
    except Exception as e: 
        print(f"!!! Error en escritura final: {e}")

print(">> PROCESO COMPLETADO CON ÉXITO.")