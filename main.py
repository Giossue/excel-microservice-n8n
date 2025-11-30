from fastapi import FastAPI, Response
from pydantic import BaseModel
from typing import List, Optional
import xlsxwriter
import io
import asyncio
import aiohttp
import requests
from PIL import Image as PILImage

app = FastAPI()

# --- MODELOS DE DATOS ---
class Product(BaseModel):
    image_product: Optional[str] = None
    id_product: str
    product_description: str
    quantity: int
    unit_price: float
    subtotal: float

class QuotationData(BaseModel):
    items: List[Product]
    Total: float

# --- CONFIGURACIÓN ---
sem = asyncio.Semaphore(5)

# Celdas y Límites
CELL_WIDTH_PX = 210
CELL_HEIGHT_PX = 160

# Límites de Imágenes (Con margen de seguridad)
MAX_PROD_W = 180
MAX_PROD_H = 130
MAX_LOGO_W = 180
MAX_LOGO_H = 100 # Un poco más alto para que quepa el texto del logo

async def process_image(session, url):
    if not url or not str(url).startswith("http"):
        return None, 0, 0
    
    async with sem: 
        try:
            headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}
            async with session.get(str(url), headers=headers, timeout=15) as response:
                if response.status == 200:
                    data = await response.read()
                    # Usamos la función de redimensionado inteligente
                    return await asyncio.to_thread(smart_resize, data, MAX_PROD_W, MAX_PROD_H)
        except Exception as e:
            print(f"Error descargando {url}: {e}")
            return None, 0, 0
    return None, 0, 0

def smart_resize(data, target_w, target_h):
    try:
        with PILImage.open(io.BytesIO(data)) as img:
            # 1. GESTIÓN DE TRANSPARENCIA (Vital para el logo)
            if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
                img = img.convert("RGBA") # Mantener transparencia
            else:
                img = img.convert("RGB") # Imagen normal sin transparencia

            # 2. CÁLCULO MATEMÁTICO DE ASPECT RATIO (Regla de tres)
            # Calculamos cuánto tendríamos que reducir para encajar en ancho y alto
            original_w, original_h = img.size
            ratio_w = target_w / original_w
            ratio_h = target_h / original_h
            
            # Elegimos el ratio menor para asegurar que quepa entera sin recortar
            scale = min(ratio_w, ratio_h)
            
            # Nuevas dimensiones
            new_w = int(original_w * scale)
            new_h = int(original_h * scale)

            # 3. REDIMENSIONADO DE ALTA CALIDAD (Lanczos)
            img = img.resize((new_w, new_h), PILImage.Resampling.LANCZOS)
            
            output_buffer = io.BytesIO()
            # Guardamos siempre como PNG para preservar calidad y transparencia
            img.save(output_buffer, format="PNG", optimize=True)
            return output_buffer, new_w, new_h
    except Exception as e:
        print(f"Error resize: {e}")
        return None, 0, 0

# --- ENDPOINT PRINCIPAL ---
@app.post("/generate-excel")
async def generate_excel(data: QuotationData):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet()

    # --- 1. FORMATOS ---
    fmt_company = workbook.add_format({'bold': True, 'font_size': 16, 'align': 'center', 'valign': 'vcenter'})
    fmt_info = workbook.add_format({'font_size': 10, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    fmt_red_title = workbook.add_format({'font_color': 'red', 'font_size': 12, 'align': 'center', 'valign': 'vcenter'})
    fmt_seller = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
    
    fmt_table_header = workbook.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    fmt_cell_center = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
    fmt_cell_left = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'text_wrap': True, 'border': 1})
    fmt_money = workbook.add_format({'num_format': '$#,##0.00', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    
    fmt_total_label = workbook.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'align': 'right', 'valign': 'vcenter', 'border': 1})
    fmt_total_value = workbook.add_format({'bold': True, 'num_format': '$#,##0.00', 'bg_color': '#4472C4', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter', 'border': 1})

    # --- 2. COLUMNAS Y FILAS ---
    worksheet.set_column('A:A', 28)
    worksheet.set_column('B:B', 15)
    worksheet.set_column('C:C', 45)
    worksheet.set_column('D:D', 10)
    worksheet.set_column('E:F', 15)

    # Dar altura a las filas de la cabecera para que el logo respire
    for r in range(5):
        worksheet.set_row(r, 25) # 25 puntos de alto por fila en el encabezado

    # --- 3. LOGO (PROCESADO CON TRANSPARENCIA) ---
    try:
        logo_url = "https://konig-kids.com/wp-content/uploads/2023/05/konigkids-logo.png"
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
        
        logo_response = requests.get(logo_url, headers=headers, timeout=5)
        
        if logo_response.status_code == 200:
            # Procesamos el logo con la función inteligente
            logo_buffer, l_w, l_h = smart_resize(logo_response.content, MAX_LOGO_W, MAX_LOGO_H)
            
            if logo_buffer:
                # Calculamos centrado vertical en el espacio de 5 filas (aprox 125 puntos = 166px)
                # Un offset simple de 10px suele bastar
                worksheet.insert_image('A1', 'logo.png', {
                    'image_data': logo_buffer,
                    'x_offset': 15, 
                    'y_offset': 15,
                    'object_position': 1
                })
        else:
            worksheet.write('A3', "LOGO ERR", fmt_cell_center)
    except Exception as e:
        print(f"Error logo: {e}")
        worksheet.write('A3', "LOGO ERR", fmt_cell_center)

    # --- 4. TEXTO CABECERA ---
    worksheet.merge_range('B1:F1', "KONIG KIDS LIMITED", fmt_company)
    worksheet.merge_range('B2:F2', "Add: NO.12 Southern Dengfeng Road, Chenghai District.", fmt_info)
    worksheet.merge_range('B3:F3', "Tel: 0754-89861629 Email: sales@konig-kids.com", fmt_info)
    worksheet.merge_range('B4:F4', "Quotation List", fmt_red_title)
    worksheet.merge_range('B5:F5', "Seller: Agent AI", fmt_seller)

    # --- 5. TABLA ---
    TABLE_HEADER_ROW = 6 
    headers = ["Photo", "Item No.", "Description", "Quantity", "Unit Price", "Amount"]
    for col, text in enumerate(headers):
        worksheet.write(TABLE_HEADER_ROW, col, text, fmt_table_header)

    # --- 6. DESCARGA ASÍNCRONA DE PRODUCTOS ---
    async with aiohttp.ClientSession() as session:
        tasks = [process_image(session, item.image_product) for item in data.items]
        processed_results = await asyncio.gather(*tasks)

    # --- 7. LLENADO ---
    START_DATA_ROW = TABLE_HEADER_ROW + 1
    row_height_points = 120
    
    for i, item in enumerate(data.items):
        current_row = START_DATA_ROW + i
        worksheet.set_row(current_row, row_height_points)
        
        img_buffer, img_w, img_h = processed_results[i]
        
        if img_buffer:
            # Centrado Matemático
            x_off = (CELL_WIDTH_PX - img_w) // 2
            y_off = (CELL_HEIGHT_PX - img_h) // 2

            worksheet.insert_image(current_row, 0, "prod.png", {
                'image_data': img_buffer,
                'x_offset': x_off, 
                'y_offset': y_off,
                'object_position': 1
            })
        else:
            worksheet.write(current_row, 0, "No Image", fmt_cell_center)

        worksheet.write(current_row, 1, item.id_product, fmt_cell_center)
        worksheet.write(current_row, 2, item.product_description, fmt_cell_left)
        worksheet.write(current_row, 3, item.quantity, fmt_cell_center)
        worksheet.write(current_row, 4, item.unit_price, fmt_money)
        worksheet.write(current_row, 5, item.subtotal, fmt_money)

    # --- 8. TOTAL ---
    last_row = START_DATA_ROW + len(data.items)
    worksheet.merge_range(last_row, 0, last_row, 4, "GRAND TOTAL:", fmt_total_label)
    worksheet.write(last_row, 5, data.Total, fmt_total_value)

    workbook.close()
    output.seek(0)

    return Response(
        content=output.getvalue(), 
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=quotation.xlsx"}
    )
