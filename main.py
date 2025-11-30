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
# Definimos el tamaño aproximado de la celda en píxeles para los cálculos de centrado
# Ancho col A (25) es aprox 200px. Alto fila (120pt) es aprox 160px.
CELL_WIDTH_PX = 200
CELL_HEIGHT_PX = 160
# Definimos un tamaño máximo para la imagen un poco menor para dejar margen
MAX_IMG_WIDTH = 190
MAX_IMG_HEIGHT = 150


async def process_image(session, url):
    if not url or not str(url).startswith("http"):
        # Devolvemos None y dimensiones 0 si no hay URL
        return None, 0, 0
    
    async with sem: 
        try:
            async with session.get(str(url), timeout=15) as response:
                if response.status == 200:
                    data = await response.read()
                    # Procesar en hilo aparte y obtener dimensiones
                    return await asyncio.to_thread(resize_image_and_get_dims, data)
        except Exception as e:
            print(f"Error descargando {url}: {e}")
            return None, 0, 0
    return None, 0, 0

def resize_image_and_get_dims(data):
    # Esta función ahora devuelve TRES cosas: el buffer, el ancho final y el alto final
    try:
        with PILImage.open(io.BytesIO(data)) as img:
            img = img.convert("RGB")
            # Redimensionar proporcionalmente para que quepa en nuestro cuadro máximo
            img.thumbnail((MAX_IMG_WIDTH, MAX_IMG_HEIGHT))
            
            # Obtener las dimensiones finales exactas de la imagen redimensionada
            final_w, final_h = img.size
            
            output_buffer = io.BytesIO()
            img.save(output_buffer, format="PNG", optimize=True)
            return output_buffer, final_w, final_h
    except:
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

    # --- 2. COLUMNAS ---
    worksheet.set_column('A:A', 25) # Aprox 200px ancho
    worksheet.set_column('B:B', 15)
    worksheet.set_column('C:C', 45)
    worksheet.set_column('D:D', 10)
    worksheet.set_column('E:F', 15)

    # --- 3. CABECERA Y LOGO ---
    
    # A. Logo (Insertar en A1 sin combinar, dejar que flote)
    try:
        logo_url = "https://konig-kids.com/wp-content/uploads/2023/05/konigkids-logo.png"
        # Usamos requests síncrono porque es solo una imagen pequeña al inicio
        logo_response = requests.get(logo_url, timeout=5)
        logo_data = io.BytesIO(logo_response.content)
        
        worksheet.insert_image('A1', 'logo.png', {
            'image_data': logo_data,
            'x_scale': 0.7, 'y_scale': 0.7, # Escala ajustada para que se vea bien
            'x_offset': 5, 'y_offset': 5    # Pequeño margen superior izquierdo
        })
    except Exception as e:
        print(f"Error logo: {e}")
        worksheet.write('A3', "LOGO ERROR", fmt_cell_center)

    # B. Texto (Combinar B a F)
    worksheet.merge_range('B1:F1', "KONIG KIDS LIMITED", fmt_company)
    worksheet.merge_range('B2:F2', "Add: NO.12 Southern Dengfeng Road, Chenghai District.", fmt_info)
    worksheet.merge_range('B3:F3', "Tel: 0754-89861629 Email: sales@konig-kids.com", fmt_info)
    worksheet.merge_range('B4:F4', "Quotation List", fmt_red_title)
    worksheet.merge_range('B5:F5', "Seller: Agent AI", fmt_seller)

    # --- 4. ENCABEZADOS TABLA ---
    headers = ["Photo", "Item No.", "Description", "Quantity", "Unit Price", "Amount"]
    for col, text in enumerate(headers):
        worksheet.write(5, col, text, fmt_table_header)

    # --- 5. DESCARGA ASÍNCRONA ---
    async with aiohttp.ClientSession() as session:
        tasks = [process_image(session, item.image_product) for item in data.items]
        # processed_results será una lista de tuplas: [(buffer, w, h), (buffer, w, h), ...]
        processed_results = await asyncio.gather(*tasks)

    # --- 6. LLENADO DE DATOS ---
    start_row = 6
    row_height_points = 120 # 120 puntos = ~160px alto
    
    for i, item in enumerate(data.items):
        current_row = start_row + i
        worksheet.set_row(current_row, row_height_points)
        
        # Obtener datos de la imagen procesada
        img_buffer, img_w, img_h = processed_results[i]
        
        if img_buffer:
            # --- CÁLCULO DE CENTRADO PERFECTO ---
            # offset = (espacio_disponible - tamaño_imagen) / 2
            x_off = max(0, (CELL_WIDTH_PX - img_w) // 2)
            y_off = max(0, (CELL_HEIGHT_PX - img_h) // 2)

            worksheet.insert_image(current_row, 0, "prod.png", {
                'image_data': img_buffer,
                'x_offset': x_off, 
                'y_offset': y_off,
                'object_position': 1 # Mover y cambiar tamaño con celdas
            })
        else:
            worksheet.write(current_row, 0, "No Image", fmt_cell_center)

        # Resto de datos
        worksheet.write(current_row, 1, item.id_product, fmt_cell_center)
        worksheet.write(current_row, 2, item.product_description, fmt_cell_left)
        worksheet.write(current_row, 3, item.quantity, fmt_cell_center)
        worksheet.write(current_row, 4, item.unit_price, fmt_money)
        worksheet.write(current_row, 5, item.subtotal, fmt_money)

    # --- 7. TOTAL FINAL ---
    last_row = start_row + len(data.items)
    worksheet.merge_range(last_row, 0, last_row, 4, "GRAND TOTAL:", fmt_total_label)
    worksheet.write(last_row, 5, data.Total, fmt_total_value)

    workbook.close()
    output.seek(0)

    return Response(
        content=output.getvalue(), 
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=quotation.xlsx"}
    )
