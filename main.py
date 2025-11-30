from fastapi import FastAPI, Response
from pydantic import BaseModel
from typing import List, Optional
import xlsxwriter
import io
import asyncio
import aiohttp
import requests # Usamos requests solo para el logo (es una sola vez)
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

# --- CONFIGURACIÓN DE DESCARGA ASÍNCRONA (Protección de RAM) ---
sem = asyncio.Semaphore(5) # Máximo 5 descargas a la vez

async def process_image(session, url):
    if not url or not str(url).startswith("http"):
        return None
    
    async with sem: 
        try:
            async with session.get(str(url), timeout=15) as response:
                if response.status == 200:
                    data = await response.read()
                    return await asyncio.to_thread(resize_image_in_memory, data)
        except Exception as e:
            print(f"Error descargando {url}: {e}")
            return None
    return None

def resize_image_in_memory(data):
    try:
        with PILImage.open(io.BytesIO(data)) as img:
            img = img.convert("RGB")
            img.thumbnail((200, 140)) # Ajustar al tamaño de celda
            
            output_buffer = io.BytesIO()
            img.save(output_buffer, format="PNG", optimize=True)
            return output_buffer
    except:
        return None

# --- ENDPOINT PRINCIPAL ---
@app.post("/generate-excel")
async def generate_excel(data: QuotationData):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet()

    # --- 1. FORMATOS DE ESTILO ---
    # Estilos de Cabecera Principal
    fmt_company = workbook.add_format({'bold': True, 'font_size': 16, 'align': 'center', 'valign': 'vcenter'})
    fmt_info = workbook.add_format({'font_size': 10, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    fmt_red_title = workbook.add_format({'font_color': 'red', 'font_size': 12, 'align': 'center', 'valign': 'vcenter'})
    fmt_seller = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})

    # Estilos de Tabla
    fmt_table_header = workbook.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    fmt_cell_center = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
    fmt_cell_left = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'text_wrap': True, 'border': 1})
    fmt_money = workbook.add_format({'num_format': '$#,##0.00', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    fmt_total_label = workbook.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'align': 'right', 'valign': 'vcenter', 'border': 1})
    fmt_total_value = workbook.add_format({'bold': True, 'num_format': '$#,##0.00', 'bg_color': '#4472C4', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter', 'border': 1})

    # --- 2. CONFIGURACIÓN DE COLUMNAS ---
    worksheet.set_column('A:A', 25) # Columna Foto
    worksheet.set_column('B:B', 15) # Item No
    worksheet.set_column('C:C', 45) # Description
    worksheet.set_column('D:D', 10) # Qty
    worksheet.set_column('E:F', 15) # Price & Amount

    # --- 3. DIBUJAR LA CABECERA (LOGO Y DATOS) ---
    
    # A. Insertar Logo (Merge A1:A5)
    worksheet.merge_range('A1:A5', '', fmt_cell_center) # Placeholder merge
    try:
        logo_url = "https://konig-kids.com/wp-content/uploads/2023/05/konigkids-logo.png"
        logo_response = requests.get(logo_url)
        logo_data = io.BytesIO(logo_response.content)
        
        # Insertamos el logo ajustado
        worksheet.insert_image('A1', 'logo.png', {
            'image_data': logo_data,
            'x_scale': 0.6, 'y_scale': 0.6, # Ajuste de escala a ojo para que quepa
            'x_offset': 10, 'y_offset': 10
        })
    except:
        worksheet.write('A3', "LOGO HERE", fmt_cell_center)

    # B. Información de la Empresa (Columnas B a F)
    worksheet.merge_range('B1:F1', "KONIG KIDS LIMITED", fmt_company)
    worksheet.merge_range('B2:F2', "Add: NO.12 Southern Dengfeng Road, Chenghai District.", fmt_info)
    worksheet.merge_range('B3:F3', "Tel: 0754-89861629 Email: sales@konig-kids.com", fmt_info)
    worksheet.merge_range('B4:F4', "Quotation List", fmt_red_title)
    worksheet.merge_range('B5:F5', "Seller: Agent AI", fmt_seller)

    # --- 4. ENCABEZADOS DE TABLA (Fila 6 / Índice 5) ---
    headers = ["Photo", "Item No.", "Description", "Quantity", "Unit Price", "Amount"]
    for col, text in enumerate(headers):
        worksheet.write(5, col, text, fmt_table_header)

    # --- 5. PROCESAMIENTO DE IMÁGENES (Async) ---
    async with aiohttp.ClientSession() as session:
        tasks = [process_image(session, item.image_product) for item in data.items]
        processed_images = await asyncio.gather(*tasks)

    # --- 6. LLENADO DE DATOS (Fila 7 en adelante) ---
    start_row = 6
    row_height = 120 # ~160px
    
    for i, item in enumerate(data.items):
        current_row = start_row + i
        worksheet.set_row(current_row, row_height)
        
        # Imagen
        image_buffer = processed_images[i]
        if image_buffer:
            worksheet.insert_image(current_row, 0, "prod.png", {
                'image_data': image_buffer,
                'x_offset': 10, 'y_offset': 10, 'object_position': 1
            })
        else:
            worksheet.write(current_row, 0, "No Image", fmt_cell_center)

        # Datos
        worksheet.write(current_row, 1, item.id_product, fmt_cell_center)
        worksheet.write(current_row, 2, item.product_description, fmt_cell_left)
        worksheet.write(current_row, 3, item.quantity, fmt_cell_center)
        worksheet.write(current_row, 4, item.unit_price, fmt_money)
        worksheet.write(current_row, 5, item.subtotal, fmt_money)

    # --- 7. TOTAL FINAL ---
    last_row = start_row + len(data.items)
    worksheet.merge_range(last_row, 0, last_row, 4, "GRAND TOTAL:", fmt_total_label) # Merge A-E
    worksheet.write(last_row, 5, data.Total, fmt_total_value)

    workbook.close()
    output.seek(0)

    return Response(
        content=output.getvalue(), 
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=quotation.xlsx"}
    )
