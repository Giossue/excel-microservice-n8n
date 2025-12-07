from fastapi import FastAPI, Response, Request
from pydantic import BaseModel, Field
from typing import List, Optional, Union
import xlsxwriter
import io
import asyncio
import aiohttp
import requests
from PIL import Image as PILImage

app = FastAPI()

# --- CONFIGURACIÓN ---
sem = asyncio.Semaphore(5)

# Celdas y Límites
CELL_WIDTH_PX = 210
CELL_HEIGHT_PX = 160
MAX_PROD_W = 180
MAX_PROD_H = 130
MAX_LOGO_W = 180
MAX_LOGO_H = 100

# --- MODELOS DE DATOS (CASO 1: COTIZACIÓN) ---
class Product(BaseModel):
    image_product: Optional[str] = None
    id_product: str
    product_description: str
    quantity: float # Cambiado a float por si vienen decimales
    unit_price: float
    subtotal: float

class QuotationData(BaseModel):
    items: List[Product]
    Total: float

# --- MODELOS DE DATOS (CASO 2: LISTA DE PRODUCTOS) ---
class ProductItemSpec(BaseModel):
    url_image: Optional[str] = None
    description: str
    ITEM_REFERENCE_NO: str
    rate: Optional[str] = ""
    CARTON_MEASUREMENT: Optional[str] = ""
    CBM: Optional[str] = ""
    GROSS_WEIGHT_KGS: Optional[str] = ""
    MOQ_PCS: Optional[str] = ""
    NET_WEIGHT_KGS: Optional[str] = ""
    PACKAGE_SIZE: Optional[str] = ""
    PACKAGING_TYPE: Optional[str] = ""
    QTY_PCS: Optional[str] = ""
    REMARKS: Optional[str] = ""

class ProductListData(BaseModel):
    products: List[ProductItemSpec]

# --- FUNCIONES AUXILIARES (IMAGENES) ---
def smart_resize(data, target_w, target_h):
    try:
        with PILImage.open(io.BytesIO(data)) as img:
            if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
                img = img.convert("RGBA")
            else:
                img = img.convert("RGB")

            original_w, original_h = img.size
            ratio_w = target_w / original_w
            ratio_h = target_h / original_h
            scale = min(ratio_w, ratio_h)
            new_w = int(original_w * scale)
            new_h = int(original_h * scale)

            img = img.resize((new_w, new_h), PILImage.Resampling.LANCZOS)
            output_buffer = io.BytesIO()
            img.save(output_buffer, format="PNG", optimize=True)
            return output_buffer, new_w, new_h
    except Exception as e:
        print(f"Error resize: {e}")
        return None, 0, 0

async def process_image(session, url):
    if not url or not str(url).startswith("http"):
        return None, 0, 0
    
    async with sem: 
        try:
            headers = {'User-Agent': 'Mozilla/5.0'}
            async with session.get(str(url), headers=headers, timeout=15) as response:
                if response.status == 200:
                    data = await response.read()
                    return await asyncio.to_thread(smart_resize, data, MAX_PROD_W, MAX_PROD_H)
        except Exception as e:
            print(f"Error descargando {url}: {e}")
            return None, 0, 0
    return None, 0, 0

# --- LÓGICA GENERACIÓN CASO 1: COTIZACIÓN (Original) ---
async def create_quotation_sheet(workbook, data: QuotationData):
    worksheet = workbook.add_worksheet("Quotation")
    
    # Formatos
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

    # Columnas
    worksheet.set_column('A:A', 28)
    worksheet.set_column('B:B', 15)
    worksheet.set_column('C:C', 45)
    worksheet.set_column('D:D', 10)
    worksheet.set_column('E:F', 15)
    
    for r in range(5): worksheet.set_row(r, 25)

    # Logo y Cabecera
    # (Aquí iría la lógica del logo, simplificada para el ejemplo)
    worksheet.merge_range('B1:F1', "KONIG KIDS LIMITED", fmt_company)
    worksheet.merge_range('B2:F2', "Add: NO.12 Southern Dengfeng Road, Chenghai District.", fmt_info)
    worksheet.merge_range('B4:F4', "Quotation List", fmt_red_title)
    worksheet.merge_range('B5:F5', "Seller: Agent AI", fmt_seller)

    # Tabla
    TABLE_HEADER_ROW = 6 
    headers = ["Photo", "Item No.", "Description", "Quantity", "Unit Price", "Amount"]
    for col, text in enumerate(headers):
        worksheet.write(TABLE_HEADER_ROW, col, text, fmt_table_header)

    # Descarga de imagenes
    async with aiohttp.ClientSession() as session:
        tasks = [process_image(session, item.image_product) for item in data.items]
        processed_results = await asyncio.gather(*tasks)

    START_DATA_ROW = TABLE_HEADER_ROW + 1
    
    for i, item in enumerate(data.items):
        current_row = START_DATA_ROW + i
        worksheet.set_row(current_row, 120) # Altura fija
        
        img_buffer, img_w, img_h = processed_results[i]
        if img_buffer:
            x_off = (CELL_WIDTH_PX - img_w) // 2
            y_off = (CELL_HEIGHT_PX - img_h) // 2
            worksheet.insert_image(current_row, 0, "prod.png", {'image_data': img_buffer, 'x_offset': x_off, 'y_offset': y_off, 'object_position': 1})
        else:
            worksheet.write(current_row, 0, "No Image", fmt_cell_center)

        worksheet.write(current_row, 1, item.id_product, fmt_cell_center)
        worksheet.write(current_row, 2, item.product_description, fmt_cell_left)
        worksheet.write(current_row, 3, item.quantity, fmt_cell_center)
        worksheet.write(current_row, 4, item.unit_price, fmt_money)
        worksheet.write(current_row, 5, item.subtotal, fmt_money)

    last_row = START_DATA_ROW + len(data.items)
    worksheet.merge_range(last_row, 0, last_row, 4, "GRAND TOTAL:", fmt_total_label)
    worksheet.write(last_row, 5, data.Total, fmt_total_value)

# --- LÓGICA GENERACIÓN CASO 2: LISTA DE PRODUCTOS (Nuevo) ---
async def create_product_list_sheet(workbook, data: ProductListData):
    worksheet = workbook.add_worksheet("Product List")
    
    # Formatos
    fmt_header = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
    fmt_cell = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
    
    # Definir Columnas y Anchos
    # A=Photo, B=Ref, C=Desc, D=Rate, E=Carton, F=CBM, G=GW, H=MOQ, I=NW, J=PkgSize, K=PkgType, L=Qty, M=Remarks
    headers = [
        "Photo", "Item Ref.", "Description", "Rate", "Carton Meas.", 
        "CBM", "G.W.", "MOQ", "N.W.", "Pkg Size", "Pkg Type", "Qty", "Remarks"
    ]
    
    # Ajuste de anchos para que quepa todo
    worksheet.set_column('A:A', 25) # Photo
    worksheet.set_column('B:B', 15) # Ref
    worksheet.set_column('C:C', 30) # Description (Más ancho)
    worksheet.set_column('D:M', 12) # El resto estándar

    # Escribir cabecera
    for col, text in enumerate(headers):
        worksheet.write(0, col, text, fmt_header)

    # Descarga de imagenes
    async with aiohttp.ClientSession() as session:
        # Nota: aquí la clave es 'url_image' no 'image_product'
        tasks = [process_image(session, item.url_image) for item in data.products]
        processed_results = await asyncio.gather(*tasks)

    # Llenado de datos
    for i, item in enumerate(data.products):
        row = i + 1
        worksheet.set_row(row, 120) # Altura fija para imágenes

        # Insertar imagen
        img_buffer, img_w, img_h = processed_results[i]
        if img_buffer:
            # Recalculamos centrado horizontal basado en ancho de col A (aprox 180px útil)
            x_off = (190 - img_w) // 2 
            y_off = (CELL_HEIGHT_PX - img_h) // 2
            worksheet.insert_image(row, 0, "prod.png", {'image_data': img_buffer, 'x_offset': max(5, x_off), 'y_offset': y_off, 'object_position': 1})
        else:
            worksheet.write(row, 0, "No Image", fmt_cell)

        # Mapeo de datos
        worksheet.write(row, 1, item.ITEM_REFERENCE_NO, fmt_cell)
        worksheet.write(row, 2, item.description, fmt_cell)
        worksheet.write(row, 3, item.rate, fmt_cell)
        worksheet.write(row, 4, item.CARTON_MEASUREMENT, fmt_cell)
        worksheet.write(row, 5, item.CBM, fmt_cell)
        worksheet.write(row, 6, item.GROSS_WEIGHT_KGS, fmt_cell)
        worksheet.write(row, 7, item.MOQ_PCS, fmt_cell)
        worksheet.write(row, 8, item.NET_WEIGHT_KGS, fmt_cell)
        worksheet.write(row, 9, item.PACKAGE_SIZE, fmt_cell)
        worksheet.write(row, 10, item.PACKAGING_TYPE, fmt_cell)
        worksheet.write(row, 11, item.QTY_PCS, fmt_cell)
        worksheet.write(row, 12, item.REMARKS, fmt_cell)

# --- ENDPOINT PRINCIPAL UNIFICADO ---
@app.post("/generate-excel")
async def generate_excel(data: Union[QuotationData, ProductListData]):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})

    # Detección inteligente de tipo
    # Si el objeto tiene el atributo 'products', es el Caso 2
    if hasattr(data, 'products') and data.products is not None:
        await create_product_list_sheet(workbook, data)
        fname = "product_list.xlsx"
    else:
        # Asumimos Caso 1 (Cotización) por defecto
        # Nota: Pydantic ya habrá validado la estructura 'items' y 'Total' si entró como QuotationData
        await create_quotation_sheet(workbook, data)
        fname = "quotation.xlsx"

    workbook.close()
    output.seek(0)

    return Response(
        content=output.getvalue(), 
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={fname}"}
    )