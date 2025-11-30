from fastapi import FastAPI, Response
from pydantic import BaseModel
from typing import List, Optional
import xlsxwriter
import io
import requests
from PIL import Image as PILImage

app = FastAPI()

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

@app.post("/generate-excel")
def generate_excel(data: QuotationData):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet()

    # Formatos
    bold = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    cell_center = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
    cell_left = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'text_wrap': True, 'border': 1})
    money_fmt = workbook.add_format({'num_format': '$#,##0.00', 'align': 'right', 'valign': 'vcenter', 'border': 1, 'font_color': '#006100', 'bold': True})
    total_fmt = workbook.add_format({'num_format': '$#,##0.00', 'align': 'right', 'valign': 'vcenter', 'border': 1, 'bg_color': '#4472C4', 'font_color': 'white', 'bold': True})

    # Columnas
    worksheet.set_column('A:A', 30)
    worksheet.set_column('B:B', 12)
    worksheet.set_column('C:C', 40)
    worksheet.set_column('D:D', 10)
    worksheet.set_column('E:F', 15)

    headers = ["Image", "Item No.", "Description", "Quantity", "Unit Price", "Amount"]
    for col, text in enumerate(headers):
        worksheet.write(0, col, text, header_fmt)

    row_idx = 1
    row_height = 120 # 120 puntos = ~160px

    for item in data.items:
        worksheet.set_row(row_idx, row_height)
        
        # Imagen
        if item.image_product and item.image_product.startswith("http"):
            try:
                img_data = requests.get(item.image_product, stream=True).raw
                with PILImage.open(img_data) as img:
                    img_buffer = io.BytesIO()
                    img = img.convert("RGB")
                    img.thumbnail((200, 140))
                    img.save(img_buffer, format="PNG")
                    worksheet.insert_image(row_idx, 0, "img.png", {
                        'image_data': img_buffer, 'x_offset': 10, 'y_offset': 10, 'object_position': 1
                    })
            except:
                worksheet.write(row_idx, 0, "Error Img", cell_center)
        else:
            worksheet.write(row_idx, 0, "No Image", cell_center)

        worksheet.write(row_idx, 1, item.id_product, cell_center)
        worksheet.write(row_idx, 2, item.product_description, cell_left)
        worksheet.write(row_idx, 3, item.quantity, cell_center)
        worksheet.write(row_idx, 4, item.unit_price, money_fmt)
        worksheet.write(row_idx, 5, item.subtotal, money_fmt)
        row_idx += 1

    worksheet.write(row_idx, 4, "GRAND TOTAL:", total_fmt)
    worksheet.write(row_idx, 5, data.Total, total_fmt)

    workbook.close()
    output.seek(0)

    return Response(
        content=output.getvalue(), 
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=quotation.xlsx"}
    )