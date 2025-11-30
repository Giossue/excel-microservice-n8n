# 🚀 Excel Microservice para n8n (FastAPI + Async I/O)

Este repositorio contiene un microservicio diseñado para generar archivos Excel (formato XLSX) de manera robusta y eficiente. Está optimizado para ser consumido desde flujos de automatización, especialmente n8n.

---

## 💡 El Problema que Resuelve

El servicio soluciona las limitaciones del método HTML para generar Excels, especialmente con grandes lotes de datos y el procesamiento de imágenes:

1.  **Adiós Alertas de Seguridad:** Genera un archivo XLSX binario nativo (no un archivo de texto disfrazado), eliminando las alertas de seguridad de Microsoft Excel.
2.  **Procesamiento Masivo (500+ Items):** Utiliza Python Asíncrono (`aiohttp`) para descargar imágenes de productos en paralelo, evitando los *Timeouts* de las plataformas PaaS (como Render o Hugging Face).
3.  **Formato Perfecto:** Incrusta imágenes, redimensiona el contenido con `Pillow` y garantiza la altura de fila exacta (160px).

---

## 🛠️ Arquitectura y Tecnologías Clave

| Componente | Función |
| :--- | :--- |
| **Servidor API** | FastAPI (Servidor Python de alto rendimiento) |
| **Generación Excel** | `xlsxwriter` (Crea el binario XLSX) |
| **Procesamiento Imágenes** | `Pillow` (Redimensionamiento) |
| **Concurrencia** | `aiohttp` y `asyncio` (Descarga paralela de URLs) |
| **Despliegue** | Docker |

---

## 🔌 Uso de la API desde n8n

El servicio expone un único endpoint que espera un JSON con la estructura del pedido.

### 1. Endpoint
-   **URL:** `http://[TU_IP_O_DOMINIO]:8000/generate-excel` (usando tu VPS)
-   **Method:** `POST`

### 2. Estructura del Body (JSON Requerido)

El nodo HTTP Request en n8n debe enviar un objeto JSON con las claves `items` (lista) y `Total` (float):

```json
{
  "items": [
    {
      "image_product": "URL de la imagen del producto (http://...)",
      "id_product": "63599",
      "product_description": "Descripción del artículo",
      "quantity": 50,
      "unit_price": 14.3,
      "subtotal": 715.00
    }
    // ... hasta 500 artículos
  ],
  "Total": 10107.50
}
