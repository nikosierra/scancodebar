import os
import cv2
import pandas as pd
import pytesseract
from pyzbar.pyzbar import decode
import cloudinary
import cloudinary.api
import cloudinary.uploader
import requests
from io import BytesIO
import numpy as np

# Configuración de Cloudinary
cloudinary.config(
    cloud_name="dndyyqimx",  # Reemplaza con tu cloud_name
    api_key="355168315266622",        # Reemplaza con tu api_key
    api_secret="_G1N4KTYKOd815Wt1hsP1ckdhnA"   # Reemplaza con tu api_secret
)

# Configuración de la carpeta y archivo de salida
OUTPUT_EXCEL = r"codigos_upc.xlsx"  # Archivo Excel de salida

# Especificar la ruta de Tesseract manualmente
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Funciones existentes (enlarge_image, preprocess_image_for_text, etc.)
# ... (Mantén todas las funciones que ya tienes en tu script)

def download_image_from_cloudinary(url):
    """Descarga una imagen desde una URL de Cloudinary y la convierte en un array de NumPy."""
    response = requests.get(url)
    if response.status_code == 200:
        img_array = np.array(bytearray(response.content), dtype=np.uint8)
        img = cv2.imdecode(img_array, cv2.IMREAD_COLOR)
        return img
    else:
        print(f"❌ Error al descargar la imagen desde {url}")
        return None

def process_images_from_cloudinary(output_excel):
    """Procesa imágenes desde Cloudinary, extrae códigos UPC y los guarda en Excel."""
    data = []  # Para almacenar todos los códigos UPC detectados

    # Obtener la lista de recursos de Cloudinary
    resources = cloudinary.api.resources(type="upload", resource_type="image")

    for resource in resources.get("resources", []):
        image_url = resource["url"]
        image_name = resource["public_id"]

        print(f"📥 Procesando imagen: {image_name}")

        # Descargar la imagen desde Cloudinary
        img = download_image_from_cloudinary(image_url)
        if img is None:
            continue

        # Detectar códigos de barras con pyzbar
        upc_codes_barcodes = read_upc_from_image(img)

        # Detectar códigos con OCR
        upc_codes_ocr = read_text_for_upc(img)

        # Realizar una segunda pasada si hay mucho texto
        upc_codes_second_pass = second_pass_for_upc(img)

        # Combinar todos los resultados (códigos de barras, OCR y segunda pasada)
        combined_upc_codes = set(upc_codes_barcodes + upc_codes_ocr + upc_codes_second_pass)

        # Guardar todos los códigos detectados
        for upc in combined_upc_codes:
            data.append({"Imagen": image_name, "Código UPC": upc.zfill(12)})

    # Consolidar todos los códigos y eliminar duplicados antes de guardar en Excel
    if data:
        df = pd.DataFrame(data).drop_duplicates(subset=["Código UPC"])
        df.to_excel(output_excel, sheet_name="Códigos Consolidados", index=False)
        print(f"\n✅ Todos los códigos UPC guardados en el archivo Excel: {output_excel}")
    else:
        print("\n❌ No se detectaron códigos UPC en ninguna imagen.")

# Ejecutar el proceso
process_images_from_cloudinary(OUTPUT_EXCEL)
