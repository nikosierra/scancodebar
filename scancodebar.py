import os
import cv2
import pandas as pd
import pytesseract
from pyzbar.pyzbar import decode

# Configuración de la carpeta y archivo de salida
IMAGE_FOLDER = r"C:\Users\KTFUS\Downloads\Resultados"  # Ruta de tu carpeta
OUTPUT_EXCEL = os.path.join(IMAGE_FOLDER, "codigos_upc.xlsx")  # Guardará en la misma carpeta

# Especificar la ruta de Tesseract manualmente
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def enlarge_image(img, scale=2):
    """Agrandar la imagen para mejorar la detección OCR."""
    width = int(img.shape[1] * scale)
    height = int(img.shape[0] * scale)
    return cv2.resize(img, (width, height), interpolation=cv2.INTER_LINEAR)

def preprocess_image_for_text(img):
    """Preprocesar la imagen para mejorar la detección de texto."""
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 128, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
    return binary

def read_upc_from_image(image_path):
    """Escanea códigos de barras en una imagen."""
    img = cv2.imread(image_path)

    # Intentar leer códigos de barras directamente
    barcodes = decode(img)
    upc_codes = []
    for barcode in barcodes:
        barcode_data = barcode.data.decode("utf-8")
        barcode_type = barcode.type

        if barcode_type in ["UPC-A", "EAN13"] and len(barcode_data) in [12, 13]:
            upc_codes.append(barcode_data[-12:])  # Extraer los últimos 12 dígitos

    print(f"📸 {os.path.basename(image_path)} -> Códigos de barras detectados: {len(upc_codes)}")

    return upc_codes

def read_text_for_upc(image_path):
    """Usa OCR para leer los números UPC si el código de barras no fue detectado."""
    img = cv2.imread(image_path)

    # Agrandar imagen para mejorar OCR
    img = enlarge_image(img)

    # Convertir a escala de grises para mejorar la lectura
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # Configuración de Tesseract OCR
    custom_config = r"--oem 3 --psm 6"
    text = pytesseract.image_to_string(gray, config=custom_config)

    print(f"\n📝 OCR detectó en {os.path.basename(image_path)}:\n{text}\n")

    # Filtrar solo números de 12 dígitos (UPC), eliminando comillas y caracteres no numéricos
    upc_from_text = []
    for line in text.split("\n"):
        cleaned_line = ''.join(filter(str.isdigit, line.strip()))  # Eliminar caracteres no numéricos
        if len(cleaned_line) == 12:  # Verificar que sea un UPC válido
            upc_from_text.append(cleaned_line)

    return upc_from_text

def second_pass_for_upc(image_path):
    """Realiza una segunda pasada para intentar detectar más códigos UPC."""
    img = cv2.imread(image_path)

    # Preprocesar la imagen para mejorar la detección de texto
    preprocessed_img = preprocess_image_for_text(img)

    # Configuración de Tesseract OCR
    custom_config = r"--oem 3 --psm 6"
    text = pytesseract.image_to_string(preprocessed_img, config=custom_config)

    print(f"\n🔍 Segunda pasada OCR en {os.path.basename(image_path)}:\n{text}\n")

    # Filtrar solo números de 12 dígitos (UPC)
    upc_from_text = []
    for line in text.split("\n"):
        cleaned_line = ''.join(filter(str.isdigit, line.strip()))  # Eliminar caracteres no numéricos
        if len(cleaned_line) == 12:  # Verificar que sea un UPC válido
            upc_from_text.append(cleaned_line)

    return upc_from_text

def enhanced_second_pass(image_path):
    """Realiza una segunda pasada mejorada para intentar detectar más códigos UPC con efectos incrementados."""
    img = cv2.imread(image_path)
    upc_codes = set()

    # Primera mejora: agrandar la imagen con un escalado incrementado
    enlarged_img = enlarge_image(img, scale=1.55)  # Escalado incrementado en un 10%
    custom_config = r"--oem 3 --psm 11"
    text = pytesseract.image_to_string(enlarged_img, config=custom_config)
    for line in text.split("\n"):
        cleaned_line = ''.join(filter(str.isdigit, line.strip()))
        if len(cleaned_line) == 12:
            upc_codes.add(cleaned_line)

    # Segunda mejora: preprocesar la imagen con técnicas avanzadas
    preprocessed_img = preprocess_image_for_text(img)
    text = pytesseract.image_to_string(preprocessed_img, config=custom_config)
    for line in text.split("\n"):
        cleaned_line = ''.join(filter(str.isdigit, line.strip()))
        if len(cleaned_line) == 12:
            upc_codes.add(cleaned_line)

    # Tercera mejora: aplicar ecualización del histograma y detección de bordes con umbrales ajustados
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    equalized = cv2.equalizeHist(gray)
    edges = cv2.Canny(equalized, 120, 210)  # Umbrales ajustados para mayor sensibilidad
    text = pytesseract.image_to_string(edges, config=custom_config)
    for line in text.split("\n"):
        cleaned_line = ''.join(filter(str.isdigit, line.strip()))
        if len(cleaned_line) == 12:
            upc_codes.add(cleaned_line)

    # Cuarta mejora: intentar detectar códigos de barras nuevamente
    barcodes = decode(enlarged_img)
    for barcode in barcodes:
        barcode_data = barcode.data.decode("utf-8")
        barcode_type = barcode.type
        if barcode_type in ["UPC-A", "EAN13"] and len(barcode_data) in [12, 13]:
            upc_codes.add(barcode_data[-12:])

    print(f"🔄 Segunda pasada mejorada en {os.path.basename(image_path)} detectó {len(upc_codes)} códigos adicionales.")
    return list(upc_codes)

def final_pass_for_upc(image_path):
    """Realiza un último intento para detectar más códigos UPC aplicando rotaciones y ajustes adicionales."""
    img = cv2.imread(image_path)
    upc_codes = set()

    # Rotar la imagen en diferentes ángulos
    for angle in [0, 90, 180, 270]:
        rotated_img = rotate_image(img, angle)
        barcodes = decode(rotated_img)
        for barcode in barcodes:
            barcode_data = barcode.data.decode("utf-8")
            barcode_type = barcode.type
            if barcode_type in ["UPC-A", "EAN13"] and len(barcode_data) in [12, 13]:
                upc_codes.add(barcode_data[-12:])

    # Ajustar brillo y contraste
    adjusted_img = adjust_brightness_contrast(img, brightness=30, contrast=50)
    barcodes = decode(adjusted_img)
    for barcode in barcodes:
        barcode_data = barcode.data.decode("utf-8")
        barcode_type = barcode.type
        if barcode_type in ["UPC-A", "EAN13"] and len(barcode_data) in [12, 13]:
            upc_codes.add(barcode_data[-12:])

    # Reintentar con OCR y configuraciones específicas
    custom_config = r"--oem 3 --psm 4"
    text = pytesseract.image_to_string(adjusted_img, config=custom_config)
    for line in text.split("\n"):
        cleaned_line = ''.join(filter(str.isdigit, line.strip()))
        if len(cleaned_line) == 12:
            upc_codes.add(cleaned_line)

    print(f"🔄 Último intento en {os.path.basename(image_path)} detectó {len(upc_codes)} códigos adicionales.")
    return list(upc_codes)

def rotate_image(image, angle):
    """Rota la imagen en el ángulo especificado."""
    (h, w) = image.shape[:2]
    center = (w // 2, h // 2)
    matrix = cv2.getRotationMatrix2D(center, angle, 1.0)
    rotated = cv2.warpAffine(image, matrix, (w, h))
    return rotated

def adjust_brightness_contrast(image, brightness=0, contrast=0):
    """Ajusta el brillo y el contraste de la imagen."""
    img = cv2.convertScaleAbs(image, alpha=1 + contrast / 100.0, beta=brightness)
    return img

def process_images(folder_path, output_excel):
    """Procesa imágenes, extrae códigos UPC de códigos de barras y OCR, y los guarda en Excel."""
    data = []  # Para almacenar todos los códigos UPC detectados
    images_with_missing_upcs = []  # Para rastrear imágenes con códigos faltantes

    for image_name in os.listdir(folder_path):
        image_path = os.path.join(folder_path, image_name)

        if image_path.lower().endswith((".png", ".jpg", ".jpeg")):
            # Detectar códigos de barras con pyzbar
            upc_codes_barcodes = read_upc_from_image(image_path)

            # Detectar códigos con OCR
            upc_codes_ocr = read_text_for_upc(image_path)

            # Realizar una segunda pasada si hay mucho texto
            upc_codes_second_pass = second_pass_for_upc(image_path)

            # Combinar todos los resultados (códigos de barras, OCR y segunda pasada)
            combined_upc_codes = set(upc_codes_barcodes + upc_codes_ocr + upc_codes_second_pass)  # Usar un set para evitar duplicados

            # Verificar si el número de códigos detectados está fuera del rango esperado
            if len(combined_upc_codes) < 12:
                print(f"⚠️ {image_name} tiene menos códigos UPC detectados: {len(combined_upc_codes)}")
                images_with_missing_upcs.append(image_name)

            # Guardar todos los códigos detectados
            for upc in combined_upc_codes:
                data.append({"Imagen": image_name, "Código UPC": upc.zfill(12)})  # Asegurar 12 dígitos con zfill

    # Intentar recuperar códigos faltantes en imágenes problemáticas
    for image_name in images_with_missing_upcs:
        print(f"🔄 Intentando recuperar códigos faltantes en {image_name} con mejoras...")
        image_path = os.path.join(folder_path, image_name)
        additional_upc_codes = enhanced_second_pass(image_path)  # Segunda pasada mejorada
        for upc in additional_upc_codes:
            data.append({"Imagen": image_name, "Código UPC": upc.zfill(12)})

        # Último intento para detectar códigos faltantes
        final_upc_codes = final_pass_for_upc(image_path)
        for upc in final_upc_codes:
            data.append({"Imagen": image_name, "Código UPC": upc.zfill(12)})

    # Consolidar todos los códigos y eliminar duplicados antes de guardar en Excel
    if data:
        df = pd.DataFrame(data).drop_duplicates(subset=["Código UPC"])  # Eliminar duplicados por código UPC
        df.to_excel(output_excel, sheet_name="Códigos Consolidados", index=False)
        print(f"\n✅ Todos los códigos UPC guardados en la hoja 'Códigos Consolidados' del archivo Excel.")
    else:
        print("\n❌ No se detectaron códigos UPC en ninguna imagen.")

# Ejecutar el proceso
process_images(IMAGE_FOLDER, OUTPUT_EXCEL)