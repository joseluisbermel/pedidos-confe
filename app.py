import io
import re
import openpyxl
from openpyxl.styles import PatternFill
from PIL import Image
import pdf2image
from pypdf import PdfReader
import streamlit as st
import cv2
import numpy as np

st.set_page_config(
    page_title="Gestor de Pedidos Escolares", page_icon="🍬", layout="wide"
)

# Estilo de alerta en amarillo chillón para celdas dudosas
YELLOW_FILL = PatternFill(
    start_color="FFFF00", end_color="FFFF00", fill_type="solid"
)


def segmentar_y_leer_cantidades(pil_image):
    """
    Usa OpenCV para localizar la tabla, aislar la columna de cantidades,
    y extraer los números fila por fila de forma geométrica estructurada.
    """
    # Convertir imagen PIL a formato OpenCV (Grises)
    img = cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2BGR)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    
    # Binarizar la imagen para encontrar las líneas negras
    _, thresh = cv2.threshold(gray, 180, 255, cv2.THRESH_BINARY_INV)
    
    # Aislar las líneas horizontales y verticales
    kernel_h = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 1))
    lineas_h = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel_h)
    
    kernel_v = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 40))
    lineas_v = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel_v)
    
    # Encontrar la cuadrícula de la tabla
    tabla_lineas = cv2.add(lineas_h, lineas_v)
    contornos, _ = cv2.findContours(tabla_lineas, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    cantidades_detectadas = {}
    if not contornos:
        return cantidades_detectadas

    # Obtener el recuadro de la tabla principal
    tabla_macro = max(contornos, key=cv2.contourArea)
    x, y, w, h = cv2.boundingRect(tabla_macro)
    
    # La columna cantidad está en el último 15% del ancho de la tabla
    x_cant = x + int(w * 0.85)
    w_cant = int(w * 0.15)
    
    # Dividir el bloque vertical simétricamente en 30 filas (C001 a C030)
    alto_fila = h / 30.0
    
    import pytesseract
    for i in range(30):
        y_fila_inicio = y + int(i * alto_fila)
        y_fila_fin = y + int((i + 1) * alto_fila)
        
        # Recortar la celda exacta de cantidad para el artículo actual
        celda = gray[y_fila_inicio:y_fila_fin, x_cant:x_cant+w_cant]
        
        # Pasar un OCR optimizado exclusivamente para dígitos en una celda pequeña
        texto_celda = pytesseract.image_to_string(
            celda, config='--psm 10 --oem 3 -c tessedit_char_whitelist=0123456789'
        ).strip()
        
        ref_code = f"C{i+1:03d}"
        if texto_celda.isdigit():
            val = int(texto_celda)
            if val > 0 and val < 50:  # Descartar ruidos o manchas que den números absurdos
                cantidades_detectadas[ref_code] = val
                
    return cantidades_detectadas


def extraer_texto_pagina(page, pdf_bytes, page_idx):
  """Extrae texto digital o aplica OCR completo de la página si es necesario."""
  texto = page.extract_text() or ""
  if len(texto.strip()) > 40:
    return texto, None

  try:
    import pytesseract
    images = pdf2image.convert_from_bytes(
        pdf_bytes, first_page=page_idx + 1, last_page=page_idx + 1
    )
    if images:
      texto_ocr = pytesseract.image_to_string(images[0], lang="spa")
      return texto_ocr, images[0]
  except Exception as e:
    st.warning(f"Aviso en página {page_idx+1}: {e}")

  return texto, None


def extraer_pedidos_agrupados(pdf_bytes):
  """Procesa el PDF unificando cabeceras y aplicando segmentación visual."""
  reader = PdfReader(io.BytesIO(pdf_bytes))
  pedidos = []
  alumno_actual = None

  for idx, page in enumerate(reader.pages):
    texto, pil_image = extraer_texto_pagina(page, pdf_bytes, idx)
    lineas = [l.strip() for l in texto.split("\n") if l.strip()]

    contiene_nombre = any(
        x in texto.upper()
        for x in ["NOMBRE Y APELLIDOS:", "HOJA DE PEDIDOS"]
    )

    if alumno_actual is None or contiene_nombre:
      if alumno_actual is not None:
        pedidos.append(alumno_actual)

      nombre = "Ficha_Anonima"
      telefono = ""

      for l_idx, linea in enumerate(lineas[:8]):
        if "NOMBRE" in linea.upper():
          if ":" in linea:
            nombre = linea.split(":")[-1].strip()
          if (not nombre or len(nombre) < 3) and l_idx + 1 < len(lineas):
            nombre = lineas[l_idx + 1].strip()
          break

      nombre = re.sub(r"(?:TELÉFONO|TELEFONO|CURSO|FECHA|CENTRO).*", "", nombre, flags=re.IGNORECASE).strip()
      nombre = re.sub(r"[^a-zA-ZÁÉÍÓÚáéíóúÑñ\s]", "", nombre).strip()

      for linea in lineas[:10]:
        match_tel = re.search(r"(?:TELÉFONO|TELEFONO|TELF)[:\s]*([\d\s-]{6,15})", linea, re.IGNORECASE)
        if match_tel:
          telefono = re.sub(r"\D", "", match_tel.group(1))
          break

      alumno_actual = {
          "nombre": nombre if len(nombre) > 2 else f"REVISAR_ALUMNO_PAG_{idx + 1}",
          "telefono": telefono,
          "cantidades": {},
          "paginas": [idx + 1],
      }
    else:
      alumno_actual["paginas"].append(idx + 1)

    # --- EXTRACCIÓN DE CANTIDADES ---
    if pil_image is not None:
        # Si es una foto/imagen, usamos el motor geométrico de OpenCV por celdas aisladas
        cantidades_img = segmentar_y_leer_cantidades(pil_image)
        alumno_actual["cantidades"].update(cantidades_img)
    else:
        # Si es texto digital nativo, aplicamos el procesado por línea seguro
        for linea in lineas:
          match_ref = re.search(r"\b(C\d{3})\b", linea)
          if match_ref:
            ref_code = match_ref.group(1)
            linea_limpia = re.sub(r"\d+\s*(?:g|gr|kg|ml)\b", "", linea, flags=re.IGNORECASE)
            numeros = [int(n) for n in re.findall(r"\b\d+\b", linea_limpia)]
            if numeros:
              cantidad = numeros[-1]
              if cantidad > 0 and cantidad != int(ref_code[1:]):
                alumno_actual["cantidades"][ref_code] = cantidad

    if any("TOTAL" in l.upper() and "ART" in l.upper() for l in lineas):
      pedidos.append(alumno_actual)
      alumno_actual = None

  if alumno_actual is not None:
    pedidos.append(alumno_actual)

  return pedidos


def rellenar_excel_dinamico(excel_bytes, pedidos):
  """Mapea los datos extraídos en las pestañas del libro de Excel original."""
  wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
  hojas_disponibles = wb.sheetnames

  hojas_alumnos_base = [
      h for h in hojas_disponibles
      if h.upper() not in ["RESUMEN", "INSTRUCCIONES", "DATOS", "TARIFA", "TOTALES"]
  ]

  if not hojas_alumnos_base:
    raise ValueError("Error: Tu plantilla Excel no contiene pestañas individuales válidas.")

  ws_plantilla = wb[hojas_alumnos_base[0]]
  alertas = []

  for i, pedido in enumerate(pedidos):
    if i < len(hojas_alumnos_base):
      ws = wb[hojas_alumnos_base[i]]
      if "REVISAR" not in pedido["nombre"]:
        ws.title = pedido["nombre"][:30]
    else:
      ws = wb.copy_worksheet(ws_plantilla)
      ws.title = f"{pedido['nombre'][:25]}_{i+1}"

    celda_nombre_encontrada = False
    for row in range(1, 15):
      for col in range(1, 10):
        val = str(ws.cell(row=row, column=col).value or "").upper()
        if "ALUMNO" in val or "NOMBRE" in val:
          celda_destino = ws.cell(row=row, column=col + 1)
          if "REVISAR" in pedido["nombre"]:
            celda_destino.value = f"REVISAR (Pág {pedido['paginas']})"
            celda_destino.fill = YELLOW_FILL
            alertas.append(f"⚠️ Páginas {pedido['paginas']}: No se pudo leer el nombre manuscrito.")
          else:
            celda_destino.value = pedido["nombre"]
          celda_nombre_encontrada = True
          break
      if celda_nombre_encontrada: break

    mapa_celdas_ref = {}
    for row in range(1, ws.max_row + 1):
      for col in range(1, 10):
        val = str(ws.cell(row=row, column=col).value or "").strip()
        if re.match(r"^C\d{3}$", val):
          mapa_celdas_ref[val] = (row, col + 3)

    for ref_code, cantidad in pedido["cantidades"].items():
      if ref_code in mapa_celdas_ref:
        r, c = mapa_celdas_ref[ref_code]
        ws.cell(row=r, column=c, value=cantidad)

  output = io.BytesIO()
  wb.save(output)
  output.seek(0)
  return output.getvalue(), alertas


# --- INTERFAZ DE USUARIO (STREAMLIT WEB) ---
st.title("🍬 Gestor Automatizado de Pedidos - E. Moreno")
st.markdown("Carga tu PDF desordenado y tu plantilla Excel. El sistema aislará visualmente las columnas de cantidades.")

st.sidebar.header("Opciones del Informe")
nombre_colegio = st.sidebar.text_input("Centro Escolar", value="IES_Lopez_de_Arenas")
fecha_pedido = st.sidebar.date_input("Fecha de emisión")

col1, col2 = st.columns(2)
with col1:
  st.subheader("1. Archivo PDF")
  pdf_file = st.file_uploader("Arrastra el PDF de las hojas de pedido", type=["pdf"])
with col2:
  st.subheader("2. Plantilla de Destino (.xlsx)")
  excel_file = st.file_uploader("Arrastra la plantilla Excel vacía", type=["xlsx"])

if st.button("🚀 Ejecutar Consolidación de Pedidos", type="primary", use_container_width=True):
  if not pdf_file or not excel_file:
    st.error("❌ Por favor, sube ambos archivos antes de ejecutar el proceso.")
  else:
    with st.spinner("Procesando celdas y alineando datos en segundo plano..."):
        try:
            pdf_bytes = pdf_file.read()
            pedidos = extraer_pedidos_agrupados(pdf_bytes)
            excel_procesado, alertas = rellenar_excel_dinamico(excel_file.read(), pedidos)

            st.success(f"✨ ¡Conseguido! Se han estructurado {len(pedidos)} alumnos en tu Excel.")

            if alertas:
              with st.expander("⚠️ Alertas de verificación manual:", expanded=True):
                for al in alertas: 
                    st.info(al)

            filename_clean = f"Pedidos_{nombre_colegio}_{fecha_pedido.strftime('%d%m_%Y')}.xlsx"
            st.download_button(
                label="📥 Descargar Libro de Excel Consolidado",
                data=excel_procesado,
                file_name=filename_clean,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"Se ha producido un error durante el volcado de datos: {str(e)}")
