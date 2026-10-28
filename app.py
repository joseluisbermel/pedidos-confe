import io
import re
import streamlit as st
from pypdf import PdfReader
import openpyxl
from openpyxl.styles import PatternFill
from PIL import Image
import pdf2image

st.set_page_config(
    page_title="Gestor de Pedidos Escolares (Imágenes y PDFs)",
    page_icon="🍬",
    layout="wide"
)

YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

def extraer_texto_de_imagen(img):
    """
    Utiliza pytesseract OCR para extraer todo el texto visible de una imagen 
    (ya sea una página de PDF escaneada o una foto suelta).
    """
    try:
        import pytesseract
        # Configuramos tesseract en español
        texto = pytesseract.image_to_string(img, lang='spa')
        return texto
    except Exception as e:
        st.error(f"Error al procesar OCR en imagen: {e}")
        return ""

def procesar_contenido_imagen(img, origen_nombre=""):
    """
    Analiza el texto extraído por OCR de una hoja de pedido (foto o PDF escaneado)
    y devuelve un diccionario estructurado con el alumno, teléfono, curso y cantidades.
    """
    texto = extraer_texto_de_imagen(img)
    lineas = [l.strip() for l in texto.split("\n") if l.strip()]
    
    nombre = ""
    telefono = ""
    curso = ""
    
    # 1. Búsqueda de Nombre y Apellidos
    for l_idx, linea in enumerate(lineas):
        if "NOMBRE" in linea.upper() or "APELLIDOS" in linea.upper():
            if ":" in linea:
                nombre = linea.split(":")[-1].strip()
            if not nombre and l_idx + 1 < len(lineas):
                nombre = lineas[l_idx + 1].strip()
            break
            
    if not nombre or len(nombre) < 3:
        for linea in lineas:
            if linea == linea.upper() and len(linea) > 6 and not "EMORENO" in linea.upper() and not "PEDIDOS" in linea.upper() and not "PROMOCIÓN" in linea.upper():
                nombre = linea
                break

    nombre = re.sub(r"TELÉFONO.*", "", nombre, flags=re.IGNORECASE).strip()
    nombre = re.sub(r"CURSO.*", "", nombre, flags=re.IGNORECASE).strip()
    nombre = re.sub(r"FECHA.*", "", nombre, flags=re.IGNORECASE).strip()

    # 2. Búsqueda de Teléfono
    for linea in lineas:
        match_tel = re.search(r"(?:TELÉFONO|TELEFONO|TELF|TELF\.)[:\s]*([\d\s]{6,15})", linea, re.IGNORECASE)
        if match_tel:
            telefono = match_tel.group(1).strip()
            break

    # 3. Búsqueda de Curso / Centro
    for linea in lineas:
        match_cur = re.search(r"(?:CURSO/CENTRO|CURSO|CENTRO)[:\s]*(.*)", linea, re.IGNORECASE)
        if match_cur:
            curso = match_cur.group(1).strip()
            break

    # 4. Extracción de Referencias (C001 - C030) y sus cantidades
    cantidades = {}
    leyendo_tabla = True
    for linea in lineas:
        if "TOTAL" in linea.upper() and ("ARTÍCULO" in linea.upper() or "ARTICULO" in linea.upper()):
            leyendo_tabla = False
            
        # Buscar código de artículo C001 hasta C030 o número de artículo suelto
        match_ref = re.search(r"\b(C?\d{1,2})\b", linea)
        # También buscamos explícitamente C001-C030
        match_c_ref = re.search(r"\b(C\d{3})\b", linea)
        
        ref_code = None
        if match_c_ref:
            ref_code = match_c_ref.group(1)
        else:
            # Si viene como número simple en la tabla (1 al 30)
            match_num_ref = re.search(r"\b([1-2]?[0-9]|30)\b", linea)
            if match_num_ref and ("HOJALDRADA" in linea.upper() or "BOMBON" in linea.upper() or "SURTIDO" in linea.upper() or "MANTECADO" in linea.upper() or "BARQUILLO" in linea.upper() or "BOLITAS" in linea.upper() or "CHOCO" in linea.upper() or "ALMENDRAS" in linea.upper() or "ROSCOS" in linea.upper() or "YEMITAS" in linea.upper() or "DELICIAS" in linea.upper()):
                num_val = int(match_num_ref.group(1))
                if 1 <= num_val <= 30:
                    ref_code = f"C{num_val:03d}"

        if ref_code:
            # Buscar números en la misma línea que representen cantidades (1 a 50, descartando gramos como 300g, 1800g)
            numeros = re.findall(r"\b\d+\b", linea)
            candidatos = []
            for n in numeros:
                val = int(n)
                # Descartar el número de referencia y valores de gramos/precios
                if val != int(ref_code[1:]) and val <= 50:
                    if not re.search(rf"{val}\s*(?:g|gr|kg|€)", linea, re.IGNORECASE):
                        candidatos.append(val)
            if candidatos:
                cantidades[ref_code] = candidatos[-1]

    return {
        "nombre": nombre if nombre else f"Alumno_{origen_nombre}",
        "telefono": telefono,
        "curso": curso,
        "cantidades": cantidades
    }


def rellenar_excel(excel_bytes, pedidos):
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
    hojas_disponibles = wb.sheetnames
    
    hojas_alumnos_base = [h for h in hojas_disponibles if h.upper() not in ["RESUMEN", "INSTRUCCIONES", "DATOS", "TARIFA"]]
    
    if not hojas_alumnos_base:
        raise ValueError("No se encontró ninguna pestaña de alumno en el Excel.")
        
    ws_plantilla = wb[hojas_alumnos_base[0]]
    alertas = []

    for i, pedido in enumerate(pedidos):
        if i < len(hojas_alumnos_base):
            ws = wb[hojas_alumnos_base[i]]
        else:
            ws = wb.copy_worksheet(ws_plantilla)
            ws.title = f"Alumno_{i+1}"

        # Volcar cabecera
        encontrado_nombre = False
        for row in range(1, 15):
            for col in range(1, 10):
                val = str(ws.cell(row=row, column=col).value or "").upper()
                if "NOMBRE" in val and not encontrado_nombre:
                    celda_dest = ws.cell(row=row, column=col + 1)
                    if "ALUMNO_" in pedido["nombre"] or not pedido["nombre"]:
                        celda_dest.value = "¡REVISAR NOMBRE!"
                        celda_dest.fill = YELLOW_FILL
                        alertas.append(f"Pedido {i+1}: Nombre no detectado.")
                    else:
                        celda_dest.value = pedido["nombre"]
                    encontrado_nombre = True

                if "TELÉFONO" in val or "TELEFONO" in val:
                    if pedido["telefono"]:
                        ws.cell(row=row, column=col + 1, value=pedido["telefono"])

        if not encontrado_nombre:
            ws["B10"] = pedido["nombre"]

        # Mapear referencias en la hoja Excel
        mapa_refs = {}
        for row in range(1, ws.max_row + 1):
            for col in range(1, 10):
                val = str(ws.cell(row=row, column=col).value or "").strip()
                match_c = re.search(r"\b(C\d{3})\b", val)
                if match_c:
                    mapa_refs[match_c.group(1)] = (row, col + 3)

        # Volcar cantidades
        for ref_code, cantidad in pedido["cantidades"].items():
            if ref_code in mapa_refs:
                r, c = mapa_refs[ref_code]
                celda_dest = ws.cell(row=r, column=c)
                celda_dest.value = cantidad

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue(), alertas


# Interfaz Streamlit
st.title("🍬 Procesador Universal de Pedidos (PDFs e Imágenes)")
st.markdown("Sube tanto archivos PDF (con páginas escaneadas) como fotos sueltas de los pedidos en formato imagen.")

col1, col2 = st.columns(2)
with col1:
    archivos_pdf = st.file_uploader("Sube PDF(s) de Pedidos", type=["pdf"], accept_multiple_files=True)
    archivos_img = st.file_uploader("O sube Fotos sueltas de Pedidos", type=["png", "jpg", "jpeg"], accept_multiple_files=True)
with col2:
    excel_file = st.file_uploader("Sube la Plantilla Excel (.xlsx)", type=["xlsx"])

if st.button("🚀 Procesar Todos los Pedidos", type="primary", use_container_width=True):
    if not (archivos_pdf or archivos_img) or not excel_file:
        st.error("❌ Sube al menos un archivo de pedidos (PDF o imágenes) y la plantilla Excel.")
    else:
        with st.spinner("Analizando imágenes y extrayendo datos con OCR..."):
            try:
                pedidos_totales = []
                
                # 1. Procesar PDFs convirtiéndolos a imágenes y aplicando OCR
                if archivos_pdf:
                    for pdf in archivos_pdf:
                        pdf_bytes = pdf.read()
                        images = pdf2image.convert_from_bytes(pdf_bytes)
                        for img_idx, img in enumerate(images):
                            datos_pedido = procesar_contenido_imagen(img, f"PDF_{pdf.name}_pag_{img_idx+1}")
                            pedidos_totales.append(datos_pedido)
                
                # 2. Procesar imágenes sueltas subidas directamente
                if archivos_img:
                    for img_file in archivos_img:
                        img = Image.open(img_file)
                        datos_pedido = procesar_contenido_imagen(img, img_file.name)
                        pedidos_totales.append(datos_pedido)
                
                excel_final, alertas = rellenar_excel(excel_file.read(), pedidos_totales)
                
                st.success(f"✅ ¡Proceso completado con éxito! Se han procesado {len(pedidos_totales)} pedidos en total.")
                
                if alertas:
                    with st.expander("⚠️ Alertas de revisión:", expanded=True):
                        for al in alertas:
                            st.warning(al)
                
                st.download_button(
                    label="📥 Descargar Excel Consolidado",
                    data=excel_final,
                    file_name="pedidos_consolidados_universal.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Error crítico en el procesamiento: {str(e)}")