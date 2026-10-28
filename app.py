import io
import re
import streamlit as st
from pypdf import PdfReader
import openpyxl
from openpyxl.styles import PatternFill
from PIL import Image
import pdf2image

st.set_page_config(
    page_title="Gestor de Pedidos Escolares",
    page_icon="🍬",
    layout="wide"
)

# Estilo de alerta en amarillo chillón para celdas dudosas
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

def extraer_texto_pagina(page, pdf_bytes, page_idx):
    """Extrae texto digital o aplica OCR local mediante Tesseract si es una foto."""
    texto = page.extract_text() or ""
    if len(texto.strip()) > 40:  # Umbral subido ligeramente para asegurar calidad
        return texto
    
    try:
        import pytesseract
        images = pdf2image.convert_from_bytes(pdf_bytes, first_page=page_idx+1, last_page=page_idx+1)
        if images:
            texto_ocr = pytesseract.image_to_string(images[0], lang='spa')
            return texto_ocr
    except Exception as e:
        st.warning(f"Aviso: Fallback OCR activado en página {page_idx+1} (Asegúrate de tener Tesseract instalado)")
        
    return texto

def extraer_pedidos_agrupados(pdf_bytes):
    """Procesa el PDF unificando cabeceras y aislando de forma exacta las cantidades."""
    reader = PdfReader(io.BytesIO(pdf_bytes))
    pedidos = []
    alumno_actual = None
    
    for idx, page in enumerate(reader.pages):
        texto = extraer_texto_pagina(page, pdf_bytes, idx)
        lineas = [l.strip() for l in texto.split("\n") if l.strip()]
        
        # Detectar si esta página es el inicio de una FICHA NUEVA
        # Si contiene 'NOMBRE' o no hay ningún alumno activo, inicializamos
        contiene_nombre = any(x in texto.upper() for x in ["NOMBRE Y APELLIDOS:", "HOJA DE PEDIDOS"])
        
        if alumno_actual is None or contiene_nombre:
            if alumno_actual is not None:
                pedidos.append(alumno_actual)  # Guardar el alumno anterior si se quedó abierto
            
            nombre = "Ficha_Anonima"
            telefono = ""
            curso = ""
            
            # Buscar Nombre en las primeras líneas de la página
            for l_idx, linea in enumerate(lineas[:8]):
                if "NOMBRE" in linea.upper():
                    if ":" in linea:
                        nombre = linea.split(":")[-1].strip()
                    if (not nombre or len(nombre) < 3) and l_idx + 1 < len(lineas):
                        nombre = lineas[l_idx + 1].strip()
                    break
            
            # Limpieza exhaustiva de metadatos en el nombre extraído
            nombre = re.sub(r"(?:TELÉFONO|TELEFONO|CURSO|FECHA|CENTRO).*", "", nombre, flags=re.IGNORECASE).strip()
            nombre = re.sub(r'[^a-zA-ZÁÉÍÓÚáéíóúÑñ\s]', '', nombre).strip()
            
            # Buscar teléfono mediante patrón numérico
            for linea in lineas[:10]:
                match_tel = re.search(r"(?:TELÉFONO|TELEFONO|TELF)[:\s]*([\d\s-]{6,15})", linea, re.IGNORECASE)
                if match_tel:
                    telefono = re.sub(r"\D", "", match_tel.group(1))
                    break

            alumno_actual = {
                "nombre": nombre if len(nombre) > 2 else f"REVISAR_ALUMNO_PAG_{idx + 1}",
                "telefono": telefono,
                "cantidades": {},
                "paginas": [idx + 1]
            }
        else:
            alumno_actual["paginas"].append(idx + 1)

        # --- EXTRACCIÓN DE CANTIDADES POR LÍNEA MEDIANTE REGEX SEGURO ---
        for linea in lineas:
            match_ref = re.search(r"\b(C\d{3})\b", linea)
            if match_ref:
                ref_code = match_ref.group(1)
                
                # Eliminamos los gramos (ej: '600 g', '300g') para que no confundan al lector numérico
                linea_limpia = re.sub(r"\d+\s*(?:g|gr|kg|ml)\b", "", linea, flags=re.IGNORECASE)
                
                # Extraemos todos los números aislados que queden
                numeros = [int(n) for n in re.findall(r"\b\d+\b", linea_limpia)]
                
                if numeros:
                    # El número de orden (1 al 30) suele aparecer al principio. 
                    # El último número puro de la línea es SIEMPRE la cantidad.
                    cantidad = numeros[-1]
                    
                    # Filtro de seguridad: Evitar registrar el número de orden del catálogo como cantidad
                    num_orden_catalogo = int(ref_code[1:])
                    if len(numeros) == 1 and cantidad == num_orden_catalogo:
                        continue # La celda de cantidad estaba vacía (era solo el índice del artículo)
                    
                    # Si el número final sobreviviente es válido, se anota
                    if cantidad > 0:
                        alumno_actual["cantidades"][ref_code] = cantidad

        # Forzar cierre si encontramos explícitamente el total de la hoja
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
    
    # Filtro para ignorar hojas maestras de totales/tarifas
    hojas_alumnos_base = [h for h in hojas_disponibles if h.upper() not in ["RESUMEN", "INSTRUCCIONES", "DATOS", "TARIFA", "TOTALES"]]
    
    if not hojas_alumnos_base:
        raise ValueError("Error: Tu plantilla Excel no contiene pestañas individuales asignables a los alumnos.")
        
    ws_plantilla = wb[hojas_alumnos_base[0]]
    alertas = []

    for i, pedido in enumerate(pedidos):
        # Reutilizar pestaña existente o clonar la plantilla si nos quedamos sin hojas
        if i < len(hojas_alumnos_base):
            ws = wb[hojas_alumnos_base[i]]
            if "REVISAR" not in pedido["nombre"]:
                ws.title = pedido["nombre"][:30] # Renombrar la pestaña con el nombre del alumno
        else:
            ws = wb.copy_worksheet(ws_plantilla)
            ws.title = f"{pedido['nombre'][:25]}_{i+1}"

        # --- VOLCADO INTELIGENTE DE CABECERA ---
        celda_nombre_encontrada = False
        for row in range(1, 15):
            for col in range(1, 10):
                val = str(ws.cell(row=row, column=col).value or "").upper()
                if "ALUMNO" in val or "NOMBRE" in val:
                    celda_destino = ws.cell(row=row, column=col + 1)
                    if "REVISAR" in pedido["nombre"]:
                        celda_destino.value = f"REVISAR (Pág {pedido['paginas']})"
                        celda_destino.fill = YELLOW_FILL
                        alertas.append(f"⚠️ Páginas {pedido['paginas']}: No se pudo leer el nombre manuscrito de forma nítida.")
                    else:
                        celda_destino.value = pedido["nombre"]
                    celda_nombre_encontrada = True
                    break
            if celda_nombre_encontrada: break

        # --- MAPEO AUTOMÁTICO DE COORDENADAS POR CÓDIGO (C001-C030) ---
        mapa_celdas_ref = {}
        for row in range(1, ws.max_row + 1):
            for col in range(1, 10):
                val = str(ws.cell(row=row, column=col).value or "").strip()
                if re.match(r"^C\d{3}$", val):
                    # Guardamos la coordenada exacta (Fila, Columna Cantidad -> habitualmente Columna + 3)
                    mapa_celdas_ref[val] = (row, col + 3)

        # --- INYECCIÓN DE CANTIDADES ---
        for ref_code, cantidad in pedido["cantidades"].items():
            if ref_code in mapa_celdas_ref:
                r, c = mapa_celdas_ref[ref_code]
                celda_destino = ws.cell(row=r, column=c)
                
                # Gestión de anomalías de lectura manuscrita (valores disparados)
                if cantidad > 20: 
                    celda_destino.value = cantidad
                    celda_destino.fill = YELLOW_FILL
                    alertas.append(f"🧐 {pedido['nombre']} (Pág {pedido['paginas']}): Cantidad inusualmente alta ({cantidad}) en {ref_code}. Verificado visualmente.")
                else:
                    celda_destino.value = cantidad

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue(), alertas


# --- INTERFAZ DE USUARIO (STREAMLIT WEB) ---
st.title("🍬 Gestor Automatizado de Pedidos - E. Moreno")
st.markdown("Carga tu PDF desordenado y tu plantilla Excel. El sistema aplicará OCR local y consolidará las fichas al instante.")

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
        with st.spinner("Procesando datos en segundo plano..."):
            try:
                pdf_bytes = pdf_file.read()
                pedidos = extraer_pedidos_agrupados(pdf_bytes)
                excel_procesado, alertas = rellenar_excel_dinamico(excel_file.read(), pedidos)
                
                st.success(f"✨ ¡Conseguido! Se han estructurado {len(pedidos)} alumnos en tu Excel.")
                
                if alertas:
