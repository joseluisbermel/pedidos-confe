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

# Estilo de alerta en amarillo chillón para celdas dudosas o anomalías
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

def extraer_texto_pagina(page, pdf_bytes, page_idx):
    """
    Intenta extraer texto de forma digital. Si está vacío (es una imagen escaneada),
    aplica OCR localmente mediante pytesseract y pdf2image.
    """
    texto = page.extract_text() or ""
    if len(texto.strip()) > 15:
        return texto
    
    # Fallback a OCR si la página es una imagen escaneada
    try:
        import pytesseract
        images = pdf2image.convert_from_bytes(pdf_bytes, first_page=page_idx+1, last_page=page_idx+1)
        if images:
            texto_ocr = pytesseract.image_to_string(images[0], lang='spa')
            return texto_ocr
    except Exception as e:
        print(f"Error en OCR para página {page_idx+1}: {e}")
        
    return texto


def extraer_pedidos_agrupados(pdf_bytes):
    """
    Agrupa las páginas del PDF por alumno de forma inteligente.
    Una ficha puede ocupar 1 o 2 páginas y finaliza al encontrar 'TOTAL ARTÍCULOS'.
    """
    reader = PdfReader(io.BytesIO(pdf_bytes))
    pedidos = []
    
    alumno_actual = None
    
    for idx, page in enumerate(reader.pages):
        texto = extraer_texto_pagina(page, pdf_bytes, idx)
        lineas = [l.strip() for l in texto.split("\n") if l.strip()]
        
        # Si no hay un alumno activo, buscamos los datos de cabecera en esta página
        if alumno_actual is None:
            nombre = ""
            telefono = ""
            curso = ""
            
            # Buscar nombre
            for l_idx, linea in enumerate(lineas):
                if "NOMBRE" in linea.upper() or "APELLIDOS" in linea.upper():
                    if ":" in linea:
                        nombre = linea.split(":")[-1].strip()
                    if not nombre and l_idx + 1 < len(lineas):
                        nombre = lineas[l_idx + 1].strip()
                    break
            
            # Si no se encontró por etiqueta, buscamos patrones habituales
            if not nombre or len(nombre) < 3:
                for linea in lineas:
                    if len(linea) > 5 and not "EMORENO" in linea.upper() and not "PEDIDOS" in linea.upper() and not "ESCOLAR" in linea.upper():
                        # Podría ser el nombre si va en mayúsculas
                        if linea == linea.upper() and " " in linea:
                            nombre = linea
                            break

            nombre = re.sub(r"TELÉFONO.*", "", nombre, flags=re.IGNORECASE).strip()
            nombre = re.sub(r"CURSO.*", "", nombre, flags=re.IGNORECASE).strip()

            # Buscar teléfono
            for linea in lineas:
                match_tel = re.search(r"(?:TELÉFONO|TELEFONO|TELF)[:\s]*([\d\s]{6,15})", linea, re.IGNORECASE)
                if match_tel:
                    telefono = match_tel.group(1).strip()
                    break

            # Buscar curso / centro
            for linea in lineas:
                match_cur = re.search(r"(?:CURSO/CENTRO|CURSO|CENTRO)[:\s]*(.*)", linea, re.IGNORECASE)
                if match_cur:
                    curso = match_cur.group(1).strip()
                    break

            alumno_actual = {
                "nombre": nombre if nombre else f"ALUMNO_PAG_{idx + 1}",
                "telefono": telefono,
                "curso": curso,
                "cantidades": {},
                "paginas": [idx + 1]
            }
        else:
            alumno_actual["paginas"].append(idx + 1)

        # Extracción de referencias (C001 a C030) y cantidades
        leyendo_tabla = True
        for linea in lineas:
            if "TOTAL" in linea.upper() and "ARTÍCULO" in linea.upper():
                leyendo_tabla = False
                break
                
            if leyendo_tabla:
                match_ref = re.search(r"\b(C\d{3})\b", linea)
                if match_ref:
                    ref_code = match_ref.group(1)
                    numeros = re.findall(r"\b\d+\b", linea)
                    candidatos = []
                    for n in numeros:
                        val = int(n)
                        # Descartar número de referencia y gramos (300g, 1800g, etc.)
                        if val != int(ref_code[1:]) and val <= 50:
                            if not re.search(rf"{val}\s*(?:g|gr|kg)", linea, re.IGNORECASE):
                                candidatos.append(val)
                    if candidatos:
                        alumno_actual["cantidades"][ref_code] = candidatos[-1]

        # Comprobar si esta página contiene el cierre del pedido ("TOTAL ARTÍCULOS")
        es_fin_pedido = any("TOTAL" in l.upper() and "ARTÍCULO" in l.upper() for l in lineas)
        if es_fin_pedido:
            pedidos.append(alumno_actual)
            alumno_actual = None

    # Si queda un alumno pendiente sin cierre explícito, se guarda
    if alumno_actual is not None:
        pedidos.append(alumno_actual)
        
    return pedidos


def rellenar_excel_dinamico(excel_bytes, pedidos):
    """
    Rellena la plantilla Excel con los pedidos agrupados, duplicando pestañas 
    automáticamente si hay más alumnos que hojas en la plantilla original.
    """
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
    hojas_disponibles = wb.sheetnames
    
    # Excluir pestañas de sistema o resumen
    hojas_alumnos_base = [h for h in hojas_disponibles if h.upper() not in ["RESUMEN", "INSTRUCCIONES", "DATOS", "TARIFA"]]
    
    if not hojas_alumnos_base:
        raise ValueError("No se ha encontrado ninguna pestaña de alumno válida en el Excel.")
        
    hoja_plantilla_nombre = hojas_alumnos_base[0]
    ws_plantilla = wb[hoja_plantilla_nombre]
    
    alertas = []

    for i, pedido in enumerate(pedidos):
        # Asignar o duplicar pestaña para el alumno actual
        if i < len(hojas_alumnos_base):
            ws = wb[hojas_alumnos_base[i]]
        else:
            ws = wb.copy_worksheet(ws_plantilla)
            ws.title = f"Alumno_{i+1}"

        # --- VOLCADO DINÁMICO DE CABECERA ---
        celda_nombre_encontrada = False
        for row in range(1, 15):
            for col in range(1, 10):
                val = str(ws.cell(row=row, column=col).value or "").upper()
                
                if "NOMBRE" in val and not celda_nombre_encontrada:
                    celda_destino = ws.cell(row=row, column=col + 1)
                    if "ALUMNO_PAG" in pedido["nombre"] or not pedido["nombre"]:
                        celda_destino.value = "¡REVISAR NOMBRE!"
                        celda_destino.fill = YELLOW_FILL
                        alertas.append(f"Páginas {pedido['paginas']}: Nombre no detectado.")
                    else:
                        celda_destino.value = pedido["nombre"]
                    celda_nombre_encontrada = True

                if "TELÉFONO" in val or "TELEFONO" in val:
                    if pedido["telefono"]:
                        ws.cell(row=row, column=col + 1, value=pedido["telefono"])

        if not celda_nombre_encontrada:
            ws["B10"] = pedido["nombre"] if "ALUMNO_PAG" not in pedido["nombre"] else "¡REVISAR NOMBRE!"

        # --- MAPEAR LA TABLA DE REFERENCIAS EN EL EXCEL ---
        mapa_celdas_ref = {}
        for row in range(1, ws.max_row + 1):
            for col in range(1, 10):
                val = str(ws.cell(row=row, column=col).value or "").strip()
                match_c = re.search(r"\b(C\d{3})\b", val)
                if match_c:
                    ref_code = match_c.group(1)
                    mapa_celdas_ref[ref_code] = (row, col + 3)

        # --- VOLCADO Y VALIDACIÓN DE CANTIDADES ---
        for ref_code, cantidad in pedido["cantidades"].items():
            if ref_code in mapa_celdas_ref:
                r, c = mapa_celdas_ref[ref_code]
                celda_destino = ws.cell(row=r, column=c)
                
                if cantidad > 50:
                    celda_destino.value = cantidad
                    celda_destino.fill = YELLOW_FILL
                    alertas.append(f"Alumno {pedido['nombre']} (Págs {pedido['paginas']}): Cantidad anómala ({cantidad}) en {ref_code}.")
                else:
                    celda_destino.value = cantidad

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue(), alertas


# --- INTERFAZ GRÁFICA STREAMLIT ---
st.title("🍬 Procesador de Pedidos de Confitura")
st.markdown("Agrupación inteligente multi-página y soporte OCR para hojas escritas a mano.")

st.sidebar.header("Configuración")
nombre_colegio = st.sidebar.text_input("Nombre del Colegio / Centro", placeholder="Ej: IES_Lopez_de_Arenas")
fecha_pedido = st.sidebar.date_input("Fecha del Pedido")

st.divider()

col1, col2 = st.columns(2)
with col1:
    st.subheader("1. PDF de Pedidos")
    pdf_file = st.file_uploader("Sube el PDF", type=["pdf"])
with col2:
    st.subheader("2. Plantilla Excel")
    excel_file = st.file_uploader("Sube la plantilla Excel (.xlsx)", type=["xlsx"])

st.divider()

if st.button("🚀 Procesar Pedidos al Completo", type="primary", use_container_width=True):
    if not nombre_colegio:
        st.error("❌ Introduce el nombre del colegio.")
    elif not pdf_file or not excel_file:
        st.error("❌ Sube ambos archivos.")
    else:
        with st.spinner("Agrupando páginas, aplicando OCR y volcando en el Excel..."):
            try:
                pdf_bytes = pdf_file.read()
                pedidos = extraer_pedidos_agrupados(pdf_bytes)
                excel_procesado, alertas = rellenar_excel_dinamico(excel_file.read(), pedidos)
                
                st.success(f"✅ ¡Proceso completado! Se han consolidado {len(pedidos)} fichas de alumnos.")
                
                if alertas:
                    with st.expander("⚠️ Alertas de revisión:", expanded=True):
                        for al in alertas:
                            st.warning(al)
                
                colegio_clean = re.sub(r"[^\w\s-]", "", nombre_colegio).strip().replace(" ", "_")
                fecha_str = fecha_pedido.strftime("%d-%m-%Y")
                nombre_archivo_final = f"{colegio_clean}_{fecha_str}.xlsx"
                
                st.download_button(
                    label=f"📥 Descargar {nombre_archivo_final}",
                    data=excel_procesado,
                    file_name=nombre_archivo_final,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Error crítico en el proceso: {str(e)}")