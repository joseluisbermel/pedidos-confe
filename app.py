import io
import re
import streamlit as st
from pypdf import PdfReader
import openpyxl
from openpyxl.styles import PatternFill

st.set_page_config(
    page_title="Gestor de Pedidos Escolares",
    page_icon="🍬",
    layout="wide"
)

# Estilo de alerta en amarillo chillón para celdas dudosas o anomalías
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

def extraer_datos_pdf(pdf_bytes):
    """
    Lee cada página del PDF, extrae los datos del alumno, busca las referencias 
    exactas y frena de golpe al encontrar 'TOTAL ARTÍCULOS'.
    """
    reader = PdfReader(io.BytesIO(pdf_bytes))
    pedidos = []
    
    for idx, page in enumerate(reader.pages):
        texto = page.extract_text() or ""
        
        # 1. Extraer Nombre y Apellidos de forma robusta
        match_nombre = re.search(r"NOMBRE Y APELLIDOS:\s*(.*?)(?=\s*TELÉFONO|\s*CURSO|\n|$)", texto, re.IGNORECASE)
        nombre = match_nombre.group(1).strip() if match_nombre else ""
        
        if not nombre or len(nombre) < 3:
            lineas_texto = texto.split("\n")
            for l_idx, linea in enumerate(lineas_texto):
                if "NOMBRE Y APELLIDOS" in linea.upper():
                    if l_idx + 1 < len(lineas_texto):
                        nombre = linea.split(":")[-1].strip() + " " + lineas_texto[l_idx + 1].strip()
                    break

        nombre = re.sub(r"TELÉFONO.*", "", nombre, flags=re.IGNORECASE).strip()
        
        # 2. Extraer Teléfono y Curso
        match_tel = re.search(r"TELÉFONO:\s*([\d\s]{6,15})", texto, re.IGNORECASE)
        telefono = match_tel.group(1).strip() if match_tel else ""
        
        match_curso = re.search(r"CURSO/CENTRO:\s*(.*?)(?=\s*FECHA|\n|$)", texto, re.IGNORECASE)
        curso = match_curso.group(1).strip() if match_curso else ""
        
        # 3. Extracción tabulada estricta de referencias hasta "TOTAL ARTÍCULOS"
        cantidades = {}
        lineas = texto.split("\n")
        leyendo_tabla = True
        
        for linea in lineas:
            if "TOTAL" in linea.upper() and "ARTÍCULO" in linea.upper():
                leyendo_tabla = False
                break
                
            if leyendo_tabla:
                match_ref = re.search(r"\b(C\d{3})\b", linea)
                if match_ref:
                    ref_code = match_ref.group(1)
                    
                    # FILTRO INTELIGENTE: Ignorar números que sean gramos (300g, 1800g, etc.) o la referencia
                    # Buscamos números en la línea que sean candidatos lógicos de cantidad (ej: entre 1 y 50)
                    numeros = re.findall(r"\b\d+\b", linea)
                    candidatos = []
                    for n in numeros:
                        val = int(n)
                        # Descartar el número de referencia y números grandes (que son gramos o precios)
                        if val != int(ref_code[1:]) and val <= 50:
                            # Asegurarnos de que no van acompañados de 'g' o 'gr' inmediatamente
                            patron_gramos = rf"{val}\s*(?:g|gr|kg)"
                            if not re.search(patron_gramos, linea, re.IGNORECASE):
                                candidatos.append(val)
                    
                    if candidatos:
                        # La cantidad real suele ser el último número válido de la línea
                        cantidades[ref_code] = candidatos[-1]

        pedidos.append({
            "pagina": idx + 1,
            "nombre": nombre if nombre else f"ALUMNO_PAG_{idx + 1}",
            "telefono": telefono,
            "curso": curso,
            "cantidades": cantidades
        })
        
    return pedidos


def rellenar_excel_dinamico(excel_bytes, pedidos):
    """
    Rellena la plantilla Excel buscando dinámicamente dónde van los datos 
    y duplicando pestañas si hay más alumnos que hojas en la plantilla original.
    """
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
    hojas_disponibles = wb.sheetnames
    
    # Identificar hojas de alumnos (excluyendo RESUMEN, etc.)
    hojas_alumnos_base = [h for h in hojas_disponibles if h.upper() not in ["RESUMEN", "INSTRUCCIONES", "DATOS", "TARIFA"]]
    
    if not hojas_alumnos_base:
        raise ValueError("No se ha encontrado ninguna pestaña de alumno válida en el Excel.")
        
    hoja_plantilla_nombre = hojas_alumnos_base[0]
    ws_plantilla = wb[hoja_plantilla_nombre]
    
    alertas = []

    for i, pedido in enumerate(pedidos):
        # Si hay más alumnos en el PDF que pestañas en el Excel, duplicamos la plantilla
        if i < len(hojas_alumnos_base):
            ws = wb[hojas_alumnos_base[i]]
        else:
            # Duplicar hoja plantilla para el nuevo alumno
            ws = wb.copy_worksheet(ws_plantilla)
            nombre_nueva_hoja = f"Alumno_{i+1}"
            ws.title = nombre_nueva_hoja

        # --- BÚSQUEDA DINÁMICA DE CABECERA EN EL EXCEL ---
        # Buscamos dónde pone "NOMBRE" o "NOMBRE DEL ALUMNO" para colocar el nombre al lado
        celda_nombre_encontrada = False
        celda_telefono_encontrada = False
        
        for row in range(1, 15): # Buscar en las primeras 15 filas
            for col in range(1, 10):
                val = str(ws.cell(row=row, column=col).value or "").upper()
                
                # Buscar celda del nombre
                if "NOMBRE" in val and not celda_nombre_encontrada:
                    celda_val = ws.cell(row=row, column=col + 1) # Celda a la derecha
                    if "ALUMNO" in val or ws.cell(row=row, column=col).value == "NOMBRE DEL ALUMNO:":
                        # A veces está debajo o a la derecha, probamos a la derecha o en col+2
                        celda_destino = ws.cell(row=row, column=col + 1)
                    else:
                        celda_destino = ws.cell(row=row, column=col + 1)
                    
                    if "ALUMNO_PAG" in pedido["nombre"]:
                        celda_destino.value = "¡REVISAR NOMBRE!"
                        celda_destino.fill = YELLOW_FILL
                        alertas.append(f"Página {pedido['pagina']}: Nombre no detectado.")
                    else:
                        celda_destino.value = pedido["nombre"]
                    celda_nombre_encontrada = True

                # Buscar celda del teléfono
                if "TELÉFONO" in val or "TELEFONO" in val:
                    if pedido["telefono"]:
                        ws.cell(row=row, column=col + 1, value=pedido["telefono"])
                    celda_telefono_encontrada = True

        # Fallback por si la etiqueta no se encuentra exactamente
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
                    # La cantidad suele estar unas columnas a la derecha de la referencia
                    mapa_celdas_ref[ref_code] = (row, col + 3)

        # --- VOLCADO Y VALIDACIÓN DE CANTIDADES ---
        for ref_code, cantidad in pedido["cantidades"].items():
            if ref_code in mapa_celdas_ref:
                r, c = mapa_celdas_ref[ref_code]
                celda_destino = ws.cell(row=r, column=c)
                
                if cantidad > 50:
                    celda_destino.value = cantidad
                    celda_destino.fill = YELLOW_FILL
                    alertas.append(f"Página {pedido['pagina']} ({pedido['nombre']}): Cantidad anómala ({cantidad}) en {ref_code}.")
                else:
                    celda_destino.value = cantidad

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue(), alertas


# --- INTERFAZ GRÁFICA STREAMLIT ---
st.title("🍬 Procesador de Pedidos de Confitura")
st.markdown("Extracción inteligente con búsqueda dinámica de campos y filtro de pesos.")

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

if st.button("🚀 Procesar con Precisión Quirúrgica", type="primary", use_container_width=True):
    if not nombre_colegio:
        st.error("❌ Introduce el nombre del colegio.")
    elif not pdf_file or not excel_file:
        st.error("❌ Sube ambos archivos.")
    else:
        with st.spinner("Analizando páginas, filtrando gramos y rellenando pestañas..."):
            try:
                pedidos = extraer_datos_pdf(pdf_file.read())
                excel_procesado, alertas = rellenar_excel_dinamico(excel_file.read(), pedidos)
                
                st.success(f"✅ ¡Lectura finalizada! Se han procesado {len(pedidos)} fichas de alumnos.")
                
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