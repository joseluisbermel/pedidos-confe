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
        
        # Si el nombre está partido en dos líneas, intentamos rescatar la segunda línea si existe
        if not nombre or len(nombre) < 3:
            lineas_texto = texto.split("\n")
            for l_idx, linea in enumerate(lineas_texto):
                if "NOMBRE Y APELLIDOS" in linea.upper():
                    # Unir con la siguiente línea por si está partido
                    if l_idx + 1 < len(lineas_texto):
                        nombre = linea.split(":")[-1].strip() + " " + lineas_texto[l_idx + 1].strip()
                    break

        # Limpiar caracteres raros del nombre
        nombre = re.sub(r"TELÉFONO.*", "", nombre, flags=re.IGNORECASE).strip()
        
        # 2. Extraer Teléfono y Curso
        match_tel = re.search(r"TELÉFONO:\s*([\d\s]{6,15})", texto, re.IGNORECASE)
        telefono = match_tel.group(1).strip() if match_tel else ""
        
        match_curso = re.search(r"CURSO/CENTRO:\s*(.*?)(?=\s*FECHA|\n|$)", texto, re.IGNORECASE)
        curso = match_curso.group(1).strip() if match_curso else ""
        
        # 3. Extracción tabulada estricta de referencias (C001 a C999) hasta "TOTAL ARTÍCULOS"
        cantidades = {}
        lineas = texto.split("\n")
        leyendo_tabla = True
        
        for linea in lineas:
            # Si llegamos al final del bloque de artículos, paramos de leer este alumno
            if "TOTAL" in linea.upper() and "ARTÍCULO" in linea.upper():
                leyendo_tabla = False
                break
                
            if leyendo_tabla:
                # Buscar código de referencia exacto (ej: C001, C019...)
                match_ref = re.search(r"\b(C\d{3})\b", linea)
                if match_ref:
                    ref_code = match_ref.group(1)
                    # Extraer números en la línea
                    numeros = re.findall(r"\b\d+\b", linea)
                    if numeros:
                        # Filtrar para no coger el propio código ni números de la columna de orden (1 al 30)
                        # Buscamos el número que está aislado al final (la cantidad escrita a mano)
                        candidatos = [int(n) for n in numeros if int(n) != int(ref_code[1:])]
                        if candidatos:
                            qty = candidatos[-1] # La cantidad suele ser el último número de la línea
                            cantidades[ref_code] = qty

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
    Rellena la plantilla Excel cruzando las referencias de forma exacta 
    y marcando en amarillo si detecta cantidades anómalas o datos vacíos.
    """
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
    hojas_disponibles = wb.sheetnames
    
    # Excluir pestañas de resumen o instrucciones
    hojas_alumnos = [h for h in hojas_disponibles if h.upper() not in ["RESUMEN", "INSTRUCCIONES", "DATOS", "TARIFA"]]
    
    alertas = []

    for i, pedido in enumerate(pedidos):
        if i >= len(hojas_alumnos):
            alertas.append(f"⚠️ Hay más páginas en el PDF que pestañas de alumnos en el Excel ({len(pedidos)} vs {len(hojas_alumnos)}).")
            break
            
        ws = wb[hojas_alumnos[i]]
        
        # --- VOLCADO DE CABECERA ---
        celda_nombre = ws["B3"]
        if "ALUMNO_PAG" in pedido["nombre"]:
            celda_nombre.value = "¡REVISAR NOMBRE!"
            celda_nombre.fill = YELLOW_FILL
            alertas.append(f"Página {pedido['pagina']}: Nombre no detectado. Revisar pestaña '{ws.title}'.")
        else:
            celda_nombre.value = pedido["nombre"]
            
        if pedido["telefono"]:
            ws["B4"] = pedido["telefono"]
        if pedido["curso"]:
            ws["B5"] = pedido["curso"]
            
        # --- MAPEAR LA TABLA DE LA PLANTILLA EXCEL ---
        # Buscamos exactamente dónde están las celdas de las referencias (C001, C002...) en esta hoja
        mapa_celdas_ref = {}
        for row in range(1, ws.max_row + 1):
            for col in range(1, 10):
                val = str(ws.cell(row=row, column=col).value or "").strip()
                match_c = re.search(r"\b(C\d{3})\b", val)
                if match_c:
                    ref_code = match_c.group(1)
                    # Asumimos que la cantidad se pone en la columna situada a la derecha de la referencia o descripción
                    # Buscamos celdas vacías o numéricas en las columnas contiguas (ej: col+3 o col+4)
                    mapa_celdas_ref[ref_code] = (row, col + 3) # Ajusta este desplazamiento si difiere en tu plantilla

        # --- VOLCADO Y VALIDACIÓN DE CANTIDADES ---
        for ref_code, cantidad in pedido["cantidades"].items():
            if ref_code in mapa_celdas_ref:
                r, c = mapa_celdas_ref[ref_code]
                celda_destino = ws.cell(row=r, column=c)
                
                # VALIDACIÓN DE ANOMALÍAS: Si la cantidad es exagerada (ej > 50), se marca en amarillo
                if cantidad > 50:
                    celda_destino.value = cantidad
                    celda_destino.fill = YELLOW_FILL
                    alertas.append(f"Página {pedido['pagina']} ({pedido['nombre']}): Cantidad anómala ({cantidad}) en referencia {ref_code}. Marcado en amarillo.")
                else:
                    celda_destino.value = cantidad

    # Guardar cambios
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue(), alertas


# --- INTERFAZ GRÁFICA STREAMLIT ---
st.title("🍬 Procesador de Pedidos de Confitura (Versión Optimizada)")
st.markdown("Extracción inteligente delimitada por 'TOTAL ARTÍCULOS' y validación de referencias.")

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

if st.button("🚀 Procesar con Precisión", type="primary", use_container_width=True):
    if not nombre_colegio:
        st.error("❌ Introduce el nombre del colegio.")
    elif not pdf_file or not excel_file:
        st.error("❌ Sube ambos archivos.")
    else:
        with st.spinner("Analizando bloques y frenando en 'TOTAL ARTÍCULOS'..."):
            try:
                pedidos = extraer_datos_pdf(pdf_file.read())
                excel_procesado, alertas = rellenar_excel_dinamico(excel_file.read(), pedidos)
                
                st.success(f"✅ ¡Lectura finalizada! Se han procesado {len(pedidos)} fichas.")
                
                if alertas:
                    with st.expander("⚠️ Alertas de revisión (Celdas marcadas en amarillo):", expanded=True):
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