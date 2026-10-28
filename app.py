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

# Relleno amarillo para destacar campos incompletos o celdas no encontradas
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

def extraer_datos_pdf(pdf_bytes):
    """
    Lee cada página del PDF y extrae la información del alumno y las cantidades asignadas.
    Retorna una lista de diccionarios, uno por cada página/alumno.
    """
    reader = PdfReader(io.BytesIO(pdf_bytes))
    pedidos = []
    
    for idx, page in enumerate(reader.pages):
        texto = page.extract_text() or ""
        
        # 1. Extraer Nombre y Apellidos
        match_nombre = re.search(r"NOMBRE Y APELLIDOS:\s*(.*?)(?=\s*TELÉFONO|\s*CURSO|\n|$)", texto, re.IGNORECASE)
        nombre = match_nombre.group(1).strip() if match_nombre else ""
        
        # 2. Extraer Teléfono
        match_tel = re.search(r"TELÉFONO:\s*([\d\s]{6,15})", texto, re.IGNORECASE)
        telefono = match_tel.group(1).strip() if match_tel else ""
        
        # 3. Extraer Curso / Centro
        match_curso = re.search(r"CURSO/CENTRO:\s*(.*?)(?=\s*FECHA|\n|$)", texto, re.IGNORECASE)
        curso = match_curso.group(1).strip() if match_curso else ""
        
        # 4. Extraer Referencias C001-C099 y sus Cantidades
        # Busca patrones tipo: C001 ... (número o marca al final)
        cantidades = {}
        lineas = texto.split("\n")
        for linea in lineas:
            # Detectar si la línea contiene un código de referencia (ej: C001, C012, etc.)
            match_ref = re.search(r"\b(C\d{3})\b", linea)
            if match_ref:
                ref_code = match_ref.group(1)
                # Extraer el último entero de la línea si existe
                numeros = re.findall(r"\b\d+\b", linea)
                if numeros:
                    # El último número suele ser la cantidad
                    qty = int(numeros[-1])
                    # Evitar coger la propia referencia (ej: de C001 no coger el 1)
                    if len(numeros) > 1 or str(qty) != ref_code[1:].lstrip("0"):
                        cantidades[ref_code] = qty
                    else:
                        cantidades[ref_code] = qty
                elif "X" in linea.upper() or "✓" in linea:
                    cantidades[ref_code] = 1

        pedidos.append({
            "pagina": idx + 1,
            "nombre": nombre,
            "telefono": telefono,
            "curso": curso,
            "cantidades": cantidades
        })
        
    return pedidos


def rellenar_excel_dinamico(excel_bytes, pedidos):
    """
    Rellena la plantilla Excel respetando la estructura dinámica de hojas y celdas.
    Aplica marcado amarillo si faltan datos clave (como el nombre).
    """
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
    hojas_disponibles = wb.sheetnames
    
    # Filtrar pestañas que sean para alumnos (excluyendo RESUMEN, Plantilla base, etc.)
    hojas_alumnos = [h for h in hojas_disponibles if h.upper() not in ["RESUMEN", "INSTRUCCIONES", "DATOS"]]
    
    alertas = []

    for i, pedido in enumerate(pedidos):
        if i >= len(hojas_alumnos):
            alertas.append(f"⚠️ El PDF tiene más de {len(hojas_alumnos)} alumnos ({len(pedidos)} en total). Las páginas sobrantes no se pudieron volcar.")
            break
            
        ws = wb[hojas_alumnos[i]]
        
        # --- VOLCADO DE CABECERA ---
        # Buscamos dinámicamente celdas que contengan los textos de cabecera o usamos celdas estándar
        # 1. Nombre del Alumno
        if pedido["nombre"]:
            # Intenta ubicar la celda donde va el nombre
            ws["B3"] = pedido["nombre"]
        else:
            ws["B3"] = "¡REVISAR NOMBRE!"
            ws["B3"].fill = YELLOW_FILL
            alertas.append(f"Página {pedido['pagina']}: Nombre no detectado. Se ha marcado en amarillo en la pestaña '{ws.title}'.")
            
        # 2. Teléfono y Curso
        if pedido["telefono"]:
            ws["B4"] = pedido["telefono"]
        if pedido["curso"]:
            ws["B5"] = pedido["curso"]
            
        # --- VOLCADO DINÁMICO DE PRODUCTOS/CANTIDADES ---
        # Recorremos la hoja buscando las referencias (ej: en la columna A o B)
        mapa_celdas_ref = {}
        for row in range(1, ws.max_row + 1):
            for col in range(1, 10): # Inspeccionar las primeras columnas
                valor = str(ws.cell(row=row, column=col).value or "").strip()
                match_c = re.search(r"\b(C\d{3})\b", valor)
                if match_c:
                    ref_code = match_c.group(1)
                    # Suponemos que la columna de cantidad es la siguiente o la columna E/F
                    # Colocamos la cantidad 3 columnas a la derecha por defecto si no se detecta otra
                    mapa_celdas_ref[ref_code] = (row, col + 3)

        # Rellenar las cantidades encontradas en el PDF
        for ref_code, cantidad in pedido["cantidades"].items():
            if ref_code in mapa_celdas_ref:
                r, c = mapa_celdas_ref[ref_code]
                ws.cell(row=r, column=c, value=cantidad)

    # Guardar en memoria
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue(), alertas


# --- INTERFAZ GRÁFICA STREAMLIT ---
st.title("🍬 Procesador de Pedidos de Confitura")
st.markdown("Herramienta para la extracción de pedidos en PDF y volcado dinámico en plantillas Excel.")

st.sidebar.header("Configuración del Pedido")
nombre_colegio = st.sidebar.text_input("Nombre del Colegio / Centro", placeholder="Ej: IES_Lopez_de_Arenas")
fecha_pedido = st.sidebar.date_input("Fecha del Pedido")

st.divider()

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. PDF de Pedidos")
    pdf_file = st.file_uploader("Sube el PDF escaneado con las hojas de los alumnos", type=["pdf"])

with col2:
    st.subheader("2. Plantilla Excel")
    excel_file = st.file_uploader("Sube la plantilla Excel de la tarifa limpia (.xlsx)", type=["xlsx"])

st.divider()

if st.button("🚀 Procesar Pedido y Generar Excel", type="primary", use_container_width=True):
    if not nombre_colegio:
        st.error("❌ Por favor, introduce el nombre del colegio en la barra lateral.")
    elif not pdf_file or not excel_file:
        st.error("❌ Debes subir tanto el PDF de los pedidos como la plantilla Excel.")
    else:
        with st.spinner("Leyendo páginas del PDF y volcando datos en el Excel..."):
            try:
                # 1. Procesar el PDF
                pedidos = extraer_datos_pdf(pdf_file.read())
                
                # 2. Volcar al Excel
                excel_procesado, alertas = rellenar_excel_dinamico(excel_file.read(), pedidos)
                
                st.success(f"✅ ¡Proceso completado! Se han procesado {len(pedidos)} hojas/alumnos.")
                
                # Mostrar avisos si hubo celdas no encontradas o nombres vacíos
                if alertas:
                    with st.expander("⚠️ Revisar advertencias (Celdas marcadas en amarillo):", expanded=True):
                        for al in alertas:
                            st.warning(al)
                
                # 3. Formatear nombre del archivo final
                colegio_clean = re.sub(r"[^\w\s-]", "", nombre_colegio).strip().replace(" ", "_")
                fecha_str = fecha_pedido.strftime("%d-%m-%Y")
                nombre_archivo_final = f"{colegio_clean}_{fecha_str}.xlsx"
                
                # 4. Botón de Descarga
                st.download_button(
                    label=f"📥 Descargar {nombre_archivo_final}",
                    data=excel_procesado,
                    file_name=nombre_archivo_final,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
            except Exception as e:
                st.error(f"Ocurrió un error al procesar el archivo: {str(e)}")