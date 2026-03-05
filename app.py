import streamlit as st
import requests
import pandas as pd
import html
from io import BytesIO
import datetime
import docx
from docx import Document

# ==========================================
# CONFIGURACIÓN INICIAL Y ESTILOS
# ==========================================
st.set_page_config(page_title="Boletín Mensual - Banxico", layout="wide")

# ==========================================
# FUNCIONES AUXILIARES (BACKEND)
# ==========================================

@st.cache_data(show_spinner="Descargando y procesando datos del BIS...")
def load_data_bis():
    url = "https://www.bis.org/api/document_lists/cbspeeches.json"
    response = requests.get(url)
    response.raise_for_status()
    data = response.json()

    speeches_dict = data.get("list", {})
    rows = []

    for path, speech in speeches_dict.items():
        title = html.unescape(speech.get("short_title", ""))
        date_str = speech.get("publication_start_date", "")

        link = "https://www.bis.org" + path + (".htm" if not path.endswith(".htm") else "")

        rows.append({
            "Date": date_str,
            "Title": title,
            "Link": link
        })

    df = pd.DataFrame(rows)
    if not df.empty:
        df["Date"] = pd.to_datetime(df["Date"])
        df = df.sort_values("Date", ascending=False)
    
    return df

def add_hyperlink(paragraph, text, url):
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    hyperlink = docx.oxml.shared.OxmlElement('w:hyperlink')
    hyperlink.set(docx.oxml.shared.qn('r:id'), r_id)

    new_run = docx.oxml.shared.OxmlElement('w:r')
    rPr = docx.oxml.shared.OxmlElement('w:rPr')

    c = docx.oxml.shared.OxmlElement('w:color')
    c.set(docx.oxml.shared.qn('w:val'), '0000EE')
    rPr.append(c)

    u = docx.oxml.shared.OxmlElement('w:u')
    u.set(docx.oxml.shared.qn('w:val'), 'single')
    rPr.append(u)

    t = docx.oxml.shared.OxmlElement('w:t')
    t.text = text
    new_run.append(rPr)
    new_run.append(t)
    
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink

def generate_word(dataframe, title="Discursos"):
    doc = Document()
    doc.add_heading(title, 0)

    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Date'
    hdr_cells[1].text = 'Title'

    for index, row in dataframe.iterrows():
        row_cells = table.add_row().cells
        
        if pd.api.types.is_datetime64_any_dtype(row['Date']):
            date_str = row['Date'].strftime('%Y-%m-%d')
        else:
            date_str = str(row['Date'])
        row_cells[0].text = date_str
        
        p = row_cells[1].paragraphs[0]
        add_hyperlink(p, str(row['Title']), str(row['Link']))

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output


# ==========================================
# INTERFAZ DE USUARIO (SIDEBAR Y NAVEGACIÓN)
# ==========================================

# 1. Logo institucional en la barra lateral
# NOTA: Sube un archivo llamado 'logo_banxico.png' a la misma carpeta de tu app en GitHub
try:
    st.sidebar.image("logo_banxico.png", use_column_width=True)
except:
    st.sidebar.markdown("### 🏦 BANCO DE MÉXICO") # Placeholder si no encuentra la imagen

st.sidebar.markdown("---")
st.sidebar.header("Menú de Navegación")

# 2. Selector de Tipo de Documento
tipo_doc = st.sidebar.selectbox(
    "Selecciona el Tipo de Documento",
    ["Reportes", "Publicaciones Institucionales", "Investigación", "Discursos"]
)

# 3. Selector de Organismo (depende del tipo de documento)
if tipo_doc == "Discursos":
    organismos_discursos = ["BIS", "FMI", "BCE", "Fed"] # Puedes agregar más
    organismo_seleccionado = st.sidebar.selectbox("Selecciona el Organismo", organismos_discursos)
else:
    # Placeholder para las otras categorías
    organismos_generales = ["BM", "BID", "CEF", "FEM", "FMI", "BPI", "OCDE", "CEMLA"]
    organismo_seleccionado = st.sidebar.selectbox("Selecciona el Organismo", organismos_generales)

st.sidebar.markdown("---")
st.sidebar.info("Herramienta de extracción automatizada para la elaboración del boletín mensual.")


# ==========================================
# CONTENIDO PRINCIPAL (MAIN PAGE)
# ==========================================

# Portada Institucional
st.title("Boletín Mensual de Organismos Internacionales")
st.markdown(f"**Explorador de {tipo_doc} - {organismo_seleccionado}**")
st.markdown("---")

# ==========================================
# MÓDULOS DE EXTRACCIÓN
# ==========================================

# MÓDULO: DISCURSOS -> BIS
# MÓDULO: DISCURSOS -> BIS
if tipo_doc == "Discursos" and organismo_seleccionado == "BIS":
    
    st.subheader("1. Selecciona el Mes y Año")

    # Diccionario para mapear el nombre del mes con su número en fechas
    meses_dict = {
        "Enero": 1, "Febrero": 2, "Marzo": 3, "Abril": 4,
        "Mayo": 5, "Junio": 6, "Julio": 7, "Agosto": 8,
        "Septiembre": 9, "Octubre": 10, "Noviembre": 11, "Diciembre": 12
    }

    col1, col2 = st.columns(2)
    with col1:
        # Selector múltiple de meses
        meses_seleccionados = st.multiselect(
            "Mes(es)",
            options=list(meses_dict.keys()),
            default=["Marzo"] # Puedes dejarlo vacío usando default=[]
        )
    with col2:
        # Selector múltiple de años con 2026 por defecto
        anios_seleccionados = st.multiselect(
            "Año(s)",
            options=["2023", "2024", "2025", "2026", "2027"],
            default=["2026"]
        )

    buscar = st.button("🔍 Buscar Discursos del BIS", type="primary")

    # Lógica de ejecución
    if buscar or "bis_df_filtrado" in st.session_state:
        
        # Validación de que al menos haya un mes y un año seleccionado
        if not meses_seleccionados or not anios_seleccionados:
            st.warning("⚠️ Por favor, selecciona al menos un mes y un año para realizar la búsqueda.")
        else:
            df = load_data_bis()
            
            # Convertimos las selecciones a números para filtrar
            meses_num = [meses_dict[m] for m in meses_seleccionados]
            anios_num = [int(a) for a in anios_seleccionados]
            
            # Filtramos extrayendo el mes y año de la columna 'Date'
            mask = (df["Date"].dt.year.isin(anios_num)) & (df["Date"].dt.month.isin(meses_num))
            filtered_df = df[mask]
            
            st.session_state["bis_df_filtrado"] = filtered_df

            st.subheader("2. Resultados de la búsqueda")
            
            if len(filtered_df) > 0:
                str_meses = ", ".join(meses_seleccionados)
                str_anios = ", ".join(anios_seleccionados)
                st.success(f"Se encontraron **{len(filtered_df)}** discursos en **{str_meses} {str_anios}**.")
                
                filtered_df_display = filtered_df.copy()
                filtered_df_display["Date"] = filtered_df_display["Date"].dt.strftime('%Y-%m-%d')
                filtered_df_display["Title"] = filtered_df_display.apply(
                    lambda x: f"[{x['Title']}]({x['Link']})", axis=1
                )

                st.markdown(
                    filtered_df_display[["Date", "Title"]].to_markdown(index=False),
                    unsafe_allow_html=True
                )

                st.markdown("---")
                st.subheader("3. Exportar Datos")
                
                word_file = generate_word(filtered_df, title="BIS Central Bank Speeches")
                st.download_button(
                    label="📄 Descargar en Word",
                    data=word_file,
                    # Nombramos el archivo dinámicamente según la selección
                    file_name=f"bis_speeches_{'_'.join(meses_seleccionados)}_{'_'.join(anios_seleccionados)}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            else:
                st.warning("No hay discursos del BIS para las fechas seleccionadas. Intenta con otro mes.")
    else:
        st.info("👆 Selecciona el mes y año arriba y presiona **'Buscar Discursos del BIS'**.")
# MÓDULOS EN CONSTRUCCIÓN (Placeholder para el resto del menú)
else:
    st.info(f"El extractor de **{tipo_doc}** para **{organismo_seleccionado}** está en construcción.")
    st.write("Próximamente podrás extraer estos documentos de forma automatizada.")

