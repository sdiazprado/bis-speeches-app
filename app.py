import streamlit as st
import requests
import pandas as pd
import html
from io import BytesIO
import datetime
from docx import Document

# Configuración inicial de la página
st.set_page_config(page_title="BIS Central Bank Speeches", layout="wide")

st.title("BIS Central Bank Speeches Extractor")
st.markdown("Extrae y descarga los discursos de los bancos centrales desde el BIS.")

# 1. Filtros en la pantalla principal (Soluciona el error del calendario cortado)
st.subheader("1. Selecciona el rango de fechas")

hoy = datetime.date.today()
hace_un_mes = hoy - datetime.timedelta(days=30)

# Usamos columnas para que se vea más ordenado
col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input("Fecha de inicio", hace_un_mes)
with col2:
    end_date = st.date_input("Fecha de fin", hoy)

# Botón principal para ejecutar la búsqueda
buscar = st.button("🔍 Buscar Discursos", type="primary")

# 2. Función para descargar los datos del BIS
@st.cache_data(show_spinner="Descargando y procesando datos del BIS...")
def load_data():
    url = "https://www.bis.org/api/document_lists/cbspeeches.json"
    response = requests.get(url)
    response.raise_for_status()
    data = response.json()

    speeches_dict = data.get("list", {})
    rows = []

    for path, speech in speeches_dict.items():
        title = html.unescape(speech.get("short_title", ""))
        date_str = speech.get("publication_start_date", "")

        if not path.endswith(".htm"):
            link = "https://www.bis.org" + path + ".htm"
        else:
            link = "https://www.bis.org" + path

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

# 3. Función para generar el documento de Word
def generate_word(dataframe):
    doc = Document()
    doc.add_heading('BIS Central Bank Speeches', 0)

    # Crear una tabla con 3 columnas y bordes
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    
    # Escribir los encabezados
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Date'
    hdr_cells[1].text = 'Title'
    hdr_cells[2].text = 'Link'

    # Llenar la tabla con los datos iterando sobre el DataFrame
    for index, row in dataframe.iterrows():
        row_cells = table.add_row().cells
        
        # Asegurarnos de que la fecha se escriba como texto en el Word
        if pd.api.types.is_datetime64_any_dtype(row['Date']):
            date_str = row['Date'].strftime('%Y-%m-%d')
        else:
            date_str = str(row['Date'])
            
        row_cells[0].text = date_str
        row_cells[1].text = str(row['Title'])
        row_cells[2].text = str(row['Link'])

    # Guardar en memoria para que Streamlit lo pueda descargar
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

st.markdown("---")

# 4. Lógica de ejecución de la app
if buscar or "df_filtrado" in st.session_state:
    
    df = load_data()
    
    # Filtrar el dataframe por fechas
    mask = (df["Date"].dt.date >= start_date) & (df["Date"].dt.date <= end_date)
    filtered_df = df[mask]
    
    # Guardar en el estado para que no desaparezca al hacer clic en descargar
    st.session_state["df_filtrado"] = filtered_df

    st.subheader("2. Resultados de la búsqueda")
    
    if len(filtered_df) > 0:
        st.success(f"Se encontraron **{len(filtered_df)}** discursos entre {start_date} y {end_date}.")
        
        # Preparar la tabla visual para la pantalla (con links clickeables)
        filtered_df_display = filtered_df.copy()
        filtered_df_display["Date"] = filtered_df_display["Date"].dt.strftime('%Y-%m-%d')
        filtered_df_display["Title"] = filtered_df_display.apply(
            lambda x: f"[{x['Title']}]({x['Link']})", axis=1
        )

        # Mostrar la tabla en Streamlit
        st.markdown(
            filtered_df_display[["Date", "Title"]].to_markdown(index=False),
            unsafe_allow_html=True
        )

        st.markdown("---")
        st.subheader("3. Exportar Datos")
        
        # Generar y mostrar el botón de descarga para Word
        word_file = generate_word(filtered_df)
        st.download_button(
            label="📄 Descargar en Word",
            data=word_file,
            file_name=f"bis_speeches_{start_date}_to_{end_date}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    else:
        st.warning("No hay discursos del BIS en el rango de fechas seleccionado. Intenta ampliar tu búsqueda.")
else:
    st.info("👆 Selecciona las fechas arriba y presiona **'Buscar Discursos'**.")
