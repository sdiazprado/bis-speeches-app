import streamlit as st
import requests
import pandas as pd
import html
from io import BytesIO

st.set_page_config(page_title="BIS Central Bank Speeches", layout="wide")

st.title("BIS Central Bank Speeches Extractor")

@st.cache_data
def load_data():
    url = "https://www.bis.org/api/document_lists/cbspeeches.json"
    response = requests.get(url)
    response.raise_for_status()
    data = response.json()

    speeches_dict = data["list"]
    rows = []

    for path, speech in speeches_dict.items():
        title = html.unescape(speech.get("short_title", ""))
        date = speech.get("publication_start_date", "")

        if not path.endswith(".htm"):
            link = "https://www.bis.org" + path + ".htm"
        else:
            link = "https://www.bis.org" + path

        rows.append({
            "Date": date,
            "Title": title,
            "Link": link
        })

    df = pd.DataFrame(rows)
    df["Date"] = pd.to_datetime(df["Date"])
    df = df.sort_values("Date", ascending=False)

    return df


df = load_data()

# 📅 Filtros de fecha
st.sidebar.header("Filter by Date")

min_date = df["Date"].min()
max_date = df["Date"].max()

start_date = st.sidebar.date_input("Start date", min_date)
end_date = st.sidebar.date_input("End date", max_date)

filtered_df = df[
    (df["Date"] >= pd.to_datetime(start_date)) &
    (df["Date"] <= pd.to_datetime(end_date))
]

st.subheader("Filtered Results")
st.write(f"Showing {len(filtered_df)} speeches")

# Mostrar tabla con links clickeables
filtered_df_display = filtered_df.copy()
filtered_df_display["Title"] = filtered_df_display.apply(
    lambda x: f"[{x['Title']}]({x['Link']})", axis=1
)

st.markdown(
    filtered_df_display[["Date", "Title"]].to_markdown(index=False),
    unsafe_allow_html=True
)

# 📥 Descargar Excel
def generate_excel(dataframe):
    output = BytesIO()
    dataframe_to_export = dataframe.copy()
    
    dataframe_to_export["Title"] = dataframe_to_export.apply(
        lambda x: f'=HYPERLINK("{x["Link"]}","{x["Title"]}")',
        axis=1
    )
    
    dataframe_to_export = dataframe_to_export.drop(columns=["Link"])
    
    dataframe_to_export.to_excel(output, index=False)
    output.seek(0)
    return output

excel_file = generate_excel(filtered_df)

st.download_button(
    label="Download Excel",
    data=excel_file,
    file_name="bis_speeches_filtered.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)