import streamlit as st
import zipfile
import os
import io
import tempfile
from src_bisantis import *

def list_files_in_zip(zip_file):
    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        file_list = zip_ref.namelist()
    return file_list

st.title("Analisi sezioni in c.a.")

# Funzione per trovare sottocartelle

# Funzione per estrarre ZIP
def extract_zip(uploaded_file):
    with zipfile.ZipFile(uploaded_file, "r") as zip_ref:
        extract_path = "temp_extracted"
        zip_ref.extractall(extract_path)
    return extract_path

# Interfaccia Streamlit
st.title("Analisi File ZIP e Generazione Report Word")

uploaded_zip = st.file_uploader("Carica un file ZIP", type=["zip"])

if uploaded_zip:
    st.success("📂 File ZIP caricato con successo!")

    # Estrarre il file ZIP
    extract_path = extract_zip(uploaded_zip)
    st.write(f"📂 File estratti in: `{extract_path}`")

    # Trovare le sottocartelle
    sottocartelle = trova_sottocartelle(extract_path)
    st.write("📂 Sottocartelle trovate:", sottocartelle)

    # Creazione documento Word
    doc = docx.Document()
    doc.add_heading("Report verifiche", level=1)

    for i, path in enumerate(sottocartelle):
        st.write(f"🔍 Analizzando: `{path}`")

        # Trovare il primo file XLSX che inizia con "cds"
        file_cds = file_per_estensione(path, estensione=".xlsx", iniziali="cds")
        
        if file_cds:
            st.write(f"📊 File trovato: `{file_cds[0]}`")
            df = pd.read_excel(file_cds[0], sheet_name="materiali")  # Leggiamo il foglio "materiali"
            st.dataframe(df.head())  # Mostriamo un'anteprima su Streamlit

            # Aggiungere intestazione al documento Word
            doc.add_heading(f"📄 {os.path.basename(path)}", level=2)

            # Salvataggio del dataframe in tabella Word
            table = doc.add_table(rows=df.shape[0]+1, cols=df.shape[1])
            for j, col_name in enumerate(df.columns):
                table.cell(0, j).text = col_name
            for i, row in df.iterrows():
                for j, val in enumerate(row):
                    table.cell(i+1, j).text = str(val)

        else:
            st.warning(f"⚠️ Nessun file `cds*.xlsx` trovato in `{path}`")

    # Salvataggio finale
    word_stream = BytesIO()
    doc.save(word_stream)
    word_stream.seek(0)

    st.download_button("📥 Scarica Report Word", word_stream, "Report_Verifiche.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    st.success("✅ Report Word generato con successo!")
