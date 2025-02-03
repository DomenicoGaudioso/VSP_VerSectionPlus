import streamlit as st
import zipfile
import os
import io

def list_files_in_zip(zip_file):
    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        file_list = zip_ref.namelist()
    return file_list

st.title("Caricamento e Visualizzazione di File ZIP")

# Caricamento del file ZIP
uploaded_file = st.file_uploader("Carica un file ZIP", type=["zip"])

if uploaded_file is not None:
    st.success("File caricato con successo!")
    
    # Leggere il contenuto del file ZIP
    file_contents = io.BytesIO(uploaded_file.read())
    file_list = list_files_in_zip(file_contents)
    
    st.subheader("Contenuti della cartella ZIP:")
    for file in file_list:
        st.write(f"- {file}")
