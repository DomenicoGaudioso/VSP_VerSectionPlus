import streamlit as st
import zipfile
import os
import io
import tempfile
#from stpyvista import stpyvista
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
    sottocartella1 = trova_sottocartelle(extract_path)
    #sottocartella2 = trova_sottocartelle(sottocartella1[0])
    sottocartelle = trova_sottocartelle(sottocartella1[0])
    st.write("📂 Sottocartelle trovate:", sottocartelle)


pathStar = sottocartella1[0]
pathSec = sottocartelle
#print(pathSec)

#path = pathSec[1]

word_save = os.path.join(pathStar, "Report Verifiche.docx")
# open an existing document
doc = docx.Document()
doc.add_heading("Report verifiche", level=1)

for i, item in enumerate(pathSec):
    #print(pathSec)
    #print(pathSec[i])
    path = pathSec[i]
    path_cds = file_per_estensione(path, estensione=".xlsx", iniziali="cds")[0]
    #print(path_cds)
    cds = pd.read_excel(path_cds, usecols=range(1, 11, 1), skiprows= 1)

    #st.write(cds)
    
    cls_dict, steel_dict = setmaterial(path) # settaggio dei materiali
    conc_sec = bildSection(path, cls_dict, steel_dict) # costruzione della sezione
    #st.write(conc_sec)
    im3d = domino3D(conc_sec, cls_dict, steel_dict, cds, n_points=5, n_level=5) # costruzione del dominio 3D
    #im3d.show(interactive=True) #, auto_close=False
    #input("Premi Invio per chiudere...")
    #im3d.screenshot(r"C:\Users\d.gaudioso\Desktop\prova.png", window_size=[2020, 3035])  # Salva l'immagine in un file PNG
    figure = subplot_figure1(im3d, conc_sec, cls_dict, steel_dict)
    image_stream = BytesIO()
    figure.savefig(image_stream, format="png")  # Salva l'immagine in memoria
    image_stream.seek(0)  # Torna all'inizio del file in memoria

    figure.savefig(path)
    print('Figure written File successfully.')

    # saving the excel
    nomefile = "verifiche_MxMyN_" + os.path.basename(path) +".xlsx"
    cds.to_excel(os.path.join(path, nomefile))
    print('DataFrame is written to Excel File successfully.')

    ### TO WORD##
    doc.add_heading(os.path.basename(path), level=2)
    #Inserimento dell'immagine dal buffer di memoria
    doc.add_picture(image_stream, width=docx.shared.Inches(5))  # Inserisce l'immagine dal buffer

    # 📌 Generiamo le immagini del dataframe
    image_paths = save_dataframe_images(cds, rows_per_page=70)

    # 📌 Inseriamo le immagini in Word
    doc.add_paragraph("Risultati in forma tabellare")

    for img in image_paths:
        doc.add_picture(img, width = docx.shared.Inches(6))
        doc.add_page_break()  # Aggiunge un'interruzione di pagina

# 📌 Salviamo il documento
doc.save(word_save)



    #st.download_button("📥 Scarica Report Word", word_stream, "Report_Verifiche.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    #st.success("✅ Report Word generato con successo!")
