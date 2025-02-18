import streamlit as st
import zipfile
import os
import io
from src_bisantis import *

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

pathStar = r"Z:\studio\CODIFICATE\200 BISantis - due volte prima\Areatecnica\Calcolo\Verifiche Domini"
pathSec = trova_sottocartelle(pathStar)
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
    
    cls_dict, steel_dict = setmaterial(path) # settaggio dei materiali
    conc_sec = bildSection(path, cls_dict, steel_dict) # costruzione della sezione
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