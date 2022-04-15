import base64
from io import StringIO
import streamlit as st
import pandas as pd
import numpy as np
from zipfile import ZipFile


from utils import *


st.title('SANLAM : INTERFACE DE TRAITEMENT <TRIANGLE>')

fichier = st.file_uploader("Choisir un fichier EXCEL", type=["xlsx", "xls"])


if fichier is not None:

    details = {"filename":fichier.name, "filetype":fichier.type, "filesize":fichier.size}
			
    st.write(details)

    data = prepare_data(fichier)
    st.write(data.head())

    
    nbranches = deter(data, "BRANCHE")
    st.write("J'AI TROUVÉ :  {} branches".format(len(nbranches)))
    st.write("JE LANCE LE TRAITEMENT MAINTENANT")

    fnow = datetime.now().strftime("%d-%m-%Y_%H:%M:%S")
    fnzip = "data_{}.zip".format(fnow)
    fpzip = "./zips/"+fnzip
    datazip = ZipFile(fpzip, 'w')

    for nbranche in nbranches[:2]:  
        st.write("==================================================")
        st.write("BRANCHE {} EN COURS DE TRAITEMENT".format(nbranche))
        dob = data[data["BRANCHE"] == nbranche]
        st.write("BRANCHE {} TRAITÉE AVEC SUCCÈS".format(nbranche))
        fp, fn = traitementv4(dob)
        datazip.write(fp)
        st.write("--------------------------------------------------")
    datazip.close()

    with open(fpzip, "rb") as f:
        bytes = f.read()
        b64 = base64.b64encode(bytes).decode()
        href = f'<a href="data:file/zip;base64,{b64}" download=\'{fnzip}\'>\
                JE PROCÈDE AU TÉLÉCHARGeMENT\
            </a>'
        st.sidebar.markdown(href, unsafe_allow_html=True)
    
