import streamlit as st
import pandas as pd
import fitz
import re
import zipfile
from io import BytesIO
from collections import defaultdict

MAP_ABREV = {
    "0": "FAC", "1": "HEV", "2": "EPI", "3": "PDX", "4": "DQX",
    "5": "RAN", "6": "CRC", "7": "TAP", "8": "TNA", "9": "FMO",
    "10": "OPF", "11": "HAM", "12": "ADRES", "13": "PDE"
}

st.set_page_config("Renombrador de PDFs – Radicación", layout="centered")
st.title("Renombrador masivo de PDFs – Radicación")

pdfs = st.file_uploader("📂 Selecciona los PDFs", type="pdf", accept_multiple_files=True)
excel = st.file_uploader("📊 Excel base de facturas", type=["xlsx", "xls"])
nit = st.text_input("🏷️ NIT del prestador", placeholder="900364721")

if st.button("🚀 Procesar"):

    if not (pdfs and excel and nit):
        st.error("Faltan archivos o NIT")
        st.stop()

    ext = os.path.splitext(excel.name)[1].lower()

if ext == ".xlsx":
    df = pd.read_excel(excel, engine="openpyxl")
elif ext == ".xls":
    df = pd.read_excel(excel, engine="xlrd")
else:
    st.error("Formato de Excel no soportado")
    st.stop()

    grupos = defaultdict(list)

    for pdf in pdfs:
        name = pdf.name.replace(".pdf", "")
        match = re.match(r"(\d+)\.(\d+)", name)
        if match:
            factura, subtipo = match.groups()
            grupos[(factura, subtipo)].append(pdf)

    buffer_zip = BytesIO()
    with zipfile.ZipFile(buffer_zip, "w") as zipf:

        for (factura, subtipo), archivos in grupos.items():
            doc = fitz.open()
            for pdf in archivos:
                doc.insert_pdf(fitz.open(stream=pdf.read(), filetype="pdf"))

            idx = int(factura) - 1
            onc = onc_list[idx]
            abrev = MAP_ABREV.get(subtipo, "OTRO")

            nombre = f"{abrev}_{nit}_{onc}.pdf"
            pdf_bytes = doc.write()
            zipf.writestr(nombre, pdf_bytes)

    st.success("Proceso completado")
    st.download_button(
        "⬇️ Descargar ZIP",
        buffer_zip.getvalue(),
        file_name="PDFs_Renombrados.zip",
        mime="application/zip"
    )
