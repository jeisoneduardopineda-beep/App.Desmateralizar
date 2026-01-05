import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import re
import zipfile
import os
from io import BytesIO
from collections import defaultdict

MAP_ABREV = {
    "0": "FAC", "1": "HEV", "2": "EPI", "3": "PDX", "4": "DQX",
    "5": "RAN", "6": "CRC", "7": "TAP", "8": "TNA", "9": "FMO",
    "10": "OPF", "11": "HAM", "12": "ADRES", "13": "PDE"
}

st.set_page_config(
    page_title="Renombrador de PDFs – Radicación",
    layout="centered"
)

st.title("Renombrador masivo de PDFs – Radicación")

pdfs = st.file_uploader(
    "📂 Selecciona los PDFs",
    type="pdf",
    accept_multiple_files=True
)

excel = st.file_uploader(
    "📊 Excel base de facturas (.xlsx)",
    type=["xlsx"]
)

nit = st.text_input("🏷️ NIT del prestador", placeholder="900364721")

if st.button("🚀 Procesar"):

    if not pdfs or not excel or not nit:
        st.error("Faltan archivos o NIT")
        st.stop()

    # 🔹 Leer Excel de forma segura (Cloud-safe)
    try:
        df = pd.read_excel(
            BytesIO(excel.getvalue()),
            engine="openpyxl"
        )
    except Exception as e:
        st.error(f"Error leyendo el Excel: {e}")
        st.stop()

    # 🔹 ONC desde la primera columna
    onc_list = df.iloc[:, 0].astype(str).tolist()

    grupos = defaultdict(list)

    # 🔹 Agrupar PDFs por factura.subtipo
    for pdf in pdfs:
        name = pdf.name.replace(".pdf", "")
        match = re.match(r"(\d+)\.(\d+)", name)
        if match:
            factura, subtipo = match.groups()
            grupos[(factura, subtipo)].append(pdf)

    if not grupos:
        st.warning("No se encontraron PDFs con formato factura.subtipo")
        st.stop()

    buffer_zip = BytesIO()

    with zipfile.ZipFile(buffer_zip, "w") as zipf:
        for (factura, subtipo), archivos in grupos.items():
            doc = fitz.open()

            for pdf in archivos:
                doc.insert_pdf(
                    fitz.open(stream=pdf.read(), filetype="pdf")
                )

            idx = int(factura) - 1

            if idx >= len(onc_list):
                st.warning(f"Factura {factura} fuera de rango en el Excel")
                continue

            onc = onc_list[idx]
            abrev = MAP_ABREV.get(subtipo, "OTRO")
            nombre = f"{abrev}_{nit}_{onc}.pdf"

            zipf.writestr(nombre, doc.write())

    st.success("✅ Proceso completado")

    st.download_button(
        "⬇️ Descargar ZIP",
        buffer_zip.getvalue(),
        file_name="PDFs_Renombrados.zip",
        mime="application/zip"
    )
