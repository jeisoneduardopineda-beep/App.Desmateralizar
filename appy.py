import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import re
import zipfile
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
    "📊 Excel base (Consecutivo | Factura)",
    type=["xlsx"]
)

nit = st.text_input("🏷️ NIT del prestador", placeholder="900364721")

if st.button("🚀 Procesar"):

    if not pdfs or not excel or not nit:
        st.error("Faltan archivos o NIT")
        st.stop()

    # 📌 Leer Excel
    df = pd.read_excel(excel, engine="openpyxl")
    df.columns = ["consecutivo", "factura"] + list(df.columns[2:])
    df["consecutivo"] = df["consecutivo"].astype(int)
    df["factura"] = df["factura"].astype(str)

    # 📌 Agrupar PDFs por consecutivo base
    grupos_pdf = defaultdict(list)

    for pdf in pdfs:
        nombre = pdf.name.replace(".pdf", "")
        match = re.match(r"(\d+)\.(\d+)", nombre)

        if not match:
            continue

        consecutivo_pdf, subtipo = match.groups()
        grupos_pdf[int(consecutivo_pdf)].append((int(subtipo), pdf))

    buffer_zip = BytesIO()
    errores = []

    with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED) as zipf:

        for _, fila in df.iterrows():
            consecutivo = fila["consecutivo"]
            factura = fila["factura"]

            if consecutivo not in grupos_pdf:
                errores.append(f"Consecutivo {consecutivo} sin PDFs")
                continue

            # Ordenar por subtipo (0,1,2…)
            archivos = sorted(grupos_pdf[consecutivo], key=lambda x: x[0])

            doc = fitz.open()

            for subtipo, pdf in archivos:
                doc.insert_pdf(
                    fitz.open(stream=pdf.read(), filetype="pdf")
                )

            # Usar el subtipo del primer archivo para la abreviatura
            subtipo_base = str(archivos[0][0])
            abrev = MAP_ABREV.get(subtipo_base, "OTRO")

            nombre_final = f"{abrev}_{nit}_{factura}.pdf"
            zipf.writestr(nombre_final, doc.write())

    if errores:
        st.warning("⚠️ Algunos consecutivos no tenían PDFs:")
        for e in errores:
            st.text(f"- {e}")

    st.success("✅ Proceso completado correctamente usando el consecutivo del Excel")

    st.download_button(
        "⬇️ Descargar ZIP",
        buffer_zip.getvalue(),
        file_name="PDFs_Renombrados.zip",
        mime="application/zip"
    )
