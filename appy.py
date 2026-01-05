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
    "📊 Excel (Consecutivo base | Factura)",
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
    pdf_por_consecutivo = defaultdict(list)

    for pdf in pdfs:
        nombre = pdf.name.replace(".pdf", "")
        match = re.match(r"^(\d+)\.(\d+)", nombre)

        if not match:
            continue

        consecutivo_pdf, subtipo = match.groups()
        pdf_por_consecutivo[int(consecutivo_pdf)].append((subtipo, pdf))

    buffer_zip = BytesIO()
    errores = []

    with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED) as zipf:

        for _, fila in df.iterrows():
            consecutivo = fila["consecutivo"]
            factura = fila["factura"]

            if consecutivo not in pdf_por_consecutivo:
                errores.append(f"Consecutivo {consecutivo} sin PDFs asociados")
                continue

            archivos = sorted(
                pdf_por_consecutivo[consecutivo],
                key=lambda x: int(x[0])
            )

            doc = fitz.open()

            for subtipo, pdf in archivos:
                doc.insert_pdf(
                    fitz.open(stream=pdf.read(), filetype="pdf")
                )

            # La abreviatura se toma del primer subtipo
            abrev = MAP_ABREV.get(archivos[0][0], "OTRO")

            nombre_final = f"{abrev}_{nit}_{factura}.pdf"
            zipf.writestr(nombre_final, doc.write())

    if errores:
        st.warning("⚠️ Consecutivos sin PDFs:")
        for e in errores:
            st.text(f"- {e}")

    st.success("✅ Proceso completado correctamente (Excel manda)")

    st.download_button(
        "⬇️ Descargar ZIP",
        buffer_zip.getvalue(),
        file_name="PDFs_Renombrados.zip",
        mime="application/zip"
    )
