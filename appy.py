import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import re
import zipfile
from io import BytesIO

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
    "📊 Excel base de facturas (con consecutivo)",
    type=["xlsx"]
)

nit = st.text_input("🏷️ NIT del prestador", placeholder="900364721")

if st.button("🚀 Procesar"):

    if not pdfs or not excel or not nit:
        st.error("Faltan archivos o NIT")
        st.stop()

    # Leer Excel
    try:
        df = pd.read_excel(excel, engine="openpyxl")
    except Exception as e:
        st.error(f"Error leyendo el Excel: {e}")
        st.stop()

    # Col A = consecutivo | Col B = ONC
    df.columns = ["consecutivo", "onc"] + list(df.columns[2:])
    df["consecutivo"] = df["consecutivo"].astype(int)
    df["onc"] = df["onc"].astype(str)

    # Diccionario rápido: consecutivo → ONC
    mapa_excel = dict(zip(df["consecutivo"], df["onc"]))

    buffer_zip = BytesIO()
    errores = []

    with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED) as zipf:

        for pdf in pdfs:
            name = pdf.name.replace(".pdf", "")
            match = re.match(r"(\d+)\.(\d+)", name)

            if not match:
                errores.append(f"{pdf.name} (formato inválido)")
                continue

            factura_pdf, subtipo = match.groups()
            factura_pdf = int(factura_pdf)

            if factura_pdf not in mapa_excel:
                errores.append(f"{pdf.name} (no existe en Excel)")
                continue

            onc = mapa_excel[factura_pdf]
            abrev = MAP_ABREV.get(subtipo, "OTRO")

            doc = fitz.open()
            doc.insert_pdf(fitz.open(stream=pdf.read(), filetype="pdf"))

            nombre = f"{abrev}_{nit}_{onc}.pdf"
            zipf.writestr(nombre, doc.write())

    if errores:
        st.warning("⚠️ Algunos archivos no se procesaron:")
        for e in errores:
            st.text(f"- {e}")

    st.success("✅ Proceso completado usando el consecutivo del Excel")

    st.download_button(
        "⬇️ Descargar ZIP",
        buffer_zip.getvalue(),
        file_name="PDFs_Renombrados.zip",
        mime="application/zip"
    )
