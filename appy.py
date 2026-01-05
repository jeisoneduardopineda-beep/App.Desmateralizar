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
    "📊 Excel base de facturas",
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

    onc_list = df.iloc[:, 0].astype(str).tolist()

    # 📌 Extraer y ordenar PDFs por el número que traen
    items = []

    for pdf in pdfs:
        name = pdf.name.replace(".pdf", "")
        match = re.match(r"(\d+)\.(\d+)", name)
        if match:
            factura, subtipo = match.groups()
            items.append((int(factura), subtipo, pdf))

    if not items:
        st.error("No se encontraron PDFs con formato factura.subtipo.pdf")
        st.stop()

    # Orden real por número original (solo para mantener orden)
    items.sort(key=lambda x: x[0])

    # 📌 Secuencia corregida
    factura_inicial = items[0][0]
    factura_virtual = factura_inicial

    buffer_zip = BytesIO()

    with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED) as zipf:

        for _, subtipo, pdf in items:

            idx = factura_virtual - factura_inicial

            if idx >= len(onc_list):
                st.warning(f"Factura virtual {factura_virtual} fuera de rango en el Excel")
                break

            doc = fitz.open()
            doc.insert_pdf(fitz.open(stream=pdf.read(), filetype="pdf"))

            onc = onc_list[idx]
            abrev = MAP_ABREV.get(subtipo, "OTRO")

            nombre = f"{abrev}_{nit}_{onc}.pdf"
            zipf.writestr(nombre, doc.write())

            # 🔥 AQUÍ SE CORRIGE EL SALTO
            factura_virtual += 1

    st.success("✅ Proceso completado con corrección automática de consecutivos")

    st.download_button(
        "⬇️ Descargar ZIP",
        buffer_zip.getvalue(),
        file_name="PDFs_Renombrados.zip",
        mime="application/zip"
    )
