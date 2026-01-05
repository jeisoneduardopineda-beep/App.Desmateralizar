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
    "📊 Excel (Col A = consecutivo | Col B = factura)",
    type=["xlsx"]
)

nit = st.text_input("🏷️ NIT del prestador", placeholder="900364721")

if st.button("🚀 Procesar"):

    if not pdfs or not excel or not nit:
        st.error("Faltan archivos o NIT")
        st.stop()

    # === LEER EXCEL ===
    df = pd.read_excel(excel, engine="openpyxl")
    df.columns = ["consecutivo", "factura"] + list(df.columns[2:])

    df["consecutivo"] = df["consecutivo"].astype(int)
    df["factura"] = df["factura"].astype(str).str.strip()

    excel_map = dict(zip(df["consecutivo"], df["factura"]))

    # === AGRUPAR PDFs POR CONSECUTIVO BASE ===
    pdf_groups = defaultdict(list)
    errores = []

    for pdf in pdfs:
        nombre = pdf.name

        m = re.match(r"^(\d+)\.(.+)\.pdf$", nombre)
        if not m:
            errores.append(f"{nombre} (formato inválido)")
            continue

        consecutivo_base = int(m.group(1))
        subtipo = m.group(2)

        pdf_groups[consecutivo_base].append((subtipo, pdf))

    buffer_zip = BytesIO()

    with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED) as zipf:

        for consecutivo, archivos in pdf_groups.items():

            if consecutivo not in excel_map:
                errores.append(f"Consecutivo {consecutivo} no existe en Excel")
                continue

            factura = excel_map[consecutivo]

            doc = fitz.open()

            for _, pdf in sorted(archivos):
                doc.insert_pdf(
                    fitz.open(stream=pdf.read(), filetype="pdf")
                )

            primer_subtipo = archivos[0][0].split(".")[0]
            abrev = MAP_ABREV.get(primer_subtipo, "OTRO")

            nombre_final = f"{abrev}_{nit}_{factura}.pdf"
            zipf.writestr(nombre_final, doc.write())

    if errores:
        st.warning("⚠️ Archivos con observaciones:")
        for e in errores:
            st.text(f"- {e}")

    st.success("✅ Proceso completado. Excel y PDFs correlacionados correctamente.")

    st.download_button(
        "⬇️ Descargar ZIP",
        buffer_zip.getvalue(),
        file_name="PDFs_Renombrados.zip",
        mime="application/zip"
    )
