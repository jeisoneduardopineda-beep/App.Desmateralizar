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

    # Diccionario exacto
    excel_map = dict(zip(df["consecutivo"], df["factura"]))

    # === AGRUPAR PDFs POR CONSECUTIVO BASE ===
    pdf_map = defaultdict(list)
    errores = []

    for pdf in pdfs:
        nombre = pdf.name.strip()

        match = re.match(r"^(\d+)\.(\d+)", nombre)
        if not match:
            errores.append(f"{nombre} (formato inválido)")
            continue

        consecutivo = int(match.group(1))
        subtipo = match.group(2)

        pdf_map[consecutivo].append((int(subtipo), pdf))

    buffer_zip = BytesIO()

    with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED) as zipf:

        # 🔥 AHORA SE RECORREN LOS PDFs, NO EL EXCEL
        for consecutivo, archivos in pdf_map.items():

            if consecutivo not in excel_map:
                errores.append(f"PDF con consecutivo {consecutivo} sin referencia en Excel")
                continue

            factura = excel_map[consecutivo]

            # Ordenar por subtipo: 0,1,2,3...
            archivos.sort(key=lambda x: x[0])

            doc = fitz.open()

            for _, pdf in archivos:
                doc.insert_pdf(
                    fitz.open(stream=pdf.read(), filetype="pdf")
                )

            abrev = MAP_ABREV.get(str(archivos[0][0]), "OTRO")

            nombre_final = f"{abrev}_{nit}_{factura}.pdf"
            zipf.writestr(nombre_final, doc.write())

    if errores:
        st.warning("⚠️ Archivos con observaciones:")
        for e in errores:
            st.text(f"- {e}")

    st.success("✅ Proceso completado (PDFs 100% asociados al Excel)")

    st.download_button(
        "⬇️ Descargar ZIP",
        buffer_zip.getvalue(),
        file_name="PDFs_Renombrados.zip",
        mime="application/zip"
    )
