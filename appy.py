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

st.set_page_config(page_title="Renombrador PDFs", layout="centered")
st.title("Renombrador masivo de PDFs – Radicación")

pdfs = st.file_uploader(
    "📂 Selecciona los PDFs",
    type="pdf",
    accept_multiple_files=True
)

excel = st.file_uploader(
    "📊 Excel base (Consecutivo | Número factura)",
    type=["xlsx"]
)

nit = st.text_input("🏷️ NIT del prestador", placeholder="900364721")

if st.button("🚀 Procesar"):

    if not pdfs or not excel or not nit:
        st.error("Faltan archivos o NIT")
        st.stop()

    # 🔹 Leer Excel
    df = pd.read_excel(excel, engine="openpyxl")
    df.columns = ["consecutivo", "factura_final"] + list(df.columns[2:])
    df["consecutivo"] = df["consecutivo"].astype(int)
    df["factura_final"] = df["factura_final"].astype(str)

    mapa_excel = dict(zip(df["consecutivo"], df["factura_final"]))

    # 🔹 Agrupar PDFs: consecutivo + tipo_documento
    grupos = defaultdict(list)

    for pdf in pdfs:
        nombre = pdf.name.replace(".pdf", "")
        match = re.match(r"(\d+)\.(\d+)(?:\.(\d+))?$", nombre)

        if not match:
            continue

        consecutivo, tipo, fragmento = match.groups()
        consecutivo = int(consecutivo)

        grupos[(consecutivo, tipo)].append((fragmento, pdf))

    buffer_zip = BytesIO()
    errores = []

    with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED) as zipf:

        for (consecutivo, tipo), archivos in grupos.items():

            if consecutivo not in mapa_excel:
                errores.append(f"{consecutivo}.{tipo} → no existe en Excel")
                continue

            factura_final = mapa_excel[consecutivo]
            abrev = MAP_ABREV.get(tipo, "OTRO")

            # 🔹 Ordenar fragmentos (None primero)
            archivos.sort(key=lambda x: (x[0] is not None, x[0] or 0))

            doc = fitz.open()

            for _, pdf in archivos:
                doc.insert_pdf(
                    fitz.open(stream=pdf.read(), filetype="pdf")
                )

            nombre_final = f"{abrev}_{nit}_{factura_final}.pdf"
            zipf.writestr(nombre_final, doc.write())

    if errores:
        st.warning("⚠️ Algunos documentos no se procesaron:")
        for e in errores:
            st.text(f"- {e}")

    st.success("✅ Proceso completado correctamente")

    st.download_button(
        "⬇️ Descargar ZIP",
        buffer_zip.getvalue(),
        file_name="PDFs_Renombrados.zip",
        mime="application/zip"
    )
