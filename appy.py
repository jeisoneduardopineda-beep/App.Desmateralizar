import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import re
import zipfile
from io import BytesIO
from collections import defaultdict

# ===============================
# MAPA DE ABREVIATURAS
# ===============================
MAP_ABREV = {
    "0": "FAC",
    "1": "HEV",
    "2": "EPI",
    "3": "PDX",
    "4": "DQX",
    "5": "RAN",
    "6": "CRC",
    "7": "TAP",
    "8": "TNA",
    "9": "FMO",
    "10": "OPF",
    "11": "HAM",
    "12": "ADRES",
    "13": "PDE"
}

# ===============================
# CONFIG STREAMLIT
# ===============================
st.set_page_config(
    page_title="Renombrador de PDFs – Radicación",
    layout="centered"
)

st.title("Renombrador de PDFs (Desmaterializacion) – Soportes-Radicación")

pdfs = st.file_uploader(
    "📂 Selecciona los PDFs",
    type="pdf",
    accept_multiple_files=True
)

excel = st.file_uploader(
    "📊 Excel base de facturas (con consecutivo)",
    type=["xlsx"]
)

nit = st.text_input("🏷️ NIT del prestador", placeholder="*********")

# ===============================
# PROCESO
# ===============================
if st.button("🚀 Procesar"):

    if not pdfs or not excel or not nit:
        st.error("Faltan PDFs, Excel o NIT")
        st.stop()

    # -------- LEER EXCEL --------
    df = pd.read_excel(excel, engine="openpyxl")
    df = df.iloc[:, :2]
    df.columns = ["consecutivo", "factura"]

    df["consecutivo"] = df["consecutivo"].astype(int)
    df["factura"] = df["factura"].astype(str)

    mapa_excel = dict(zip(df["consecutivo"], df["factura"]))

    # -------- AGRUPAR PDFs --------
    pdf_groups = defaultdict(list)
    errores = []

    for pdf in pdfs:
        nombre = pdf.name.strip()

        # Acepta:
        # 37.1.pdf
        # 37.1 (2).pdf
        # 37.1 (3).pdf
        match = re.match(
            r"^(\d+)\.(\d+)(?:\s*\(\d+\))?\.pdf$",
            nombre
        )

        if not match:
            errores.append(f"{nombre} → formato inválido")
            continue

        consecutivo = int(match.group(1))
        subtipo = match.group(2)  # ESTE es el subtipo real

        pdf_groups[(consecutivo, subtipo)].append(pdf)

    # -------- PROCESAR --------
    buffer_zip = BytesIO()

    with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED) as zipf:

        for (consecutivo, subtipo), archivos in pdf_groups.items():

            if consecutivo not in mapa_excel:
                errores.append(f"Consecutivo {consecutivo} no existe en Excel")
                continue

            if subtipo not in MAP_ABREV:
                errores.append(
                    f"Subtipo {subtipo} no definido en MAP_ABREV (consecutivo {consecutivo})"
                )
                continue  # 🔥 NO SE CREA PDF

            factura = mapa_excel[consecutivo]
            abrev = MAP_ABREV[subtipo]

            # Unificar PDFs
            doc = fitz.open()
            for pdf in archivos:
                doc.insert_pdf(
                    fitz.open(stream=pdf.read(), filetype="pdf")
                )

            nombre_final = f"{abrev}_{nit}_{factura}.pdf"
            zipf.writestr(nombre_final, doc.write())

    # -------- RESULTADOS --------
    if errores:
        st.warning("⚠️ Archivos NO procesados:")
        for e in errores:
            st.text(f"- {e}")

    st.success("✅ Proceso completado (sin OTRO, validación estricta)")

    st.download_button(
        "⬇️ Descargar ZIP",
        buffer_zip.getvalue(),
        file_name="PDFs_Renombrados.zip",
        mime="application/zip"
    )
