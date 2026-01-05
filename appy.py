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

# ===============================
# PROCESO
# ===============================
if st.button("🚀 Procesar"):

    if not pdfs or not excel or not nit:
        st.error("Faltan PDFs, Excel o NIT")
        st.stop()

    # -------- LEER EXCEL --------
    try:
        df = pd.read_excel(excel, engine="openpyxl")
    except Exception as e:
        st.error(f"Error leyendo el Excel: {e}")
        st.stop()

    # Forzamos estructura:
    # Col A = consecutivo | Col B = factura
    df = df.iloc[:, :2]
    df.columns = ["consecutivo", "factura"]

    df["consecutivo"] = df["consecutivo"].astype(int)
    df["factura"] = df["factura"].astype(str)

    # Diccionario CLAVE → VALOR
    mapa_excel = dict(zip(df["consecutivo"], df["factura"]))

    # -------- AGRUPAR PDFs POR CONSECUTIVO BASE --------
    pdf_groups = defaultdict(list)
    errores = []

    for pdf in pdfs:
        nombre = pdf.name.strip()

        # Captura:
        # 37.0.pdf
        # 37.2.1.pdf
        match = re.match(r"^(\d+)\.(.+)\.pdf$", nombre)
        if not match:
            errores.append(f"{nombre} (formato inválido)")
            continue

        consecutivo = int(match.group(1))
        subtipo = match.group(2)

        pdf_groups[consecutivo].append((subtipo, pdf))

    # -------- DEBUG VISUAL --------
    st.write("📌 Consecutivos en Excel:", sorted(mapa_excel.keys()))
    st.write("📌 Consecutivos detectados en PDFs:", sorted(pdf_groups.keys()))

    # -------- PROCESAR --------
    buffer_zip = BytesIO()

    with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED) as zipf:

        for consecutivo, archivos in pdf_groups.items():

            if consecutivo not in mapa_excel:
                errores.append(f"Consecutivo {consecutivo} no existe en Excel")
                continue

            factura = mapa_excel[consecutivo]

            for subtipo, pdf in archivos:
                # Tomar SOLO el primer número del subtipo
                base_subtipo = subtipo.split(".")[0]
                abrev = MAP_ABREV.get(base_subtipo, "OTRO")

                doc = fitz.open()
                doc.insert_pdf(
                    fitz.open(stream=pdf.read(), filetype="pdf")
                )

                nombre_final = f"{abrev}_{nit}_{factura}.pdf"
                zipf.writestr(nombre_final, doc.write())

    # -------- RESULTADOS --------
    if errores:
        st.warning("⚠️ Archivos con problemas:")
        for e in errores:
            st.text(f"- {e}")

    st.success("✅ Proceso completado correctamente")

    st.download_button(
        "⬇️ Descargar ZIP",
        buffer_zip.getvalue(),
        file_name="PDFs_Renombrados.zip",
        mime="application/zip"
    )
