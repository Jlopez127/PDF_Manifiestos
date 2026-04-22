from __future__ import annotations

import io
from datetime import datetime

import streamlit as st
from openpyxl import load_workbook

GENERATOR_IMPORT_ERROR: ModuleNotFoundError | None = None

try:
    from generate_shipping_labels import generate_zip_bytes, read_rows
except ModuleNotFoundError as exc:
    GENERATOR_IMPORT_ERROR = exc
    generate_zip_bytes = None
    read_rows = None


st.set_page_config(
    page_title="Generador de PDFs por Manifiesto",
    page_icon="📦",
    layout="centered",
)


def get_sheet_names(file_bytes: bytes) -> list[str]:
    workbook = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    return workbook.sheetnames


st.title("Generador de etiquetas 4x6")
st.write("Sube un Excel y descarga un ZIP con un PDF por cada fila del manifiesto.")

if GENERATOR_IMPORT_ERROR is not None:
    st.error("Falta una dependencia necesaria para generar los PDFs.")
    st.code("pip install -r requirements.txt")
    st.write("Dependencia reportada:", f"`{GENERATOR_IMPORT_ERROR}`")
    st.stop()

uploaded_file = st.file_uploader(
    "Cargar archivo Excel",
    type=["xlsx"],
    accept_multiple_files=False,
)

if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()
    sheet_names = get_sheet_names(file_bytes)

    selected_sheet = st.selectbox(
        "Hoja a procesar",
        options=sheet_names,
        index=0,
    )

    if st.button("Preparar ZIP", type="primary"):
        with st.spinner("Generando PDFs..."):
            rows = read_rows(io.BytesIO(file_bytes), selected_sheet)

            if not rows:
                st.error("No se encontraron filas validas con numero de Envio o Guia.")
            else:
                zip_bytes = generate_zip_bytes(rows)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                zip_name = f"etiquetas_{selected_sheet}_{timestamp}.zip"

                st.success(f"Se prepararon {len(rows)} etiquetas.")
                st.download_button(
                    label="Descargar ZIP con PDFs",
                    data=zip_bytes,
                    file_name=zip_name,
                    mime="application/zip",
                )
else:
    st.info("Selecciona un archivo Excel para comenzar.")
