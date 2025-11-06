import streamlit as st
import pandas as pd
import pytesseract
from pytesseract import Output
from pdf2image import convert_from_bytes
from docx import Document
import textract
import tempfile, os, zipfile, io, re
from datetime import datetime
from PIL import Image

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Ford Fiorasi – Procesador de Antecedentes Disciplinarios",
    page_icon="⚖️",
    layout="wide",
)

# --- COLORES BASE ---
if "base_color" not in st.session_state:
    st.session_state.base_color = "#003399"  # Azul Ford

# --- ENCABEZADO ---
col_logo, col_title, col_conf = st.columns([1, 6, 1])
with col_logo:
    st.image("logo_ford_fiorasi.png", width=140)
with col_title:
    st.markdown(
        f"<h2 style='color:{st.session_state.base_color};'>Ford Fiorasi – Procesador de Antecedentes Disciplinarios</h2>",
        unsafe_allow_html=True)
with col_conf:
    with st.expander("⚙️ Ajustes"):
        color = st.color_picker("Color institucional", st.session_state.base_color)
        if color:
            st.session_state.base_color = color

st.write("---")

# --- SUBIR ARCHIVOS ---
st.subheader("📂 Seleccione los archivos (.pdf / .doc / .docx)")
uploaded_files = st.file_uploader("Arrastre o seleccione múltiples archivos", type=["pdf", "docx", "doc"], accept_multiple_files=True)

# --- FUNCIONES AUXILIARES ---
def extract_text_from_docx(file):
    """Extrae texto de archivos .docx"""
    try:
        doc = Document(file)
        return "\n".join(p.text for p in doc.paragraphs)
    except Exception as e:
        return f"ERROR_DOCX: {e}"

def extract_text_from_doc(file_bytes):
    """Intenta leer archivos .doc antiguos usando textract"""
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".doc") as tmp:
            tmp.write(file_bytes)
            tmp_path = tmp.name
        text = textract.process(tmp_path).decode("utf-8", errors="ignore")
        return text
    except Exception as e:
        return f"ERROR_DOC: {e}"

def extract_text_from_pdf(file_bytes):
    """Aplica OCR en español para PDFs"""
    try:
        pages = convert_from_bytes(file_bytes)
        text = ""
        for page in pages:
            text += pytesseract.image_to_string(page, lang="spa")
        return text
    except Exception as e:
        return f"ERROR_PDF: {e}"

def extract_data_from_text(text):
    data = {}
    nombre = re.findall(r"([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+(?:\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+)+)", text)
    fecha = re.findall(r"\d{1,2}/\d{1,2}/\d{2,4}", text)
    tipo = re.findall(r"(llamado de atención|apercibimiento|descargo|contestación)", text.lower())
    data["Empleado"] = nombre[0] if nombre else "No detectado"
    data["Fecha"] = fecha[0] if fecha else ""
    data["Tipo"] = tipo[0].capitalize() if tipo else "No identificado"
    data["Descripción"] = text[:200].replace("\n", " ") + "..."
    data["Descargo"] = "Sí" if "descargo" in text.lower() else "No"
    return data

# --- PROCESAMIENTO ---
if st.button("🚀 Procesar antecedentes") and uploaded_files:
    registros = []
    output_dir = tempfile.mkdtemp()

    for file in uploaded_files:
        with st.spinner(f"Procesando {file.name}..."):
            file_ext = file.name.lower().split(".")[-1]
            if file_ext == "docx":
                text = extract_text_from_docx(file)
            elif file_ext == "doc":
                text = extract_text_from_doc(file.read())
            else:
                text = extract_text_from_pdf(file.read())

            if "ERROR_" in text:
                st.error(f"❌ No se pudo leer {file.name}.")
                continue

            datos = extract_data_from_text(text)
            registros.append(datos)

            emp_folder = os.path.join(output_dir, datos["Empleado"].replace(" ", "_"))
            os.makedirs(emp_folder, exist_ok=True)
            with open(os.path.join(emp_folder, file.name), "wb") as f:
                f.write(file.getbuffer())

    if registros:
        df = pd.DataFrame(registros)
        df = df.sort_values(by="Empleado")
        resumen = df.groupby("Empleado").agg({
            "Tipo": lambda x: ", ".join(sorted(set(x))),
            "Fecha": "last",
            "Descripción": lambda x: " | ".join(x)[:500]
        }).reset_index()
        resumen["Cantidad"] = df.groupby("Empleado")["Tipo"].count().values

        año = datetime.now().year
        excel_name = f"FordFiorasi_Antecedentes_Base_{año}.xlsx"
        excel_path = os.path.join(output_dir, excel_name)
        with pd.ExcelWriter(excel_path) as writer:
            df.to_excel(writer, sheet_name="Base Completa", index=False)
            resumen.to_excel(writer, sheet_name="Resumen por Empleado", index=False)

        with open(excel_path, "rb") as f:
            st.download_button("⬇️ Descargar Excel", data=f, file_name=excel_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as z:
            for root, _, files in os.walk(output_dir):
                for file in files:
                    z.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), output_dir))
        st.download_button("🗂️ Descargar ZIP (Excel + Carpetas)", data=zip_buffer.getvalue(), file_name="FordFiorasi_Procesados.zip")

    else:
        st.warning("⚠️ No se procesaron archivos válidos.")
