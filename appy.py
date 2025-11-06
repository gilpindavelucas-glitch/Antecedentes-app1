import streamlit as st
from io import BytesIO
import os, re, shutil, datetime, pandas as pd, zipfile
from pathlib import Path
import docx2txt
import fitz
from PIL import Image
import pytesseract
import dateparser
import PyPDF2

# ----- Configuración -----
DEFAULT_BLUE = "#003399"
LOGO_PATH = "logo_ford_fiorasi.png"
APP_TITLE = "Ford Fiorasi – Procesador de Antecedentes Disciplinarios"
OUTPUT_FOLDER = "procesados_output"
LANG_OCR = "spa"
# --------------------------

st.set_page_config(page_title=APP_TITLE, layout="wide", page_icon=":briefcase:")

if "color" not in st.session_state:
    st.session_state.color = DEFAULT_BLUE
if "use_ocr" not in st.session_state:
    st.session_state.use_ocr = True

# ---------- Encabezado ----------
col1, col2 = st.columns([1,6])
with col1:
    st.image(LOGO_PATH, width=180)
with col2:
    st.markdown(f"<h1 style='color:{st.session_state.color};margin-left:20px'>{APP_TITLE}</h1>", unsafe_allow_html=True)

# ---------- Configuración ----------
with st.expander("⚙️ Ajustes / Configuración"):
    st.session_state.color = st.color_picker("Color institucional", st.session_state.color)
    st.session_state.use_ocr = st.checkbox("Activar OCR español para PDF escaneados", True)

st.markdown("---")
st.write("Subí tus archivos (.docx o .pdf) y presioná **Procesar antecedentes**.")

uploaded_files = st.file_uploader("Archivos a procesar", accept_multiple_files=True, type=["pdf", "docx"])

# ---------- Funciones auxiliares ----------
def extract_text_docx(bts):
    tmp = "tmp.docx"
    with open(tmp, "wb") as f:
        f.write(bts)
    text = docx2txt.process(tmp)
    os.remove(tmp)
    return text or ""

def extract_text_pdf(bts):
    text = ""
    try:
        reader = PyPDF2.PdfReader(BytesIO(bts))
        for p in reader.pages:
            txt = p.extract_text() or ""
            text += txt + "\n"
    except:
        pass
    if text.strip():
        return text
    if st.session_state.use_ocr:
        try:
            doc = fitz.open(stream=bts, filetype="pdf")
            alltext = []
            for page in doc:
                pix = page.get_pixmap(matrix=fitz.Matrix(2,2))
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                t = pytesseract.image_to_string(img, lang=LANG_OCR)
                alltext.append(t)
            return "\n".join(alltext)
        except Exception as e:
            st.error(f"OCR falló: {e}")
            return ""
    return ""

def detect_date(text):
    m = re.search(r"(\d{1,2}\s+de\s+[A-Za-zñáéíóú]+\s+de\s+\d{4})", text)
    if m:
        dt = dateparser.parse(m.group(1), languages=['es'])
        if dt:
            return dt.date().isoformat()
    m = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", text)
    if m:
        return m.group(1)
    return ""

def detect_type(text):
    t = text.lower()
    if "llamado" in t and "atencion" in t:
        return "Llamado de atención"
    if "apercib" in t:
        return "Apercibimiento"
    if "descargo" in t and "solic" in t:
        return "Solicitud de descargo"
    if "contestac" in t or "respuesta" in t:
        return "Contestación"
    return "No determinado"

def detect_response(text):
    return "Sí" if any(k in text.lower() for k in ["descargo", "contestac"]) else "No"

def detect_name(text):
    m = re.search(r"(Apellido.?y.?Nombre.?[:\s]+)([A-ZÁÉÍÓÚÑ][A-Za-zÁÉÍÓÚñáéíóú\s]+)", text, flags=re.IGNORECASE)
    if m:
        return m.group(2).strip()
    m = re.search(r"(Sr\.|Sra\.|Señor|Señora)\s+([A-Z][a-zA-ZñáéíóúÁÉÍÓÚ\s]+)", text)
    if m:
        return m.group(2).strip()
    return "SIN_NOMBRE"

def resumen_texto(text):
    parts = re.split(r"\n\s*\n", text)
    for p in parts:
        if len(p.strip()) > 50:
            return p.strip().replace("\n", " ")
    return text[:200]

# ---------- Procesamiento ----------
if st.button("Procesar antecedentes"):
    if not uploaded_files:
        st.error("Debe subir al menos un archivo.")
    else:
        out_dir = Path(OUTPUT_FOLDER)
        out_dir.mkdir(exist_ok=True)

        registros = []
        agrupado = {}
        progress = st.progress(0)

        for i, f in enumerate(uploaded_files):
            content = f.read()
            if f.name.lower().endswith(".docx"):
                texto = extract_text_docx(content)
            else:
                texto = extract_text_pdf(content)

            nombre = detect_name(texto)
            fecha = detect_date(texto)
            tipo = detect_type(texto)
            desc = resumen_texto(texto)
            resp = detect_response(texto)

            carpeta = out_dir / re.sub(r"[^\w\s-]", "_", nombre)
            carpeta.mkdir(exist_ok=True)
            with open(carpeta / f.name, "wb") as x:
                x.write(content)

            registros.append({
                "Apellido y Nombre": nombre,
                "Fecha": fecha,
                "Tipo": tipo,
                "Descripción breve": desc,
                "Contestación/Descargo": resp,
                "Archivo": f.name
            })
            agrupado.setdefault(nombre, []).append({
                "tipo": tipo, "fecha": fecha, "desc": desc
            })
            progress.progress((i+1)/len(uploaded_files))

        df = pd.DataFrame(registros).sort_values(by="Apellido y Nombre")
        resumen = []
        for n, docs in agrupado.items():
            tipos = ", ".join(sorted(set(d["tipo"] for d in docs)))
            ult = max([d["fecha"] for d in docs if d["fecha"]], default="")
            sint = " | ".join(d["desc"][:100] for d in docs)
            resumen.append({
                "Apellido y Nombre": n,
                "Cantidad": len(docs),
                "Tipos": tipos,
                "Última fecha": ult,
                "Síntesis": sint
            })
        df_resumen = pd.DataFrame(resumen).sort_values(by="Apellido y Nombre")

        year = datetime.date.today().year
        excel_name = f"FordFiorasi_Antecedentes_Base_{year}.xlsx"
        excel_path = out_dir / excel_name
        with pd.ExcelWriter(excel_path) as w:
            df.to_excel(w, sheet_name="Base completa", index=False)
            df_resumen.to_excel(w, sheet_name="Resumen por empleado", index=False)

        st.success("Procesamiento finalizado ✅")
        with open(excel_path, "rb") as f:
            data = f.read()
        st.download_button("📘 Descargar Excel", data, file_name=excel_name)

        zip_io = BytesIO()
        with zipfile.ZipFile(zip_io, "w", zipfile.ZIP_DEFLATED) as z:
            z.write(excel_path, excel_name)
            for root, dirs, files in os.walk(out_dir):
                for file in files:
                    fp = os.path.join(root, file)
                    z.write(fp, os.path.relpath(fp, out_dir))
        zip_io.seek(0)
        st.download_button("📦 Descargar ZIP completo (Excel + carpetas)", zip_io, file_name=f"FordFiorasi_Completo_{year}.zip")
