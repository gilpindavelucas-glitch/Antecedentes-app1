# ⚖️ Ford Fiorasi – Procesador de Antecedentes Disciplinarios

Aplicación web institucional para automatizar la gestión de antecedentes disciplinarios (llamados de atención, apercibimientos, descargos, etc.) del personal de Ford Fiorasi.

## 🚀 Funcionalidades principales
- Procesa archivos `.docx` y `.pdf` automáticamente.
- OCR en español integrado (sin necesidad de instalación).
- Detecta nombre, fecha, tipo de antecedente, descripción y si hay descargo.
- Genera:
  - Excel con base completa y resumen por empleado.
  - Carpetas por empleado con sus archivos.
  - ZIP completo descargable.
- Selector de color institucional (rueda de ajustes).
- Branding Ford Fiorasi con logo y colores corporativos.

## 💻 Cómo usar
1. Sube todos los archivos desde el navegador.
2. Presiona **“Procesar antecedentes”**.
3. Descarga el **Excel** o el **ZIP completo**.

## 🌐 Despliegue en Streamlit Cloud
1. Crea un nuevo repositorio en GitHub.
2. Sube los archivos incluidos.
3. Entra a [https://share.streamlit.io](https://share.streamlit.io) y conecta el repositorio.
4. Espera a que instale dependencias — la app estará lista en minutos.
