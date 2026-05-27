import streamlit as st
import google.generativeai as genai
import io
from pathlib import Path

import docx as python_docx
import pdfplumber


# --- Helpers: extracción de texto ---

def extract_text(uploaded_file) -> str:
    """Extrae texto de un archivo subido por el usuario."""
    name = uploaded_file.name.lower()
    raw = uploaded_file.read()

    if name.endswith(".txt"):
        for encoding in ("utf-8", "latin-1", "cp1252"):
            try:
                return raw.decode(encoding)
            except UnicodeDecodeError:
                continue
        return raw.decode("utf-8", errors="replace")

    if name.endswith(".docx"):
        doc = python_docx.Document(io.BytesIO(raw))
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())

    if name.endswith(".pdf"):
        pages = []
        with pdfplumber.open(io.BytesIO(raw)) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    pages.append(text)
        return "\n\n".join(pages)

    return ""


def extract_text_from_path(path: Path) -> str:
    """Extrae texto de un archivo local (instrucciones, manual)."""
    name = path.name.lower()

    if name.endswith(".txt"):
        for encoding in ("utf-8", "latin-1", "cp1252"):
            try:
                return path.read_text(encoding=encoding).strip()
            except UnicodeDecodeError:
                continue
        return path.read_text(encoding="utf-8", errors="replace").strip()

    if name.endswith(".docx"):
        doc = python_docx.Document(str(path))
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())

    if name.endswith(".pdf"):
        pages = []
        with pdfplumber.open(str(path)) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    pages.append(text)
        return "\n\n".join(pages)

    return ""


# --- Helpers: archivos de referencia ---

BASE_DIR = Path(__file__).parent


def load_instructions() -> str:
    path = BASE_DIR / "instrucciones.txt"
    if path.exists():
        return extract_text_from_path(path)
    return ""


def load_manual() -> tuple[str, str]:
    for candidate in ("manual.txt", "manual.docx", "manual.pdf"):
        path = BASE_DIR / candidate
        if path.exists():
            return extract_text_from_path(path), candidate
    return "", ""
    if path.exists():
        return extract_text_from_path(path)
    return ""


def load_manual() -> tuple[str, str]:
    """
    Busca el manual en la carpeta del proyecto.
    Acepta: manual.txt, manual.docx, manual.pdf
    Devuelve (contenido, nombre_archivo) o ("", "") si no existe.
    """
    for candidate in ("manual.txt", "manual.docx", "manual.pdf"):
        path = Path(candidate)
        if path.exists():
            return extract_text_from_path(path), candidate
    return "", ""


def get_api_key() -> str:
    try:
        return st.secrets["GOOGLE_API_KEY"]
    except Exception:
        return ""


# --- Análisis con Gemini ---

def stream_report(contract_text: str, instructions: str, manual_text: str, api_key: str):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-2.0-flash")

    manual_block = ""
    if manual_text.strip():
        manual_block = f"""---

MANUAL DE REFERENCIA DE LA EMPRESA (usa este documento para comparar y clasificar cada cláusula):
{manual_text}

"""

    prompt = f"""INSTRUCCIONES DE REVISIÓN:
{instructions}

{manual_block}---

CONTRATO A REVISAR:
{contract_text}

---

Genera el informe de revisión siguiendo exactamente las instrucciones proporcionadas.
Usa formato Markdown con secciones bien definidas, negritas para conceptos clave y listas cuando proceda.
Escribe en español."""

    response = model.generate_content(prompt, stream=True)
    for chunk in response:
        if chunk.text:
            yield chunk.text


# =====================
# --- Interfaz ---
# =====================

st.set_page_config(
    page_title="PEPA - Revisión preliminar Contratos SaaS",
    page_icon="⚖️",
    layout="wide",
)

st.title("⚖️ PEPA - Revisión preliminar Contratos SaaS")
st.caption("Sube un Contrato SaaS y obtén un informe de revisión automático basado en las instrucciones corporativas de PEPA.")
st.caption("El presente proyecto es un proyecto educativo y no sustituye al asesoramiento legal.")

# --- Barra lateral ---
with st.sidebar:
    st.header("Cómo usar la app")
    st.markdown("""
1. Coloca `instrucciones.txt` y `manual.txt` (o `.docx`/`.pdf`) en la carpeta del proyecto.
2. **Sube** el contrato (`.txt`, `.docx` o `.pdf`).
3. Pulsa **Generar informe**.
4. Descarga el resultado si lo necesitas.
""")
    st.divider()
    st.header("Estado de archivos de referencia")

    instructions_preview = load_instructions()
    manual_text_sidebar, manual_name_sidebar = load_manual()

    if instructions_preview:
        st.success("✅ instrucciones.txt cargado")
        with st.expander("Ver instrucciones"):
            st.text(instructions_preview[:600] + ("..." if len(instructions_preview) > 600 else ""))
    else:
        st.error("❌ `instrucciones.txt` no encontrado")

    if manual_text_sidebar:
        st.success(f"✅ {manual_name_sidebar} cargado ({len(manual_text_sidebar):,} caracteres)")
        with st.expander("Ver extracto del manual"):
            st.text(manual_text_sidebar[:600] + ("..." if len(manual_text_sidebar) > 600 else ""))
    else:
        st.warning("⚠️ Manual no encontrado. La revisión usará estándares genéricos de mercado.")

    st.divider()
    st.header("Configuración")
    st.markdown("""
**API Key de Google:**
Debe estar definida en `.streamlit/secrets.toml`:
```toml
GOOGLE_API_KEY = "AIzaSy..."
```
[Obtén tu clave gratuita →](https://aistudio.google.com/apikey)
""")

# --- Comprobación de API key ---
api_key = get_api_key()
if not api_key:
    st.error(
        "❌ No se encontró la API key de Google. "
        "Añade tu clave en la configuración de Streamlit Cloud (Secrets). "
        "Consulta la barra lateral para más información."
    )
    st.stop()

# --- Subida de archivo ---
uploaded_file = st.file_uploader(
    "Sube el contrato a revisar",
    type=["txt", "docx", "pdf"],
    label_visibility="visible",
)

if uploaded_file is None:
    st.info("👆 Sube un documento para comenzar.")
    st.stop()

# --- Extracción de texto ---
with st.spinner("Leyendo documento..."):
    contract_text = extract_text(uploaded_file)

if not contract_text.strip():
    st.error("No se pudo extraer texto del documento. Verifica que el archivo no esté vacío ni protegido.")
    st.stop()

col1, col2 = st.columns([3, 1])
with col1:
    st.success(f"✅ Documento cargado: **{uploaded_file.name}** ({len(contract_text):,} caracteres)")
with col2:
    with st.expander("Ver texto extraído"):
        st.text_area(
            "Contenido",
            contract_text[:4000] + ("\n\n[...]" if len(contract_text) > 4000 else ""),
            height=250,
            disabled=True,
            label_visibility="collapsed",
        )

# --- Botón de análisis ---
instructions = load_instructions()
manual_text, manual_name = load_manual()

if not instructions:
    st.warning("⚠️ El archivo `instrucciones.txt` está vacío. Añade tus instrucciones antes de continuar.")
    st.stop()

if manual_text:
    st.info(f"📘 Manual de referencia cargado: **{manual_name}** ({len(manual_text):,} caracteres)")
else:
    st.warning("⚠️ No se encontró manual de referencia. La revisión se basará en estándares genéricos de mercado.")

if st.button("🔍 Generar informe de revisión", type="primary", use_container_width=True):
    st.divider()
    st.subheader("📊 Informe de Revisión")

    report_area = st.empty()
    full_report = ""

    try:
        for chunk in stream_report(contract_text, instructions, manual_text, api_key):
            full_report += chunk
            report_area.markdown(full_report + "▌")
        report_area.markdown(full_report)

    except Exception as e:
        mensaje = str(e).lower()
        if "api key" in mensaje or "permission" in mensaje or "invalid" in mensaje:
            st.error("❌ API key incorrecta o no válida. Revisa el valor en los Secrets de Streamlit Cloud.")
        elif "quota" in mensaje or "resource" in mensaje or "limit" in mensaje:
            st.error("❌ Límite de uso de Google AI alcanzado. Espera unos minutos e inténtalo de nuevo.")
        else:
            st.error(f"❌ Error inesperado: {e}")
        st.stop()

    st.divider()
    st.download_button(
        label="⬇️ Descargar informe (.md)",
        data=full_report,
        file_name=f"informe_{Path(uploaded_file.name).stem}.md",
        mime="text/markdown",
        use_container_width=True,
    )
