import os
import tempfile
from pathlib import Path

from dotenv import load_dotenv
load_dotenv()

import streamlit as st
from markitdown import MarkItDown

from convert_main import (
    convert_file,
    LLM_AVAILABLE,
    LLM_MODEL,
    OCR_AVAILABLE,
    FITZ_AVAILABLE,
    SUPPORTED_EXTENSIONS
)

try:
    import anthropic as _anthropic_module
except ImportError:
    _anthropic_module = None

# =====================================================
# PAGE CONFIG
# =====================================================

st.set_page_config(
    page_title="Document → Markdown Converter",
    page_icon="📄",
    layout="centered",
)

# =====================================================
# CACHED RESOURCES
# =====================================================

@st.cache_resource
def get_converter():
    return MarkItDown()


@st.cache_resource
def get_llm_client():
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if api_key and _anthropic_module:
        return _anthropic_module.Anthropic(api_key=api_key)
    return None


converter = get_converter()
llm_client = get_llm_client()

# =====================================================
# SIDEBAR — ENGINE STATUS
# =====================================================

with st.sidebar:
    st.header("Conversion Engine")

    st.markdown("**Tier 1 — markitdown**")
    st.success("Ready")

    st.markdown("**Tier 2 — OCR (pytesseract)**")
    if OCR_AVAILABLE:
        st.success("Ready")
    else:
        st.warning("Unavailable\n\nInstall `pytesseract` + the [Tesseract binary](https://github.com/UB-Mannheim/tesseract/wiki)")

    st.markdown("**Tier 3 — Claude LLM**")
    if LLM_AVAILABLE and llm_client:
        st.success(f"Ready — `{LLM_MODEL}`")
    else:
        st.warning("Unavailable\n\nSet the `ANTHROPIC_API_KEY` environment variable")

    st.divider()
    st.caption(
        "Files are converted locally and never stored. "
        "The LLM tier sends page images to Anthropic's API only when needed."
    )

# =====================================================
# MAIN UI
# =====================================================

st.title("📄 Document → Markdown")
st.caption(
    "Upload a document and download a clean `.md` file. "
    "Falls back to OCR and then Claude vision if standard extraction fails."
)

accepted = sorted(ext.lstrip(".") for ext in SUPPORTED_EXTENSIONS)
uploaded = st.file_uploader(
    "Choose a file",
    type=accepted,
    help=f"Supported: {', '.join(accepted)}",
)

if uploaded:
    col1, col2 = st.columns([3, 1])
    col1.write(f"**{uploaded.name}**  ·  {uploaded.size / 1024:.1f} KB")

    if col2.button("Convert", type="primary", use_container_width=True):
        suffix = Path(uploaded.name).suffix or ".bin"
        tmp_path = None

        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                tmp.write(uploaded.read())
                tmp_path = Path(tmp.name)

            with st.spinner("Converting…"):
                markdown_text, method = convert_file(tmp_path, converter, llm_client)

            METHOD_LABELS = {
                "markitdown": "Tier 1 — markitdown",
                "ocr":        "Tier 2 — OCR (pytesseract)",
                "llm":        f"Tier 3 — Claude LLM ({LLM_MODEL})",
            }
            st.success(f"Converted via **{METHOD_LABELS.get(method, method)}**")

            # --- Download ---
            output_name = Path(uploaded.name).stem + ".md"
            st.download_button(
                label="⬇ Download Markdown",
                data=markdown_text.encode("utf-8"),
                file_name=output_name,
                mime="text/markdown",
                type="primary",
                use_container_width=True,
            )

            # --- Preview tabs ---
            tab_rendered, tab_raw = st.tabs(["Rendered preview", "Raw Markdown"])
            with tab_rendered:
                st.markdown(markdown_text)
            with tab_raw:
                st.text_area(
                    label="raw",
                    value=markdown_text,
                    height=400,
                    label_visibility="collapsed",
                )

        except Exception as e:
            st.error(f"**Conversion failed:** {e}")

        finally:
            if tmp_path and tmp_path.exists():
                tmp_path.unlink()
