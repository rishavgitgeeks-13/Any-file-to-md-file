import os
import platform
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
    COM_AVAILABLE,
    SUPPORTED_EXTENSIONS,
)

try:
    import anthropic as _anthropic_module
except ImportError:
    _anthropic_module = None

st.set_page_config(page_title="Document → Markdown Converter", page_icon="📄", layout="centered")

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

with st.sidebar:
    st.header("Conversion Engine")

    st.markdown("**Tier 1 — markitdown**")
    st.success("Ready")

    st.markdown("**Tier 2 — OCR (pytesseract)**")
    if OCR_AVAILABLE:
        st.success("Ready")
    else:
        st.warning("Unavailable\n\nInstall pytesseract and the Tesseract binary.")

    st.markdown("**Tier 3 — Claude LLM**")
    if LLM_AVAILABLE and llm_client:
        st.success(f"Ready — {LLM_MODEL}")
    else:
        st.warning("Unavailable\n\nSet ANTHROPIC_API_KEY.")

    st.markdown("**Legacy formats (.doc / .ppt)**")
    if COM_AVAILABLE:
        st.success("Ready — Microsoft Office COM")
    elif platform.system() == "Windows":
        st.warning("Unavailable\n\nInstall Microsoft Office and pywin32.")
    else:
        st.info("Not available on Streamlit Cloud (Linux).\nPlease upload .docx or .pptx instead.")

    st.divider()
    st.caption("Files are converted locally and are not stored.")

st.title("📄 Document → Markdown")
st.caption("Upload a document and download a Markdown file.")

accepted = sorted(
    ext.lstrip(".")
    for ext in SUPPORTED_EXTENSIONS
    if platform.system() == "Windows" or ext not in {".doc", ".ppt"}
)

uploaded = st.file_uploader(
    "Choose a file",
    type=accepted,
    help=f"Supported: {', '.join(accepted)}",
)

if uploaded:
    col1, col2 = st.columns([3,1])
    col1.write(f"**{uploaded.name}** · {uploaded.size/1024:.1f} KB")

    if col2.button("Convert", type="primary", use_container_width=True):
        tmp_path = None
        try:
            suffix = Path(uploaded.name).suffix or ".bin"
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                tmp.write(uploaded.read())
                tmp_path = Path(tmp.name)

            with st.spinner("Converting..."):
                markdown_text, method = convert_file(tmp_path, converter, llm_client)

            labels = {
                "markitdown":"Tier 1 — markitdown",
                "ocr":"Tier 2 — OCR",
                "llm":f"Tier 3 — Claude ({LLM_MODEL})"
            }

            st.success(f"Converted via **{labels.get(method, method)}**")

            st.download_button(
                "⬇ Download Markdown",
                markdown_text.encode("utf-8"),
                file_name=f"{Path(uploaded.name).stem}.md",
                mime="text/markdown",
                use_container_width=True,
                type="primary",
            )

            t1, t2 = st.tabs(["Rendered preview", "Raw Markdown"])
            with t1:
                st.markdown(markdown_text)
            with t2:
                st.text_area("Markdown", markdown_text, height=400)

        except Exception as e:
            st.exception(e)

        finally:
            if tmp_path and tmp_path.exists():
                tmp_path.unlink()
