
import tempfile
from pathlib import Path

import streamlit as st

from convert_main import (
    convert_file,
    OCR_AVAILABLE,
    FITZ_AVAILABLE,
    SUPPORTED_EXTENSIONS,
    LLAMA_API_KEY,
)

from llama_parse import LlamaParse

st.set_page_config(
    page_title="Document → Markdown Converter",
    page_icon="📄",
    layout="centered",
)

@st.cache_resource
def get_llama_parser():
    if not LLAMA_API_KEY:
        return None

    return LlamaParse(
        api_key=LLAMA_API_KEY,
        result_type="markdown",
        verbose=True,
    )

llama_parser = get_llama_parser()

with st.sidebar:
    st.header("Conversion Engine")

    st.markdown("### Tier 1 — LlamaParse")
    if llama_parser:
        st.success("Ready")
    else:
        st.error("LLAMA_CLOUD_API_KEY not found")

    st.markdown("### Tier 2 — OCR")
    if OCR_AVAILABLE:
        st.success("Ready")
    else:
        st.warning("Install pytesseract and the Tesseract binary")

    st.divider()

    st.write("PyMuPDF:", "✅" if FITZ_AVAILABLE else "❌")
    st.write("OCR:", "✅" if OCR_AVAILABLE else "❌")

st.title("📄 Document → Markdown Converter")

st.caption(
    "Extract Markdown using LlamaParse with OCR fallback."
)

accepted = sorted(ext.lstrip(".") for ext in SUPPORTED_EXTENSIONS)

uploaded = st.file_uploader(
    "Choose a file",
    type=accepted,
)

if uploaded:

    st.write(f"**{uploaded.name}**")
    st.write(f"Size: {uploaded.size/1024:.1f} KB")

    if st.button("Convert", type="primary", use_container_width=True):

        suffix = Path(uploaded.name).suffix
        tmp_path = None

        try:

            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                tmp.write(uploaded.getbuffer())
                tmp_path = Path(tmp.name)

            with st.spinner("Converting document..."):
                markdown, method = convert_file(
                    file=tmp_path,
                    llama_parser=llama_parser,
                )

            labels = {
                "llamaparse": "LlamaParse",
                "ocr": "OCR",
            }

            st.success(f"Converted using **{labels.get(method, method)}**")

            st.download_button(
                "⬇ Download Markdown",
                markdown.encode("utf-8"),
                file_name=Path(uploaded.name).stem + ".md",
                mime="text/markdown",
                use_container_width=True,
            )

            tab1, tab2 = st.tabs(["Preview", "Markdown"])

            with tab1:
                st.markdown(markdown)

            with tab2:
                st.text_area(
                    "Markdown",
                    markdown,
                    height=500,
                    label_visibility="collapsed",
                )

        except Exception as e:
            st.error(str(e))

        finally:
            if tmp_path and tmp_path.exists():
                tmp_path.unlink()
