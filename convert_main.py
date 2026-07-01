import io
import os
import base64
from pathlib import Path
from dotenv import load_dotenv
from markitdown import MarkItDown

load_dotenv()

# =====================================================
# OPTIONAL DEPENDENCIES
# =====================================================

try:
    import fitz  # pymupdf — PDF page rendering
    FITZ_AVAILABLE = True
except ImportError:
    FITZ_AVAILABLE = False

try:
    from PIL import Image as PILImage
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    import pytesseract
    OCR_AVAILABLE = PIL_AVAILABLE  # tesseract needs PIL too
except ImportError:
    OCR_AVAILABLE = False

try:
    import anthropic as _anthropic_module
    _ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
    LLM_AVAILABLE = bool(_ANTHROPIC_API_KEY)
except ImportError:
    _anthropic_module = None
    LLM_AVAILABLE = False
    _ANTHROPIC_API_KEY = ""

# =====================================================
# CONFIGURATION
# =====================================================

INPUT_FOLDER = r"C:\Users\rishav.patel\Documents\ConversionTool\IT Policy"
OUTPUT_FOLDER = r"C:\Users\rishav.patel\Documents\ConversionTool\md.files"

LLM_MODEL = "claude-opus-4-6"
MIN_TEXT_LENGTH = 100   # chars below this → treat extraction as failed
PDF_DPI = 200           # resolution for PDF-to-image rendering

SUPPORTED_EXTENSIONS = {
    ".pdf", ".doc", ".docx", ".ppt", ".pptx",
    ".txt", ".png", ".jpg", ".jpeg", ".gif",
    ".bmp", ".tiff", ".webp", ".html", ".htm",
}

# markitdown cannot meaningfully extract text from raster images;
# skip tier-1 for these and go straight to OCR/LLM.
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".webp"}

# =====================================================
# HELPERS
# =====================================================

def is_content_poor(text: str) -> bool:
    return not text or len(text.strip()) < MIN_TEXT_LENGTH


def _pdf_to_png_bytes(pdf_path: Path, dpi: int = PDF_DPI) -> list[bytes]:
    """Render every page of a PDF as PNG bytes using pymupdf."""
    pages = []
    try:
        doc = fitz.open(str(pdf_path))
        zoom = dpi / 72
        mat = fitz.Matrix(zoom, zoom)
        for page in doc:
            pix = page.get_pixmap(matrix=mat)
            pages.append(pix.tobytes("png"))
        doc.close()
    except Exception as e:
        print(f"    [PDF-RENDER] {e}")
    return pages


def _image_file_to_png_bytes(image_path: Path) -> list[bytes]:
    """Load a raster image file and return it as a list of one PNG bytes blob."""
    if not PIL_AVAILABLE:
        return []
    try:
        buf = io.BytesIO()
        PILImage.open(str(image_path)).save(buf, format="PNG")
        return [buf.getvalue()]
    except Exception as e:
        print(f"    [IMG-LOAD] {e}")
        return []


def _gather_images(file: Path) -> list[bytes]:
    """Return page images (as PNG bytes) for a PDF or image file."""
    ext = file.suffix.lower()
    if ext == ".pdf" and FITZ_AVAILABLE:
        return _pdf_to_png_bytes(file)
    if ext in IMAGE_EXTENSIONS:
        return _image_file_to_png_bytes(file)
    return []


def _ocr_pages(png_pages: list[bytes]) -> str:
    """Run pytesseract on a list of PNG byte blobs; return joined Markdown."""
    parts = []
    for i, data in enumerate(png_pages, 1):
        try:
            img = PILImage.open(io.BytesIO(data))
            text = pytesseract.image_to_string(img).strip()
            if text:
                parts.append(f"## Page {i}\n\n{text}")
        except Exception as e:
            print(f"    [OCR page {i}] {e}")
    return "\n\n".join(parts)


def _llm_extract_pages(png_pages: list[bytes], client, source_name: str) -> str:
    """Send page images to Claude and request Markdown extraction."""
    content = []
    for data in png_pages:
        b64 = base64.standard_b64encode(data).decode("utf-8")
        content.append({
            "type": "image",
            "source": {"type": "base64", "media_type": "image/png", "data": b64},
        })
    content.append({
        "type": "text",
        "text": (
            f"These images are pages from the document: {source_name}\n\n"
            "Extract all text and content and format it as well-structured Markdown. "
            "Preserve headings, lists, tables, and paragraph layout faithfully. "
            "Return only the Markdown — no preamble or explanation."
        ),
    })

    response = client.messages.create(
        model=LLM_MODEL,
        max_tokens=8192,
        messages=[{"role": "user", "content": content}],
    )
    parts = [block.text for block in response.content if block.type == "text"]
    return "\n\n".join(parts)


# =====================================================
# THREE-TIER CONVERSION PIPELINE
# =====================================================

def convert_file(
    file: Path, converter: MarkItDown, llm_client=None
) -> tuple[str, str]:
    """
    Convert a file to Markdown.

    Returns (markdown_text, method) where method is one of:
      'markitdown', 'ocr', 'llm'

    Raises RuntimeError if all tiers fail or are unavailable.
    """
    ext = file.suffix.lower()
    skip_markitdown = ext in IMAGE_EXTENSIONS

    # --------------------------------------------------
    # Tier 1 — markitdown
    # --------------------------------------------------
    if not skip_markitdown:
        try:
            result = converter.convert(str(file))
            text = result.text_content or ""
            if not is_content_poor(text):
                return text, "markitdown"
            print(
                f"    [TIER1] insufficient ({len(text.strip())} chars), escalating to OCR..."
            )
        except Exception as e:
            print(f"    [TIER1] markitdown failed ({e}), escalating to OCR...")

    # --------------------------------------------------
    # Gather images (shared by Tier 2 and Tier 3)
    # --------------------------------------------------
    png_pages = _gather_images(file)

    # If we can't produce images, there's nothing left to try
    if not png_pages:
        if ext in IMAGE_EXTENSIONS:
            raise RuntimeError(
                "Could not load image file — is Pillow installed? (pip install pillow)"
            )
        if ext == ".pdf" and not FITZ_AVAILABLE:
            raise RuntimeError(
                "PDF rendering requires pymupdf — run: pip install pymupdf"
            )
        if ext == ".pdf":
            raise RuntimeError(
                "PDF could not be rendered to images. "
                "The file may be corrupted, password-protected, or empty."
            )
        # docx, pptx, etc. — markitdown is the only tier available
        raise RuntimeError(
            f"markitdown failed on {ext} and this format cannot be rendered to "
            "images for OCR/LLM fallback. Check the file is not corrupted."
        )

    # --------------------------------------------------
    # Tier 2 — OCR via pytesseract
    # --------------------------------------------------
    if OCR_AVAILABLE:
        text = _ocr_pages(png_pages)
        if not is_content_poor(text):
            return text, "ocr"
        print(
            f"    [TIER2] OCR insufficient ({len(text.strip())} chars), escalating to LLM..."
        )
    else:
        print("    [TIER2] pytesseract unavailable, skipping OCR...")

    # --------------------------------------------------
    # Tier 3 — Anthropic Claude vision
    # --------------------------------------------------
    if LLM_AVAILABLE and llm_client is not None:
        text = _llm_extract_pages(png_pages, llm_client, file.name)
        if not is_content_poor(text):
            return text, "llm"
        raise RuntimeError("LLM extraction returned insufficient content")

    if not LLM_AVAILABLE:
        raise RuntimeError(
            "OCR also failed. Set ANTHROPIC_API_KEY in .env to enable LLM fallback."
        )
    raise RuntimeError("All extraction tiers failed")


# =====================================================
# MAIN
# =====================================================

def main():
    input_dir = Path(INPUT_FOLDER)
    output_dir = Path(OUTPUT_FOLDER)

    if not input_dir.exists():
        raise FileNotFoundError(f"Input folder not found: {INPUT_FOLDER}")
    output_dir.mkdir(parents=True, exist_ok=True)

    converter = MarkItDown()

    llm_client = None
    if LLM_AVAILABLE:
        llm_client = _anthropic_module.Anthropic(api_key=_ANTHROPIC_API_KEY)
        print(f"LLM fallback  : ENABLED  (model: {LLM_MODEL})")
    else:
        print("LLM fallback  : DISABLED (set ANTHROPIC_API_KEY to enable)")

    ocr_status = "ENABLED" if OCR_AVAILABLE else "DISABLED (install pytesseract + Tesseract binary)"
    print(f"OCR fallback  : {ocr_status}")
    pdf_status = "ENABLED" if FITZ_AVAILABLE else "DISABLED (install pymupdf)"
    print(f"PDF rendering : {pdf_status}")
    print(f"\nInput  : {INPUT_FOLDER}")
    print(f"Output : {OUTPUT_FOLDER}\n")

    success = 0
    failed = 0
    skipped = 0
    unsupported = 0
    method_counts = {"markitdown": 0, "ocr": 0, "llm": 0}

    for file in input_dir.rglob("*"):
        if not file.is_file():
            continue

        ext = file.suffix.lower()
        rel = file.relative_to(input_dir)

        if ext == ".md":
            skipped += 1
            print(f"[SKIPPED]     {rel}")
            continue

        if ext not in SUPPORTED_EXTENSIONS:
            unsupported += 1
            print(f"[UNSUPPORTED] {rel}")
            continue

        output_file = output_dir / rel.parent / f"{file.stem}.md"
        output_file.parent.mkdir(parents=True, exist_ok=True)

        if output_file.exists():
            skipped += 1
            print(f"[EXISTS]      {rel}")
            continue

        print(f"[CONVERTING]  {rel}")
        try:
            text, method = convert_file(file, converter, llm_client)
            output_file.write_text(text, encoding="utf-8")
            success += 1
            method_counts[method] += 1
            print(f"[OK:{method.upper():<11}] {rel}")
        except Exception as e:
            failed += 1
            print(f"[FAILED]      {rel}")
            print(f"  Reason: {e}")

    print("\n" + "=" * 55)
    print("Conversion Summary")
    print("=" * 55)
    print(f"Converted    : {success}")
    print(f"  markitdown : {method_counts['markitdown']}")
    print(f"  ocr        : {method_counts['ocr']}")
    print(f"  llm        : {method_counts['llm']}")
    print(f"Skipped      : {skipped}")
    print(f"Unsupported  : {unsupported}")
    print(f"Failed       : {failed}")
    print("=" * 55)


if __name__ == "__main__":
    main()
