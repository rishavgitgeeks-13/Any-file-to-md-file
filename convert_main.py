
import io
import os
from pathlib import Path
from dotenv import load_dotenv
from llama_parse import LlamaParse

load_dotenv()

LLAMA_API_KEY = os.getenv("LLAMA_CLOUD_API_KEY")

try:
    import fitz
    FITZ_AVAILABLE = True
except ImportError:
    FITZ_AVAILABLE = False

try:
    from PIL import Image
    import pytesseract
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

INPUT_FOLDER = r"C:\Users\rishav.patel\Documents\ConversionTool\IT Policy"
OUTPUT_FOLDER = r"C:\Users\rishav.patel\Documents\ConversionTool\md.files"

PDF_DPI = 200
MIN_TEXT_LENGTH = 300

SUPPORTED_EXTENSIONS = {
    ".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".gif", ".webp"
}

def is_content_poor(text: str) -> bool:
    if not text:
        return True
    text = text.strip()
    return len(text) < MIN_TEXT_LENGTH or len(text.split()) < 50

def pdf_to_png(pdf_path: Path):
    pages = []
    doc = fitz.open(str(pdf_path))
    mat = fitz.Matrix(PDF_DPI / 72, PDF_DPI / 72)
    for page in doc:
        pix = page.get_pixmap(matrix=mat)
        pages.append(pix.tobytes("png"))
    doc.close()
    return pages

def image_to_png(image_path: Path):
    buf = io.BytesIO()
    Image.open(image_path).save(buf, format="PNG")
    return [buf.getvalue()]

def gather_images(path: Path):
    if path.suffix.lower() == ".pdf":
        if not FITZ_AVAILABLE:
            raise RuntimeError("PyMuPDF is required.")
        return pdf_to_png(path)
    return image_to_png(path)

def ocr_extract(images):
    output = []
    for i, img in enumerate(images, 1):
        text = pytesseract.image_to_string(Image.open(io.BytesIO(img))).strip()
        if text:
            output.append(f"## Page {i}\n\n{text}")
    return "\n\n".join(output)

def llamaparse_extract(path: Path, parser):
    docs = parser.load_data(str(path))
    return "\n\n".join(doc.text for doc in docs)

def convert_file(path: Path, llama_parser=None):
    if path.suffix.lower() == ".pdf":
        try:
            md = llamaparse_extract(path, parser)
            if not is_content_poor(md):
                return md, "llamaparse"
            print("  [LLAMAPARSE] Poor extraction, falling back to OCR...")
        except Exception as e:
            print(f"  [LLAMAPARSE] {e}")

    if not OCR_AVAILABLE:
        raise RuntimeError("OCR unavailable. Install pillow and pytesseract.")

    images = gather_images(path)
    md = ocr_extract(images)
    if not is_content_poor(md):
        return md, "ocr"

    raise RuntimeError("Extraction failed.")

def main():
    parser = None
    if LLAMA_API_KEY:
        parser = LlamaParse(
            api_key=LLAMA_API_KEY,
            result_type="markdown",
            verbose=True,
        )
    else:
        print("WARNING: LLAMA_CLOUD_API_KEY not found.")

    input_dir = Path(INPUT_FOLDER)
    output_dir = Path(OUTPUT_FOLDER)
    output_dir.mkdir(parents=True, exist_ok=True)

    stats = {"llamaparse":0,"ocr":0,"failed":0}

    for file in input_dir.rglob("*"):
        if not file.is_file():
            continue
        if file.suffix.lower() not in SUPPORTED_EXTENSIONS:
            continue

        out = output_dir / file.relative_to(input_dir).with_suffix(".md")
        out.parent.mkdir(parents=True, exist_ok=True)

        print(f"[PROCESSING] {file.name}")

        try:
            text, method = convert_file(file, parser)
            out.write_text(text, encoding="utf-8")
            stats[method]+=1
            print(f"[OK:{method.upper()}] {file.name}")
        except Exception as e:
            stats["failed"]+=1
            print(f"[FAILED] {file.name}: {e}")

    print("\\nSummary")
    print(f"LlamaParse : {stats['llamaparse']}")
    print(f"OCR        : {stats['ocr']}")
    print(f"Failed     : {stats['failed']}")

if __name__ == "__main__":
    main()
