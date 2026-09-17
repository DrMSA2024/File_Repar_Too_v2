"""
Universal File Repair API
FastAPI backend for repairing corrupted files.
"""
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pathlib import Path
import tempfile
import zipfile
import os
from datetime import datetime

from pypdf import PdfReader, PdfWriter
from PIL import Image, ImageFile

# Allow truncated images to be loaded (helps repair)
ImageFile.LOAD_TRUNCATED_IMAGES = True

app = FastAPI(title="Universal File Repair API")

# CORS - allow frontend to call this API
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition"],
)

# Magic bytes (file signatures) for common formats
MAGIC = {
    ".pdf": b"%PDF-",
    ".png": b"\x89PNG\r\n\x1a\n",
    ".jpg": b"\xff\xd8\xff",
    ".jpeg": b"\xff\xd8\xff",
    ".gif": b"GIF8",
    ".docx": b"PK\x03\x04",
    ".pptx": b"PK\x03\x04",
    ".xlsx": b"PK\x03\x04",
    ".zip": b"PK\x03\x04",
    ".doc": b"\xd0\xcf\x11\xe0",
    ".ppt": b"\xd0\xcf\x11\xe0",
}

# Max file size: 4 MB (Vercel Hobby tier limit)
MAX_BYTES = 4 * 1024 * 1024


# ============================================================
# PDF REPAIR
# ============================================================
def repair_pdf(data: bytes) -> bytes:
    """Repair PDF: fix header, add EOF, rebuild object tree."""
    # Fix missing header
    if not data.startswith(b"%PDF-"):
        idx = data.find(b"%PDF-")
        if idx > 0:
            data = data[idx:]
        else:
            data = b"%PDF-1.4\n" + data

    # Ensure %%EOF at end
    if b"%%EOF" not in data[-1024:]:
        data = data.rstrip() + b"\n%%EOF\n"

    # Try to rebuild PDF structure with pypdf
    tmp_in = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    tmp_in.write(data)
    tmp_in.close()

    tmp_out = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    tmp_out.close()

    try:
        reader = PdfReader(tmp_in.name, strict=False)
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        with open(tmp_out.name, "wb") as f:
            writer.write(f)
        return Path(tmp_out.name).read_bytes()
    except Exception:
        # If rebuild fails, return header-fixed version
        return data
    finally:
        os.unlink(tmp_in.name)
        if os.path.exists(tmp_out.name):
            os.unlink(tmp_out.name)


# ============================================================
# IMAGE REPAIR (PNG, JPG, GIF, BMP, TIFF)
# ============================================================
def repair_image(data: bytes, ext: str) -> bytes:
    """Repair image: strip junk before header, re-encode with Pillow."""
    # Strip bytes before magic header
    magic = MAGIC.get(ext)
    if magic and not data.startswith(magic):
        idx = data.find(magic)
        if idx > 0:
            data = data[idx:]

    # Write to temp file
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
    tmp.write(data)
    tmp.close()

    out = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
    out.close()

    try:
        img = Image.open(tmp.name)
        img.load()

        fmt = {
            ".png": "PNG",
            ".jpg": "JPEG",
            ".jpeg": "JPEG",
            ".gif": "GIF",
            ".bmp": "BMP",
            ".tiff": "TIFF",
        }.get(ext, "PNG")

        # JPEG doesn't support transparency
        if fmt == "JPEG" and img.mode in ("RGBA", "P"):
            img = img.convert("RGB")

        img.save(out.name, format=fmt)
        return Path(out.name).read_bytes()
    except Exception as e:
        raise HTTPException(400, f"Image repair failed: {e}")
    finally:
        os.unlink(tmp.name)
        if os.path.exists(out.name):
            os.unlink(out.name)


# ============================================================
# ZIP-BASED REPAIR (DOCX, PPTX, XLSX, ZIP, ODF)
# ============================================================
def repair_zip_based(data: bytes, ext: str) -> bytes:
    """Repair ZIP containers (Office docs, EPUB, etc.)."""
    tmp_in = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
    tmp_in.write(data)
    tmp_in.close()

    tmp_out = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
    tmp_out.close()

    try:
        with zipfile.ZipFile(tmp_in.name, "r") as zin:
            with zipfile.ZipFile(tmp_out.name, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    try:
                        zout.writestr(item, zin.read(item.filename))
                    except Exception:
                        # Skip corrupt entries, keep going
                        pass
        return Path(tmp_out.name).read_bytes()
    except Exception as e:
        raise HTTPException(400, f"ZIP container repair failed: {e}")
    finally:
        os.unlink(tmp_in.name)
        if os.path.exists(tmp_out.name):
            os.unlink(tmp_out.name)


# ============================================================
# TEXT REPAIR (TXT, TEX, MD, CSV, JSON, XML, HTML)
# ============================================================
def repair_text(data: bytes, ext: str) -> bytes:
    """Repair text files: strip nulls, fix encoding, balance braces."""
    # Strip null bytes
    data = data.replace(b"\x00", b"")

    # Try different encodings
    text = None
    for enc in ("utf-8", "utf-16", "latin-1", "cp1252"):
        try:
            text = data.decode(enc)
            break
        except UnicodeDecodeError:
            continue

    if text is None:
        text = data.decode("utf-8", errors="replace")

    # LaTeX-specific fixes
    if ext == ".tex":
        open_b = text.count("{")
        close_b = text.count("}")
        if open_b > close_b:
            text += "\n" + "}" * (open_b - close_b)

        if "\\begin{document}" in text and not text.rstrip().endswith("\\end{document}"):
            text = text.rstrip() + "\n\\end{document}\n"

    return text.encode("utf-8")


# ============================================================
# GENERIC REPAIR (fallback for any file)
# ============================================================
def repair_generic(data: bytes, ext: str) -> bytes:
    """Generic repair: trim to known header."""
    magic = MAGIC.get(ext)
    if magic:
        idx = data.find(magic)
        if 0 < idx < 1_000_000:
            data = data[idx:]
    return data


# ============================================================
# MAIN ENDPOINT
# ============================================================
@app.post("/api/repair")
async def repair(file: UploadFile = File(...)):
    """Receive file, repair, return repaired version."""
    raw = await file.read()

    if len(raw) > MAX_BYTES:
        raise HTTPException(
            413,
            f"File too large. Max {MAX_BYTES // 1024 // 1024} MB on free tier."
        )

    ext = Path(file.filename or "").suffix.lower()
    if not ext:
        raise HTTPException(400, "File must have an extension.")

    # Route to correct repair function
    try:
        if ext == ".pdf":
            fixed = repair_pdf(raw)
        elif ext in (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff"):
            fixed = repair_image(raw, ext)
        elif ext in (".docx", ".pptx", ".xlsx", ".zip", ".odt", ".ods", ".odp"):
            fixed = repair_zip_based(raw, ext)
        elif ext in (".txt", ".tex", ".md", ".csv", ".json", ".xml", ".html"):
            fixed = repair_text(raw, ext)
        else:
            fixed = repair_generic(raw, ext)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, f"Repair failed: {e}")

    # Save to temp file and return with proper download headers
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
    tmp.write(fixed)
    tmp.close()

    stem = Path(file.filename).stem
    download_name = f"{stem}_repaired{ext}"

    return FileResponse(
        tmp.name,
        media_type="application/octet-stream",
        filename=download_name,
        headers={
            "Content-Disposition": f'attachment; filename="{download_name}"',
            "Access-Control-Expose-Headers": "Content-Disposition",
        },
    )


# ============================================================
# HEALTH CHECK
# ============================================================
@app.get("/api/health")
def health():
    return {"status": "ok", "time": datetime.utcnow().isoformat()}


# ============================================================
# ROOT
# ============================================================
@app.get("/")
def root():
    return {"message": "Universal File Repair API", "endpoint": "/api/repair"}
