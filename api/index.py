from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pathlib import Path
import tempfile, zipfile, uuid, os
from datetime import datetime

from pypdf import PdfReader, PdfWriter
from PIL import Image, ImageFile

ImageFile.LOAD_TRUNCATED_IMAGES = True

app = FastAPI(title="Universal File Repair API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

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

MAX_BYTES = 4 * 1024 * 1024  # 4 MB Vercel Hobby limit


def repair_pdf(data: bytes) -> bytes:
    if not data.startswith(b"%PDF-"):
        idx = data.find(b"%PDF-")
        data = data[idx:] if idx > 0 else b"%PDF-1.4\n" + data

    if b"%%EOF" not in data[-1024:]:
        data = data.rstrip() + b"\n%%EOF\n"

    tmp_in = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    tmp_in.write(data); tmp_in.close()
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
        return data
    finally:
        os.unlink(tmp_in.name)
        if os.path.exists(tmp_out.name):
            os.unlink(tmp_out.name)


def repair_image(data: bytes, ext: str) -> bytes:
    magic = MAGIC.get(ext)
    if magic and not data.startswith(magic):
        idx = data.find(magic)
        if idx > 0:
            data = data[idx:]

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
    tmp.write(data); tmp.close()
    out = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
    out.close()

    try:
        img = Image.open(tmp.name)
        img.load()
        fmt = {
            ".png": "PNG", ".jpg": "JPEG", ".jpeg": "JPEG",
            ".gif": "GIF", ".bmp": "BMP", ".tiff": "TIFF",
        }.get(ext, "PNG")
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


def repair_zip_based(data: bytes, ext: str) -> bytes:
    tmp_in = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
    tmp_in.write(data); tmp_in.close()
    tmp_out = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
    tmp_out.close()

    try:
        with zipfile.ZipFile(tmp_in.name, "r") as zin:
            with zipfile.ZipFile(tmp_out.name, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    try:
                        zout.writestr(item, zin.read(item.filename))
                    except Exception:
                        pass
        return Path(tmp_out.name).read_bytes()
    except Exception as e:
        raise HTTPException(400, f"ZIP container repair failed: {e}")
    finally:
        os.unlink(tmp_in.name)
        if os.path.exists(tmp_out.name):
            os.unlink(tmp_out.name)


def repair_text(data: bytes, ext: str) -> bytes:
    data = data.replace(b"\x00", b"")
    text = None
    for enc in ("utf-8", "utf-16", "latin-1", "cp1252"):
        try:
            text = data.decode(enc); break
        except UnicodeDecodeError:
            continue
    if text is None:
        text = data.decode("utf-8", errors="replace")
    if ext == ".tex":
        ob, cb = text.count("{"), text.count("}")
        if ob > cb:
            text += "\n" + "}" * (ob - cb)
        if "\\begin{document}" in text and not text.rstrip().endswith("\\end{document}"):
            text = text.rstrip() + "\n\\end{document}\n"
    return text.encode("utf-8")


def repair_generic(data: bytes, ext: str) -> bytes:
    magic = MAGIC.get(ext)
    if magic:
        idx = data.find(magic)
        if 0 < idx < 1_000_000:
            data = data[idx:]
    return data


@app.post("/api/repair")
async def repair(file: UploadFile = File(...)):
    raw = await file.read()

    if len(raw) > MAX_BYTES:
        raise HTTPException(413, f"File too large. Max {MAX_BYTES // 1024 // 1024} MB on free tier.")

    ext = Path(file.filename or "").suffix.lower()
    if not ext:
        raise HTTPException(400, "File must have an extension.")

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

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
    tmp.write(fixed); tmp.close()
    stem = Path(file.filename).stem

    return FileResponse(
        tmp.name,
        media_type="application/octet-stream",
        filename=f"{stem}_repaired{ext}",
    )


@app.get("/api/health")
def health():
    return {"status": "ok", "time": datetime.utcnow().isoformat()}
