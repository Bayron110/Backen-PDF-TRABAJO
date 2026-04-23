from fastapi import FastAPI, UploadFile, File, HTTPException, Form
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pathlib import Path
from pypdf import PdfReader, PdfWriter
import shutil
import subprocess
import uuid
import os

app = FastAPI(title="Convertidor DOCX a PDF")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = Path(__file__).resolve().parent
TEMP_DIR = BASE_DIR / "temp"
TEMP_DIR.mkdir(exist_ok=True)

# WINDOWS
LIBREOFFICE_CMD = r"C:\Program Files\LibreOffice\program\soffice.exe"


def texto_pagina(pdf_reader: PdfReader, page_index: int) -> str:
    if page_index < 0 or page_index >= len(pdf_reader.pages):
        return ""

    try:
        text = pdf_reader.pages[page_index].extract_text()
        return text or ""
    except Exception:
        return ""


def pagina_esta_vacia_o_casi_vacia(pdf_reader: PdfReader, page_index: int) -> bool:
    texto = texto_pagina(pdf_reader, page_index)
    texto_limpio = " ".join(texto.split())

    # Sin texto o con texto mínimo
    if not texto_limpio:
        return True

    # Si tiene muy poquitos caracteres, también la tratamos como vacía
    if len(texto_limpio) < 20:
        return True

    return False


def limpiar_pdf_patrocinio(pdf_path: str) -> str:
    reader = PdfReader(pdf_path)
    total_paginas = len(reader.pages)

    # Si ya tiene 3 o menos, no hacemos nada
    if total_paginas <= 3:
        return pdf_path

    writer = PdfWriter()
    paginas_contenido = []

    # Conservamos solo páginas que tengan contenido real
    for i in range(total_paginas):
        if not pagina_esta_vacia_o_casi_vacia(reader, i):
            paginas_contenido.append(i)

    # Si no detectó nada, por seguridad devolver original
    if not paginas_contenido:
        return pdf_path

    # Si después de limpiar quedan más de 3, nos quedamos con las primeras 3
    paginas_finales = paginas_contenido[:3]

    for i in paginas_finales:
        writer.add_page(reader.pages[i])

    nuevo_pdf = pdf_path.replace(".pdf", "_limpio.pdf")
    with open(nuevo_pdf, "wb") as f:
        writer.write(f)

    return nuevo_pdf


@app.get("/")
def root():
    return {"ok": True, "mensaje": "Backend convertidor PDF activo"}


@app.post("/convertir-pdf")
async def convertir_pdf(
    file: UploadFile = File(...),
    tipo_documento: str = Form("general")
):
    nombre_original = (file.filename or "").lower()

    if not nombre_original.endswith(".docx"):
        raise HTTPException(status_code=400, detail="Solo se permiten archivos .docx")

    uid = str(uuid.uuid4())
    ruta_docx = TEMP_DIR / f"{uid}.docx"
    ruta_pdf = TEMP_DIR / f"{uid}.pdf"

    try:
        with open(ruta_docx, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        comando = [
            LIBREOFFICE_CMD,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", str(TEMP_DIR),
            str(ruta_docx)
        ]

        resultado = subprocess.run(
            comando,
            capture_output=True,
            text=True
        )

        if resultado.returncode != 0:
            raise HTTPException(
                status_code=500,
                detail=f"Error convirtiendo a PDF: {resultado.stderr or resultado.stdout}"
            )

        if not ruta_pdf.exists():
            raise HTTPException(status_code=500, detail="No se generó el PDF")

        pdf_final = ruta_pdf

        # SOLO para patrocinio
        if tipo_documento == "patrocinio":
            pdf_final = Path(limpiar_pdf_patrocinio(str(ruta_pdf)))

        return FileResponse(
            path=pdf_final,
            media_type="application/pdf",
            filename="documento.pdf"
        )

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error interno: {str(e)}")
    finally:
        try:
            file.file.close()
        except Exception:
            pass