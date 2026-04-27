from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pathlib import Path
import shutil
import subprocess
import uuid

app = FastAPI(title="Convertidor DOCX a PDF")

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://127.0.0.1:5500",
        "http://localhost:5500",
        "http://127.0.0.1:5501",
        "http://localhost:5501",
        "http://127.0.0.1:4200",
        "http://localhost:4200",
        "https://capacitacindocenteitsqmet.netlify.app"
    ],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = Path(__file__).resolve().parent
TEMP_DIR = BASE_DIR / "temp"
TEMP_DIR.mkdir(exist_ok=True)

LIBREOFFICE_CMD = "soffice"


@app.get("/")
def root():
    return {"ok": True, "mensaje": "Backend convertidor PDF activo"}


@app.post("/convertir-pdf")
async def convertir_pdf(file: UploadFile = File(...)):
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

        return FileResponse(
            path=str(ruta_pdf),
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