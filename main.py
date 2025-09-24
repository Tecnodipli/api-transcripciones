import os
import re
import unicodedata
import zipfile
import io
import uuid
from collections import Counter
from datetime import datetime, timedelta
from io import BytesIO

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from docx import Document

app = FastAPI(title="Validador de Transcripciones")

# ==========================
# CORS: habilitar solo tus dominios
# ==========================
ALLOWED_ORIGINS = [
    "https://www.dipli.ai",
    "https://dipli.ai",
    "https://isagarcivill09.wixsite.com/turop",
    "https://isagarcivill09.wixsite.com/turop/tienda",
    "https://isagarcivill09-wixsite-com.filesusr.com",
    "https://www.dipli.ai/preparaci%C3%B3n",
    "https://www-dipli-ai.filesusr.com"
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,   # ✅ ahora sí aplica tu lista
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ==========================
# Almacenamiento temporal de descargas
# ==========================
DOWNLOADS = {}
EXP_MINUTES = 5  # tiempo de expiración del link

def cleanup_downloads():
    """Eliminar tokens expirados"""
    now = datetime.utcnow()
    expired = [t for t, (_, exp) in DOWNLOADS.items() if exp <= now]
    for t in expired:
        DOWNLOADS.pop(t, None)

# ==========================
# Validaciones
# ==========================
ETIQUETAS_VALIDAS = ["ENTREVISTADOR:", "ENTREVISTADORA:", "ENTREVISTADO:", "ENTREVISTADA:"]
REGEX_PERMITIDOS = r"[^A-Za-zÁÉÍÓÚáéíóúÑñ0-9\s\.,:\?¿]"

def char_human(ch: str) -> str:
    code = f"U+{ord(ch):04X}"
    name = unicodedata.name(ch, "UNKNOWN")
    visible = ch if not ch.isspace() else repr(ch)
    return f"{visible} ({code} {name})"

def validar_y_limpiar(doc: Document, filename: str):
    errores = []
    especiales_count = 0
    char_counter = Counter()

    for i, para in enumerate(doc.paragraphs):
        texto = para.text.strip()
        if not texto:
            continue

        texto_norm = texto.lower()

        # Ignorar timestamps tipo mm:ss
        if re.fullmatch(r"\d{1,2}:\d{2}", texto):
            continue

        # Validaciones etiquetas
        if "speaker" in texto_norm:
            errores.append((i+1, "Etiqueta inválida", f"Se encontró '{texto}'. Usa 'ENTREVISTADOR:' o 'ENTREVISTADO:'."))

        if "usuario" in texto_norm:
            errores.append((i+1, "Etiqueta inválida", f"Se encontró 'Usuario'. Usa 'ENTREVISTADO:' o 'ENTREVISTADOR:'."))

        if "xxx" in texto_norm:
            errores.append((i+1, "Etiqueta inválida", f"Se encontró '{texto}'. Reemplázalo por la etiqueta correcta."))

        match = re.match(r"^([A-ZÁÉÍÓÚÑ]+:)", texto)
        if match:
            etiqueta = match.group(1)
            if etiqueta not in ETIQUETAS_VALIDAS:
                errores.append((i+1, "Etiqueta inválida", f"Se encontró '{etiqueta}'. Usa solo {ETIQUETAS_VALIDAS}"))
            else:
                if texto == etiqueta:
                    errores.append((i+1, "Formato incorrecto", f"La etiqueta '{etiqueta}' está sola. Debe ir junto con el texto."))

                if etiqueta in ["ENTREVISTADOR:", "ENTREVISTADORA:"]:
                    encabezado_ok = any(run.text.strip().startswith(etiqueta) and run.bold for run in para.runs)
                    if not encabezado_ok:
                        errores.append((i+1, "Encabezado sin negrita", f"La etiqueta '{etiqueta}' debería estar en negrita."))

                    all_bold = all(run.bold or not run.text.strip() for run in para.runs)
                    if not all_bold:
                        errores.append((i+1, "Formato en negrita", f"El texto de '{etiqueta}' debería estar completamente en negrita."))

        # Fuente/tamaño
        for run in para.runs:
            fuente = run.font.name
            tamano = run.font.size.pt if run.font.size else None
            if fuente and fuente.lower() != "arial":
                errores.append((i+1, "Fuente incorrecta", f"Se detectó la fuente '{fuente}' en vez de Arial."))
                break
            if tamano and tamano != 12:
                errores.append((i+1, "Tamaño incorrecto", f"Se detectó {tamano}pt en vez de 12pt."))
                break

        # Limpieza
        for run in para.runs:
            if run.text:
                encontrados = re.findall(REGEX_PERMITIDOS, run.text)
                if encontrados:
                    especiales_count += len(encontrados)
                    char_counter.update(encontrados)
                    run.text = re.sub(REGEX_PERMITIDOS, "", run.text)

    # Guardar en memoria
    docx_bytes = BytesIO()
    doc.save(docx_bytes)
    docx_bytes.seek(0)

    # Crear reporte
    txt_bytes = BytesIO()
    resumen = {}

    txt_bytes.write(f"📋 REPORTE DE ERRORES\nArchivo: {filename}\nGenerado: {datetime.now()}\n\n".encode("utf-8"))
    for linea, tipo, desc in errores:
        txt_bytes.write(f"Línea {linea}: {tipo} → {desc}\n".encode("utf-8"))
        resumen[tipo] = resumen.get(tipo, 0) + 1

    if resumen:
        txt_bytes.write("\n--- RESUMEN DE ERRORES ---\n".encode("utf-8"))
        for tipo, count in resumen.items():
            txt_bytes.write(f"{tipo}: {count} ocurrencias\n".encode("utf-8"))

    txt_bytes.write("\n--- LIMPIEZA DE TEXTO ---\n".encode("utf-8"))
    txt_bytes.write(f"Total de caracteres especiales eliminados: {especiales_count}\n".encode("utf-8"))
    txt_bytes.write(f"Tipos únicos eliminados: {len(char_counter)}\n".encode("utf-8"))
    if char_counter:
        txt_bytes.write("\nDetalle por carácter:\n".encode("utf-8"))
        for ch, cnt in char_counter.most_common():
            txt_bytes.write(f"  {char_human(ch)} → {cnt}\n".encode("utf-8"))

    txt_bytes.seek(0)

    return docx_bytes, txt_bytes


@app.post("/procesar/")
async def procesar(file: UploadFile = File(...)):
    try:
        doc = Document(file.file)
    except Exception as e:
        return {"error": f"No se pudo abrir el archivo: {e}"}

    docx_bytes, txt_bytes = validar_y_limpiar(doc, file.filename)

    # Crear ZIP en memoria
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zipf:
        zipf.writestr(file.filename.replace(".docx", "_limpio.docx"), docx_bytes.read())
        zipf.writestr(file.filename.replace(".docx", "_errores.txt"), txt_bytes.read())
    zip_buffer.seek(0)

    # Crear token único
    token = str(uuid.uuid4())
    DOWNLOADS[token] = (zip_buffer.getvalue(), datetime.utcnow() + timedelta(minutes=EXP_MINUTES))

    return JSONResponse({"token": token, "expires_in": EXP_MINUTES})


@app.get("/download/{token}")
def download_token(token: str):
    cleanup_downloads()
    item = DOWNLOADS.get(token)
    if not item:
        raise HTTPException(status_code=404, detail="Link expirado o inválido")
    data, exp = item
    if exp <= datetime.utcnow():
        DOWNLOADS.pop(token, None)
        raise HTTPException(status_code=410, detail="Link expirado")
    headers = {"Content-Disposition": "attachment; filename=reportes_transcripciones.zip"}
    return StreamingResponse(io.BytesIO(data), media_type="application/zip", headers=headers)
