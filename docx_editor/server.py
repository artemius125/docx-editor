"""FastAPI-сервер: загрузка .docx, правка запросом (пока заглушка), скачивание."""

import json
import time
import uuid
from pathlib import Path

from docx import Document
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse

from docx_editor.parse import doc_map, index

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "out"
OUT.mkdir(exist_ok=True)

app = FastAPI()


def _session_path(session: str) -> Path:
    path = OUT / f"{session}.docx"
    if not path.exists():
        raise HTTPException(404, f"сессия {session} не найдена")
    return path


@app.get("/")
def root():
    return HTMLResponse((ROOT / "web" / "index.html").read_text())


@app.post("/upload")
async def upload(file: UploadFile = File(...)):
    """Заводит новую сессию из .docx, отдаёт карту документа для немедленного рендера."""
    session = uuid.uuid4().hex
    path = OUT / f"{session}.docx"
    path.write_bytes(await file.read())
    doc = Document(str(path))
    return {"session": session, "blocks": doc_map(doc, index(doc))}


@app.post("/edit")
async def edit(prompt: str = Form(...), session: str = Form(...)):
    """Правка поверх out/{session}.docx. Ф1: заглушка — planning, затем пустой result."""
    _session_path(session)  # проверка, что сессия существует

    def stream():
        started = time.perf_counter()
        yield json.dumps({"type": "planning"}) + "\n"
        yield json.dumps({
            "type": "result",
            "session": session,
            "done": [],
            "failed": [],
            "at": round((time.perf_counter() - started) * 1000),
        }) + "\n"

    return StreamingResponse(stream(), media_type="application/x-ndjson")


@app.get("/logs")
def logs(session: str | None = None, n: int = 50):
    """Журнал правок. Ф1: журнала ещё нет — пустой хвост."""
    return []


@app.get("/preview/{session}")
def preview(session: str):
    path = _session_path(session)
    doc = Document(str(path))
    return doc_map(doc, index(doc))


@app.get("/download/{session}")
def download(session: str):
    path = _session_path(session)
    return FileResponse(
        str(path),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=f"{session}.docx",
    )
