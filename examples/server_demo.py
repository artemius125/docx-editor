"""Проверка server.py целиком: upload -> preview -> edit-заглушка -> download.

Модель не вызывается, сеть не нужна: fastapi.testclient.TestClient держит
приложение в процессе.
"""

import io
import json
import tempfile

from docx import Document
from fastapi.testclient import TestClient

from docx_editor.server import app

DOC = "/home/artem/Загрузки/Архитектура_ColBERT.docx"


def main():
    client = TestClient(app)

    assert client.get("/").status_code == 200

    with open(DOC, "rb") as f:
        data = f.read()
    files = {"file": ("doc.docx", io.BytesIO(data), "application/octet-stream")}
    up = client.post("/upload", files=files)
    assert up.status_code == 200, up.text
    payload = up.json()
    session = payload["session"]
    assert len(payload["blocks"]) == 116, payload["blocks"][:1]

    prev = client.get(f"/preview/{session}")
    assert prev.status_code == 200
    assert len(prev.json()) == 116

    edit = client.post("/edit", data={"prompt": "тестовая правка", "session": session})
    assert edit.status_code == 200
    lines = [line for line in edit.text.splitlines() if line.strip()]
    events = [json.loads(line) for line in lines]
    assert events[0]["type"] == "planning"
    assert events[-1]["type"] == "result"
    assert events[-1]["done"] == []
    assert events[-1]["failed"] == []

    assert client.get("/logs").json() == []

    dl = client.get(f"/download/{session}")
    assert dl.status_code == 200
    with tempfile.NamedTemporaryFile(suffix=".docx") as tmp:
        tmp.write(dl.content)
        tmp.flush()
        doc = Document(tmp.name)
        assert len(doc.paragraphs) == 116, len(doc.paragraphs)

    print("server_demo: ok")


if __name__ == "__main__":
    main()
