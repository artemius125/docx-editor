"""Приёмка server.py: upload -> /edit по настоящему циклу (Навигатор/Редактор/
Проверяющий подменены фикстурой llm.chat — без сети и без денег) -> download.

edit.py зовёт модель через docx_editor.llm.chat как атрибут модуля, поэтому
подмена llm.chat отражается и в вызовах edit.run_edit изнутри /edit — без
передачи фейковых navigator/editor/checker через HTTP-контракт (там их нет).
"""

import io
import json
import tempfile

from docx import Document
from fastapi.testclient import TestClient

import docx_editor.llm as llm
from docx_editor.server import app

DOC = "/home/artem/Загрузки/Архитектура_ColBERT.docx"

# Правка 1 — существующий в документе текст (p7), Навигатор находит его сразу,
# Редактор предлагает реальную замену, Проверяющий подтверждает: verdict=done.
OLD_TEXT = "высочайшая скорость на этапе инференса"
NEW_TEXT = "высокую скорость на этапе инференса"

# Правка 2 — Навигатор промахивается (несуществующий id, ненаходимый anchor), а
# в тексте самой правки нет ни одной «ёлочной» цитаты, которая реально есть в
# документе — запасной поиск по цитатам тоже пуст: честный verdict=failed.
PROMPT = (
    f"1. Замени «{OLD_TEXT}» на «{NEW_TEXT}».\n"
    "2. Впиши требование «совершенно вымышленная и точно отсутствующая в документе фраза»."
)

# Правка 2' — намеренный сбой Навигатора (не осознанный отказ, а падение) на
# ВТОРОЙ правке второй, отдельной сессии: проверяет, что (а) исключение не
# глотается где-то в потоке /edit, долетает до вызывающего, и (б) файл на
# диске уже несёт правку 1, применённую и сохранённую ДО падения — сохранение
# внутри цикла не должно "отставать" от событий, которые уже ушли в поток.
CRASH_MARKER = "ломает конвейер"
PROMPT_CRASH = (
    f"1. Замени «{OLD_TEXT}» на «{NEW_TEXT}».\n"
    f"2. Правка, которая {CRASH_MARKER} (тест громких ошибок)."
)

_REQUIRED_OP_FIELDS = {
    "type", "status", "text", "task", "task_text", "model", "at", "dt", "blocks", "verdict",
}


def _fake_chat(messages):
    system, user = messages[0]["content"], messages[1]["content"]
    if "навигатор" in system:
        if OLD_TEXT in user:
            return {"kind": "local", "rule": None, "ids": ["p7"], "anchors": []}
        if CRASH_MARKER in user:
            raise RuntimeError("сбой навигатора (тест громких ошибок)")
        return {"kind": "local", "rule": None, "ids": ["p999"], "anchors": ["текст, которого точно нет"]}
    if "редактор" in system:
        return {"ops": [{"op": "replace_text", "id": "p7", "old": OLD_TEXT, "new": NEW_TEXT}]}
    if "проверяющий" in system:
        return {"ok": True, "reason": "текст заменён корректно"}
    raise AssertionError(f"неожиданный system prompt: {system[:40]!r}")


def main():
    llm.chat = _fake_chat
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

    edit = client.post("/edit", data={"prompt": PROMPT, "session": session})
    assert edit.status_code == 200
    lines = [line for line in edit.text.splitlines() if line.strip()]
    events = [json.loads(line) for line in lines]

    assert events[0]["type"] == "planning"
    assert events[-1]["type"] == "result"
    op_events = [e for e in events[1:-1] if e["type"] == "op"]
    assert len(op_events) == 2, op_events

    done_evt, failed_evt = op_events
    for evt in op_events:
        assert _REQUIRED_OP_FIELDS <= evt.keys(), evt.keys()

    assert done_evt["verdict"] == "done", done_evt
    assert done_evt["status"] == "done"
    assert done_evt["task"] == 1
    assert done_evt["task_text"].startswith("Замени «")
    p7 = next(b for b in done_evt["blocks"] if b["id"] == "p7")
    assert NEW_TEXT in p7["text"], p7["text"]
    assert OLD_TEXT not in p7["text"], p7["text"]

    assert failed_evt["verdict"] == "failed", failed_evt
    assert failed_evt["status"] == "failed"
    assert failed_evt["task"] == 2
    assert failed_evt["text"], "у отказа должен быть непустой текст причины"

    result = events[-1]
    assert result["done"] == [done_evt["text"]]
    assert result["failed"] == [failed_evt["text"]]

    assert client.get("/logs").json() == []

    dl = client.get(f"/download/{session}")
    assert dl.status_code == 200
    with tempfile.NamedTemporaryFile(suffix=".docx") as tmp:
        tmp.write(dl.content)
        tmp.flush()
        doc = Document(tmp.name)
        assert len(doc.paragraphs) == 116, len(doc.paragraphs)
        assert NEW_TEXT in doc.paragraphs[7].text, "правка не дошла до файла на диске"

    print("server_demo: ok")


def test_crash_does_not_lose_saved_progress():
    """Правка 2 роняет Навигатора. Инвариант 6 — падение обязано долететь
    наружу, а не превратиться в тихий result. Инвариант 5 — правка 1 к этому
    моменту уже сохранена на диск (save внутри цикла, а не после него), так
    что /download отдаёт файл С этой правкой, а не старую версию."""
    llm.chat = _fake_chat
    client = TestClient(app)
    with open(DOC, "rb") as f:
        data = f.read()
    files = {"file": ("doc.docx", io.BytesIO(data), "application/octet-stream")}
    session = client.post("/upload", files=files).json()["session"]

    try:
        client.post("/edit", data={"prompt": PROMPT_CRASH, "session": session})
        raised = False
    except RuntimeError:
        raised = True
    assert raised, "падение Навигатора обязано долететь до вызывающего, а не потеряться в потоке"

    dl = client.get(f"/download/{session}")
    assert dl.status_code == 200
    with tempfile.NamedTemporaryFile(suffix=".docx") as tmp:
        tmp.write(dl.content)
        tmp.flush()
        doc = Document(tmp.name)
        assert NEW_TEXT in doc.paragraphs[7].text, "правка 1 обязана была сохраниться на диск до падения правки 2"

    print("server_demo: падение внутри /edit долетает наружу, правка 1 уже на диске")


if __name__ == "__main__":
    main()
    test_crash_does_not_lose_saved_progress()
