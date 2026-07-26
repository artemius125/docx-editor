"""Журнал правок: одна JSON-строка на правку в out/log.jsonl. Только запись
и чтение хвоста — без ротации и без класса (см. BUILD_PLAN «Чего нет»)."""

import json
from pathlib import Path

_PATH = Path(__file__).resolve().parent.parent / "out" / "log.jsonl"


def append(session: str, record: dict) -> None:
    """Дописывает одну запись правки. record несёт task/task_text/model/iter/
    verdict/reason/ids/reply (см. docx_editor.edit.run_edit — оттуда и
    значения iter/reply) — session хранится отдельным полем, по нему
    фильтрует tail()."""
    _PATH.parent.mkdir(exist_ok=True)
    line = json.dumps({"session": session, **record}, ensure_ascii=False)
    with _PATH.open("a", encoding="utf-8") as f:
        f.write(line + "\n")


def tail(session: str | None = None, n: int = 50) -> list[dict]:
    """Последние n записей журнала; при заданном session — только его записи."""
    if not _PATH.exists():
        return []
    records = [json.loads(line) for line in _PATH.read_text(encoding="utf-8").splitlines() if line.strip()]
    if session is not None:
        records = [r for r in records if r["session"] == session]
    return records[-n:]
