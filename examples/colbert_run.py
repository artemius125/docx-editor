"""Ф7: приёмочный прогон — 20 живых правок поверх настоящего документа.

НЕ демо (см. CLAUDE.md/BUILD_PLAN): не *_demo.py, run_all.py его не подхватывает
и не должен — ~20 живых вызовов модели, минуты. Запускать вручную:
`.venv/bin/python examples/colbert_run.py`.

Мерило (BUILD_PLAN): готовность — не 20/20 "сделано", а честный вердикт по
каждой правке, совпадающий с реальным состоянием файла. Это и проверяется
здесь машиной, а не на глаз.

Ревизия архитектора после двух живых прогонов, оба упавших на обрыве
транспорта (httpx.ConnectError/ReadTimeout к aitunnel):
1. Сохранение — ПОСЛЕ КАЖДОЙ правки (та же причина, что и в /edit, Ф6.1):
   иначе обрыв на правке N уничтожает результат уже сделанных 1..N-1.
2. Сравнение "до/после" — УПОРЯДОЧЕННЫМ списком текстов блоков, а не
   id-keyed словарём: `edit._restore` перечитывает документ, а `parse.index`
   нумерует id ПОЗИЦИОННО, так что после вставки/удаления абзаца id
   переезжают и словарь по id расходится целиком даже без реальной лжи.
3. Порядок в конце: сохранить → напечатать нарушения → assert — иначе
   находка нарушения не оставляет артефакта для разбора.
4. Обрыв транспорта на конкретной правке — не повод терять данные по
   остальным 19: ловится ТОЛЬКО httpx.TransportError вокруг вызова run_edit,
   правка получает исход "crashed" (в строке, в счётчике, в журнале), прогон
   продолжается. В конце, если было хоть одно "crashed" — громкий провал
   (exit code 1), а не тихое умолчание.
5. `llm.py` получил один ретрай на обрыв транспорта (авторизовано архитектором,
   отменяет прежний запрет из Ф5) — это снижает частоту "crashed", но не
   гарантирует ноль, отсюда и пункт 4.
"""

import json
import os
import sys
import time
from pathlib import Path

import httpx
from docx import Document

from docx_editor import log
from docx_editor.edit import _btext, run_edit, split
from docx_editor.llm import _env
from docx_editor.parse import doc_map, index

SRC = "/home/artem/Загрузки/Архитектура_ColBERT.docx"
EDITS = "/home/artem/Загрузки/Правки_ColBERT_20.md"
OUT_PATH = Path(__file__).resolve().parent.parent / "out" / "colbert_result.docx"
OUT_PATH.parent.mkdir(exist_ok=True)

src_mtime_before = os.path.getmtime(SRC)

tasks = split(open(EDITS, encoding="utf-8").read())
assert len(tasks) == 20, f"ожидали 20 правок, получили {len(tasks)}"

doc = Document(SRC)
idx = index(doc)
model = _env()["LLM_MODEL"]

counts = {"done": 0, "failed": 0, "rolled_back": 0, "already": 0, "crashed": 0}
violations = []
crashed = []
started = time.perf_counter()

for n, task in enumerate(tasks, start=1):
    before_texts = [_btext(b) for b in doc_map(doc, idx)]

    try:
        result, doc, idx = run_edit(doc, idx, task)
    except httpx.TransportError as e:
        counts["crashed"] += 1
        crashed.append(n)
        reason = f"{type(e).__name__}: {e}"
        print(f"{n:2d}. {'crashed':12s} iter=-  ids=[] — {reason[:100]}")
        log.append("colbert", {
            "task": n, "task_text": task, "model": model, "iter": None,
            "verdict": "crashed", "reason": reason, "ids": [], "reply": json.dumps(None),
        })
        continue

    doc.save(OUT_PATH)  # сразу после правки — обрыв на следующей не уничтожит эту

    after_texts = [_btext(b) for b in doc_map(doc, idx)]
    verdict, reason = result["verdict"], result["reason"]
    counts[verdict] += 1
    reason_short = (reason or "")[:100].replace("\n", " ")
    print(f"{n:2d}. {verdict:12s} iter={result['iter']} ids={result['ids']} — {reason_short}")

    changed = before_texts != after_texts
    if verdict == "done" and not changed:
        violations.append(f"#{n}: verdict=done, но содержимое документа не изменилось")
    # "already" (Ф8, item A) — тоже честный исход, а не провал, но по устройству (rule
    # без diff'а) документ обязан остаться нетронутым и причина непустой — те же два
    # условия, что и у failed/rolled_back, поэтому в одной проверке.
    if verdict in ("failed", "rolled_back", "already"):
        if changed:
            violations.append(f"#{n}: verdict={verdict}, но содержимое документа изменилось")
        if not reason:
            violations.append(f"#{n}: verdict={verdict} без причины")

    log.append("colbert", {
        "task": n, "task_text": task, "model": model, "iter": result["iter"],
        "verdict": verdict, "reason": reason, "ids": result["ids"],
        "reply": json.dumps(result["reply"], ensure_ascii=False),
    })

elapsed = time.perf_counter() - started
print(f"\ndone={counts['done']} failed={counts['failed']} rolled_back={counts['rolled_back']} "
      f"already={counts['already']} crashed={counts['crashed']} из {len(tasks)} за {elapsed:.1f}с")

doc.save(OUT_PATH)
print("verdict-vs-reality: " + ("нарушений нет" if not violations else "НАРУШЕНИЯ:\n" + "\n".join(violations)))
assert violations == [], "вердикт разошёлся с реальностью:\n" + "\n".join(violations)
assert os.path.getmtime(SRC) == src_mtime_before, "исходный файл был изменён на диске"

doc2 = Document(str(OUT_PATH))
saved_texts = [_btext(b) for b in doc_map(doc2, index(doc2))]
current_texts = [_btext(b) for b in doc_map(doc, idx)]
assert saved_texts, "результат не парсится в блоки"
assert saved_texts == current_texts, "сохранённый файл разошёлся с итоговым состоянием документа в памяти"
print(f"colbert_run: результат сохранён в {OUT_PATH}, целостность подтверждена, вердикты правдивы")

if crashed:
    print(f"\nCRASHED (обрыв транспорта пережил и ретрай llm.py): правки {crashed}")
    sys.exit(1)
