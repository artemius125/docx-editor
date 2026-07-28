"""Кумулятивный прогон: все 20 правок корпуса — на ОДНОМ документе, по очереди
(single-threaded, документ общий и мутируется), каждая правка видит результат
предыдущих. В отличие от run.py, где каждая правка — на своём свежем документе
(изоляция), здесь проверяется реалистичный сценарий: способность системы
пронести весь список правок по цепочке.

Запуск: python3 run_seq.py colbert|math   (BENCH_LABEL=... в окружении, default baseline)
"""

import json
import os
import sys
import time

import httpx
from docx import Document

from run import CORPORA, HERE, _styles
from docx_editor.edit import _btext, run_edit, split
from docx_editor.parse import doc_map, index


def main():
    if len(sys.argv) != 2 or sys.argv[1] not in CORPORA:
        sys.exit(f"usage: run_seq.py <corpus>, где corpus один из {list(CORPORA)}")
    corpus = sys.argv[1]
    label = os.environ.get("BENCH_LABEL", "baseline")
    src = CORPORA[corpus]["doc"]
    edits_path = CORPORA[corpus]["edits"]

    tasks = split(open(edits_path, encoding="utf-8").read())
    assert len(tasks) == 20, f"ожидали 20 правок, получили {len(tasks)}"

    out_dir = HERE / "runs" / label
    out_dir.mkdir(parents=True, exist_ok=True)
    docx_path = out_dir / f"{corpus}_seq.docx"
    jsonl_path = out_dir / f"{corpus}_seq.jsonl"

    doc = Document(src)
    idx = index(doc)

    started = time.perf_counter()
    counts = {}
    with open(jsonl_path, "w", encoding="utf-8") as f:
        for n, task in enumerate(tasks, start=1):
            blocks_before = doc_map(doc, idx)
            texts_before = [_btext(b) for b in blocks_before]
            styles_before = _styles(blocks_before)
            chars_before = sum(len(t) for t in texts_before)

            t0 = time.perf_counter()
            for attempt in range(3):
                try:
                    result, doc, idx = run_edit(doc, idx, task)
                except httpx.TransportError as e:
                    # обрыв транспорта — ретрай ТОЙ ЖЕ правки на месте (до 3 попыток),
                    # не второй проход в конце: в кумулятивной цепочке повтор правки N
                    # после N+1..20 поменял бы смысл. doc/idx либо не тронуты, либо
                    # run_edit сам откатил их на месте — в любом случае это ТЕ ЖЕ doc/idx.
                    if attempt < 2:
                        time.sleep(5)
                        continue
                    elapsed = time.perf_counter() - t0
                    blocks_after = doc_map(doc, idx)
                    texts_after = [_btext(b) for b in blocks_after]
                    styles_after = _styles(blocks_after)
                    record = {
                        "n": n, "task": task, "verdict": "crashed", "reason": f"{type(e).__name__}: {e}",
                        "applied": [], "ids": [], "iter": None,
                        "blocks_before": texts_before, "blocks_after": texts_after,
                        "styles_before": styles_before, "styles_after": styles_after,
                        "chars_before": chars_before, "chars_after": sum(len(t) for t in texts_after),
                        "seconds": elapsed,
                    }
                else:
                    elapsed = time.perf_counter() - t0
                    blocks_after = doc_map(doc, idx)
                    texts_after = [_btext(b) for b in blocks_after]
                    styles_after = _styles(blocks_after)
                    record = {
                        "n": n, "task": task, "verdict": result["verdict"], "reason": result["reason"],
                        "applied": result["applied"], "ids": result["ids"], "iter": result["iter"],
                        "blocks_before": texts_before, "blocks_after": texts_after,
                        "styles_before": styles_before, "styles_after": styles_after,
                        "chars_before": chars_before, "chars_after": sum(len(t) for t in texts_after),
                        "seconds": elapsed,
                    }
                break

            doc.save(docx_path)  # после КАЖДОЙ правки — крах на следующей не уничтожит уже сделанное
            f.write(json.dumps(record, ensure_ascii=False) + "\n")
            f.flush()

            counts[record["verdict"]] = counts.get(record["verdict"], 0) + 1
            dchars = record["chars_after"] - record["chars_before"]
            iter_s = record["iter"] if record["iter"] is not None else "-"
            reason = (record["reason"] or "")[:90].replace("\n", " ")
            print(f'{record["n"]:2d} {record["verdict"]:12s} iter={iter_s} dchars={dchars:+d} '
                  f'ids={record["ids"]} — {reason}')

    elapsed = time.perf_counter() - started
    summary = " ".join(f"{k}={v}" for k, v in sorted(counts.items()))
    print(f"\n{summary} из {len(tasks)} за {elapsed:.1f}с")


if __name__ == "__main__":
    main()
