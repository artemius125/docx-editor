"""Измерительный прогон: 20 правок корпуса, каждая на СВОЁМ свежем документе
(изоляция — правка N не видит результата правки N-1). Пишет runs/<label>/...

Запуск: python3 run.py colbert|math   (BENCH_LABEL=... в окружении, default baseline)
"""

import json
import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import httpx
from docx import Document

REPO = "/home/artem/Документы/artemius125/docx-editor"
sys.path.insert(0, REPO)

from docx_editor.edit import _btext, run_edit, split  # noqa: E402
from docx_editor.parse import doc_map, index  # noqa: E402

CORPORA = {
    "colbert": {
        "doc": "/home/artem/Загрузки/Архитектура_ColBERT.docx",
        "edits": "/home/artem/Загрузки/Правки_ColBERT_20.md",
    },
    "math": {
        "doc": "/home/artem/Загрузки/Математика как основа.docx",
        "edits": "/home/artem/Загрузки/Правки_Математика_20.md",
    },
}

HERE = Path(__file__).resolve().parent


def _styles(blocks):
    return [b["style"] if b["kind"] == "p" else "[table]" for b in blocks]


def _run_one(n, task, src):
    doc = Document(src)
    idx = index(doc)
    blocks_before = doc_map(doc, idx)
    texts_before = [_btext(b) for b in blocks_before]
    styles_before = _styles(blocks_before)
    chars_before = sum(len(t) for t in texts_before)

    t0 = time.perf_counter()
    try:
        result, doc, idx = run_edit(doc, idx, task)
    except httpx.TransportError as e:
        # обрыв транспорта — не повод терять остальные правки (см. colbert_run.py);
        # doc/idx либо не тронуты, либо run_edit сам откатил их на месте
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
        return record, doc

    elapsed = time.perf_counter() - t0
    blocks_after = doc_map(doc, idx)
    texts_after = [_btext(b) for b in blocks_after]
    styles_after = _styles(blocks_after)
    record = {
        "n": n, "task": task, "verdict": result["verdict"], "reason": result["reason"],
        "applied": result["applied"], "ids": result["ids"], "iter": result["iter"],
        # что реально предложили Навигатор и Редактор: без этого у failed-строк
        # нет следа вообще и причину отказа не приписать к операции
        "reply": result["reply"],
        "blocks_before": texts_before, "blocks_after": texts_after,
        "styles_before": styles_before, "styles_after": styles_after,
        "chars_before": chars_before, "chars_after": sum(len(t) for t in texts_after),
        "seconds": elapsed,
    }
    return record, doc


def main():
    if len(sys.argv) != 2 or sys.argv[1] not in CORPORA:
        sys.exit(f"usage: run.py <corpus>, где corpus один из {list(CORPORA)}")
    corpus = sys.argv[1]
    label = os.environ.get("BENCH_LABEL", "baseline")
    src = CORPORA[corpus]["doc"]
    edits_path = CORPORA[corpus]["edits"]

    tasks = split(open(edits_path, encoding="utf-8").read())
    assert len(tasks) == 20, f"ожидали 20 правок, получили {len(tasks)}"

    out_dir = HERE / "runs" / label / corpus
    out_dir.mkdir(parents=True, exist_ok=True)
    jsonl_path = HERE / "runs" / label / f"{corpus}.jsonl"

    def _dump(results):
        # перезаписываем jsonl целиком после каждой правки — 20 строк, дёшево,
        # зато прогон, убитый на середине, оставляет валидные данные
        with open(jsonl_path, "w", encoding="utf-8") as f:
            for n in sorted(results):
                f.write(json.dumps(results[n], ensure_ascii=False) + "\n")

    def _print_line(r, suffix=""):
        dchars = r["chars_after"] - r["chars_before"]
        iter_s = r["iter"] if r["iter"] is not None else "-"
        reason = (r["reason"] or "")[:90].replace("\n", " ")
        print(f'{r["n"]:2d} {r["verdict"]:12s} iter={iter_s} dchars={dchars:+d} '
              f'ids={r["ids"]} — {reason}{suffix}', flush=True)

    workers = int(os.environ.get("BENCH_WORKERS", "2"))
    started = time.perf_counter()
    results = {}
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = {pool.submit(_run_one, n, task, src): n for n, task in enumerate(tasks, start=1)}
        for fut in futures:
            n = futures[fut]
            record, doc = fut.result()
            doc.save(out_dir / f"e{n:02d}.docx")
            results[n] = record
            _dump(results)
            _print_line(record)

    # ретрай-проход: crashed чинится серийно (пул тут только вредит — обрывы от
    # конкурентности), с паузой на остывание соединения перед каждой попыткой
    crashed_ns = sorted(n for n, r in results.items() if r["verdict"] == "crashed")
    recovered = []
    if crashed_ns:
        print(f"\nretry pass: {len(crashed_ns)} crashed — {crashed_ns}", flush=True)
        for n in crashed_ns:
            for attempt in (2, 3):
                print(f"retry {n:02d} (попытка {attempt}/3)", flush=True)
                time.sleep(5)
                record, doc = _run_one(n, tasks[n - 1], src)
                results[n] = record
                _dump(results)
                if record["verdict"] != "crashed":
                    doc.save(out_dir / f"e{n:02d}.docx")
                    recovered.append(n)
                    _print_line(record, " [recovered]")
                    break
                _print_line(record)
        still_crashed = sorted(set(crashed_ns) - set(recovered))
        print(f"\nretry pass done: recovered {len(recovered)}/{len(crashed_ns)}, "
              f"ещё crashed: {still_crashed}", flush=True)

    counts = {}
    for r in results.values():
        counts[r["verdict"]] = counts.get(r["verdict"], 0) + 1
    elapsed = time.perf_counter() - started
    summary = " ".join(f"{k}={v}" for k, v in sorted(counts.items()))
    print(f"\n{summary} из {len(tasks)} за {elapsed:.1f}с", flush=True)


if __name__ == "__main__":
    main()
