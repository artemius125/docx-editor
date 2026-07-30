"""Прогон за рулём: роли модели исполняет человек (или агент), а не LLM.

Ответы лежат в runs/drive/<corpus>/e<NN>.json — список ответов по порядку
вызовов. Скрипт прокручивает правку с начала на свежей копии документа,
подставляя записанные ответы; на первом вызове без ответа печатает то, что
УВИДЕЛА БЫ модель (system + user), и останавливается. Дописал ответ в конец
файла — запустил снова. Когда ответы кончились, а run_edit вернул вердикт,
печатается вердикт и сохраняется .docx.

Запуск: python3 drive.py <corpus> <N> [--sys]
"""

import json
import sys
from pathlib import Path

from docx import Document

REPO = "/home/artem/Документы/artemius125/docx-editor"
sys.path.insert(0, REPO)

from docx_editor import llm  # noqa: E402
from docx_editor.edit import run_edit, split  # noqa: E402
from docx_editor.parse import index  # noqa: E402
from run import CORPORA  # noqa: E402

HERE = Path(__file__).resolve().parent

# Структурный корпус (таблица, заголовки, список) лежит в репозитории и в
# CORPORA стенда не входит — у run.py 20 правок, здесь их 10.
CORPORA = dict(CORPORA, reglament={
    "doc": str(HERE / "fixtures" / "Регламент.docx"),
    "edits": str(HERE / "fixtures" / "Правки_Регламент_10.md"),
})


class Stop(Exception):
    """Ответа на этот вызов ещё нет — печатаем запрос и выходим."""


def main():
    if len(sys.argv) < 3 or sys.argv[1] not in CORPORA:
        sys.exit(f"usage: drive.py <corpus> <N> [--sys], где corpus один из {list(CORPORA)}")
    corpus, n = sys.argv[1], int(sys.argv[2])
    show_sys = "--sys" in sys.argv

    src = CORPORA[corpus]["doc"]
    tasks = split(open(CORPORA[corpus]["edits"], encoding="utf-8").read())
    task = tasks[n - 1]

    out_dir = HERE / "runs" / "drive" / corpus
    out_dir.mkdir(parents=True, exist_ok=True)
    ans_path = out_dir / f"e{n:02d}.json"
    answers = json.loads(ans_path.read_text(encoding="utf-8")) if ans_path.exists() else []

    state = {"i": 0}

    def fake_chat(messages, **kw):
        i = state["i"]
        state["i"] += 1
        if i < len(answers):
            return answers[i]
        print(f"=== ВЫЗОВ #{i} — ответа нет, дописать в {ans_path}")
        if show_sys:
            print(f"--- system ---\n{messages[0]['content']}")
        print(f"--- user ---\n{messages[1]['content']}")
        raise Stop

    llm.chat = fake_chat

    doc = Document(src)
    idx = index(doc)
    print(f"=== {corpus}#{n}: {task}\n")
    try:
        result, doc, idx = run_edit(doc, idx, task)
    except Stop:
        return

    doc.save(out_dir / f"e{n:02d}.docx")
    print(f'=== ВЕРДИКТ {result["verdict"]} за {state["i"]} вызовов\nпричина: {result["reason"]}')
    for line in result["applied"]:
        print(f"  применено: {line}")


if __name__ == "__main__":
    main()
