"""Приёмка edit.py: split() на настоящем файле правок, отказ/запасной поиск и
откат — с фейковыми ролями (без вызова модели), и ровно один живой вызов —
правка 12 на настоящем документе (см. CLAUDE.md: демо, закрываемое фикстурой,
закрываем фикстурой; живой вызов — только там, где без него нельзя).
"""

from docx import Document

from docx_editor.edit import run_edit, split
from docx_editor.parse import doc_map, index

REAL_DOC = "/home/artem/Загрузки/Архитектура_ColBERT.docx"
EDITS_FILE = "/home/artem/Загрузки/Правки_ColBERT_20.md"


def test_split_real_file():
    text = open(EDITS_FILE, encoding="utf-8").read()
    items = split(text)
    assert len(items) == 20, f"ожидали 20 правок, получили {len(items)}"
    assert items[11].startswith("Написано «снижает требования")
    print("edit_demo: split() даёт ровно 20 правок из реального файла")


def _fake_doc():
    doc = Document()
    doc.add_paragraph("Обычный абзац без ничего примечательного.")
    idx = index(doc)
    return doc, idx


def test_fallback_search_then_honest_refusal():
    # Навигатор промахнулся: ids содержит несуществующий блок, anchor не
    # находится — код обязан САМ попробовать запасной поиск по дословной
    # цитате в «ёлочках» из текста правки. Цитата тоже отсутствует в
    # документе, поэтому итог — честный отказ, а не тихое "сделано".
    doc, idx = _fake_doc()
    before = [b["text"] for b in doc_map(doc, idx)]

    def fake_navigator(outline_text, request):
        return {"kind": "local", "rule": None, "ids": ["p99"], "anchors": ["совсем другой текст"]}

    def fake_editor(*a, **kw):
        raise AssertionError("Редактор не должен вызываться — резолв не прошёл ни на одном этапе")

    request = "Правка по цитате «текст, которого точно нет в этом документе»."
    result, doc, idx = run_edit(doc, idx, request, navigator=fake_navigator, editor=fake_editor)

    assert result["verdict"] == "failed", result
    assert result["reason"], "у отказа должна быть причина"
    assert result["applied"] == [] and result["ids"] == []
    after = [b["text"] for b in doc_map(doc, idx)]
    assert after == before, "документ не должен измениться при отказе"
    print(f"edit_demo: запасной поиск отработал, честный отказ: {result['reason']!r}")


def test_rollback_on_checker_reject():
    # Редактор и Навигатор — валидные фейки (предлагают реальную, применимую
    # операцию), а Проверяющий говорит "не то" — цикл обязан откатить документ
    # к снимку и вернуть rolled_back, а не оставить недоделанную правку.
    doc, idx = _fake_doc()
    before = [b["text"] for b in doc_map(doc, idx)]

    def fake_navigator(outline_text, request):
        return {"kind": "local", "rule": None, "ids": ["p0"], "anchors": []}

    def fake_editor(fragment_text, request, feedback=None):
        return {"ops": [{"op": "replace_text", "id": "p0", "old": "Обычный", "new": "Особенный"}]}

    def fake_checker(request, diff):
        return {"ok": False, "reason": "тестовый отказ проверяющего"}

    result, doc, idx = run_edit(
        doc, idx, "замени 'Обычный' на 'Особенный'",
        navigator=fake_navigator, editor=fake_editor, checker=fake_checker,
    )

    assert result["verdict"] == "rolled_back", result
    assert result["reason"] == "тестовый отказ проверяющего"
    after = [b["text"] for b in doc_map(doc, idx)]
    assert after == before, "после отката текст обязан совпадать с исходным побайтово"
    print("edit_demo: откат по вердикту Проверяющего вернул документ как был")


def test_live_edit_12():
    # Единственный живой вызов модели во всём демо (см. CLAUDE.md — экономить
    # вызовы не обязательно, но лишний живой вызов делает прогон медленнее и
    # недетерминированнее без нужды). Правка 12 архитектор проверил живьём:
    # Навигатор надёжно находит p76.
    doc = Document(REAL_DOC)
    idx = index(doc)
    blocks = doc_map(doc, idx)
    p76_before = next(b["text"] for b in blocks if b["id"] == "p76")

    request = split(open(EDITS_FILE, encoding="utf-8").read())[11]
    result, doc, idx = run_edit(doc, idx, request)

    print(f"edit_demo: live verdict={result['verdict']!r} reason={result['reason']!r} "
          f"applied={result['applied']!r} ids={result['ids']!r}")

    assert result["verdict"] == "done", result
    p76_after = next(b["text"] for b in doc_map(doc, idx) if b["id"] == "p76")
    assert p76_after != p76_before, "p76 должен был измениться"
    print(f"edit_demo: p76 было: {p76_before!r}")
    print(f"edit_demo: p76 стало: {p76_after!r}")


if __name__ == "__main__":
    test_split_real_file()
    test_fallback_search_then_honest_refusal()
    test_rollback_on_checker_reject()
    test_live_edit_12()
