"""Приёмка edit.py: split() на настоящем файле правок, отказ/запасной поиск и
откат — с фейковыми ролями (без вызова модели), и ровно один живой вызов —
правка 12 на настоящем документе (см. CLAUDE.md: демо, закрываемое фикстурой,
закрываем фикстурой; живой вызов — только там, где без него нельзя).
"""

from docx import Document

from docx_editor.edit import _btext, run_edit, split
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


def test_resolve_normalizes_integer_id():
    # Находка Ф5 (реопен, зеркало уже описанной в BUILD_PLAN "id внутрь
    # цитаты"): живой Навигатор вернул ids: [112] — ЧИСЛОМ, а не строкой
    # "p112". Старый _resolve сравнивал напрямую со множеством строк и тихо
    # терял такой id, резолв выживал только на якоре и уезжал не туда.
    # _resolve обязан нормализовать форму сам, а не полагаться на промпт.
    doc, idx = _fake_doc()
    reached = []

    def fake_navigator(outline_text, request):
        return {"kind": "local", "rule": None, "ids": [0], "anchors": []}  # int, не "p0"

    def fake_editor(fragment_text, request, feedback=None):
        reached.append(fragment_text)
        return {"ops": [{"op": "replace_text", "id": "p0", "old": "Обычный", "new": "Особенный"}]}

    def fake_checker(request, diff):
        return {"ok": True, "reason": "ok"}

    result, doc, idx = run_edit(
        doc, idx, "замени 'Обычный' на 'Особенный'",
        navigator=fake_navigator, editor=fake_editor, checker=fake_checker,
    )

    assert reached, "целочисленный id (0) от Навигатора обязан резолвиться в p0 и дойти до Редактора"
    assert result["verdict"] == "done", result
    print("edit_demo: целочисленный id от Навигатора (0) нормализован в p0 и резолвится")


def test_live_edit_12():
    # Единственный живой вызов модели во всём демо (см. CLAUDE.md — экономить
    # вызовы не обязательно, но лишний живой вызов делает прогон медленнее и
    # недетерминированнее без нужды).
    #
    # Мерило (BUILD_PLAN «Мерило») — НЕ "сделано", а честный вердикт,
    # совпадающий с реальным состоянием файла. 26B иногда честно отказывается
    # там, где обычно проходит (см. Ф5 «Известные риски») — это не провал
    # демо, а ровно то поведение, которое демо обязано пропустить, не соврав.
    # Поэтому здесь не проверяется verdict == "done": проверяется, что каждый
    # из трёх legal-исходов соответствует правде о документе.
    doc = Document(REAL_DOC)
    idx = index(doc)
    blocks = doc_map(doc, idx)
    p76_before = next(b["text"] for b in blocks if b["id"] == "p76")
    # Сравнение по СОДЕРЖИМОМУ блоков (текст абзацев + ячеек таблиц), а не по
    # сырым байтам файла: python-docx пишет .docx через zipfile.writestr,
    # который штампует каждую запись архива текущим временем — один и тот же
    # нетронутый документ даёт разные байты на каждом save. _btext — тот же
    # хелпер, что уже использует run_edit._diff для сравнения "до/после".
    before_texts = [_btext(b) for b in blocks]

    request = split(open(EDITS_FILE, encoding="utf-8").read())[11]
    result, doc, idx = run_edit(doc, idx, request)

    print(f"edit_demo: live verdict={result['verdict']!r} reason={result['reason']!r}")

    assert result["verdict"] in ("done", "failed", "rolled_back"), result

    if result["verdict"] == "done":
        p76_after = next(b["text"] for b in doc_map(doc, idx) if b["id"] == "p76")
        assert p76_after != p76_before, "verdict=done обязан означать реальное изменение p76"
        print(f"edit_demo: p76 было: {p76_before!r}")
        print(f"edit_demo: p76 стало: {p76_after!r}")
    else:
        assert result["reason"], "у отказа/отката обязана быть непустая причина"
        after_texts = [_btext(b) for b in doc_map(doc, idx)]
        assert after_texts == before_texts, "не-done обязан оставить содержимое документа нетронутым"
        print(f"edit_demo: содержимое документа не тронуто, ничего не применено вполовину (verdict={result['verdict']!r})")


if __name__ == "__main__":
    test_split_real_file()
    test_fallback_search_then_honest_refusal()
    test_rollback_on_checker_reject()
    test_resolve_normalizes_integer_id()
    test_live_edit_12()
