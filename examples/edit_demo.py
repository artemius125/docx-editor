"""Приёмка edit.py: split() на настоящем файле правок, отказ/запасной поиск и
откат — с фейковыми ролями (без вызова модели), и ровно один живой вызов —
правка 12 на настоящем документе (см. CLAUDE.md: демо, закрываемое фикстурой,
закрываем фикстурой; живой вызов — только там, где без него нельзя).
"""

import os

from docx import Document

from docx_editor import edit as edit_mod
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


def test_ellipsis_truncation_retries_to_full_text():
    # Фикстура реального провала приёмки (Ф8, правка 18/p35): живой Редактор
    # ответил set_text, укоротив абзац и оборвав его буквальным «…» —
    # смысловая правка была на месте, поэтому Проверяющий сказал ok, а
    # машинная сверка "текст изменился" такое по устройству не ловит.
    # validate (Layer 1) обязан отбить обрезанный ответ ДО применения и
    # вернуть Редактору текст ошибки; получив обратную связь, Редактор
    # обязан прислать текст ЦЕЛИКОМ — и цикл обязан ПРОЙТИ ретраем, а не
    # просто честно отказать.
    doc = Document()
    original = "Первое предложение абзаца. Второе предложение абзаца, длиннее первого для правдоподобия."
    doc.add_paragraph(original)
    idx = index(doc)

    truncated = "Первое предложение абзаца. Второе предложение…"
    full = "Первое предложение абзаца полностью исправлено. Второе предложение абзаца, длиннее первого для правдоподобия."

    feedbacks = []

    def fake_navigator(outline_text, request):
        return {"kind": "local", "rule": None, "ids": ["p0"], "anchors": []}

    def fake_editor(fragment_text, request, feedback=None):
        feedbacks.append(feedback)
        if feedback is None:
            return {"ops": [{"op": "set_text", "id": "p0", "text": truncated}]}
        return {"ops": [{"op": "set_text", "id": "p0", "text": full}]}

    def fake_checker(request, diff):
        return {"ok": True, "reason": "ok"}

    result, doc, idx = run_edit(
        doc, idx, "исправь первое предложение",
        navigator=fake_navigator, editor=fake_editor, checker=fake_checker,
    )

    assert result["verdict"] == "done", result
    assert result["iter"] == 2, "guard обязан отбить первый ответ и провести настоящий ретрай, а не убить правку"
    after_text = next(b["text"] for b in doc_map(doc, idx) if b["id"] == "p0")
    assert after_text == full, after_text
    assert not after_text.rstrip().endswith(("…", "...")), "результат не должен обрываться многоточием"
    assert len(after_text) >= len(original), "текст не должен схлопнуться короче исходного"
    assert len(feedbacks) == 2 and feedbacks[0] is None and feedbacks[1], \
        "второй вызов редактора обязан нести текст ошибки валидатора"
    print(f"edit_demo: обрезанный set_text отбит validate, ретрай дал полный текст, iter={result['iter']}")


def test_check_diff_carries_lengths():
    # Layer 2 (Ф8): _check обязан показывать Проверяющему длины "было"/"стало",
    # иначе числовой сигнал "текст вдвое короче" молча сгниёт при рефакторинге.
    from docx_editor.edit import _check

    seen = {}

    def fake_llm_chat(messages):
        seen["user"] = messages[1]["content"]
        return {"ok": True, "reason": "ok"}

    import docx_editor.llm as llm_module
    real_chat = llm_module.chat
    llm_module.chat = fake_llm_chat
    try:
        diff = [{"id": "p35", "before": "а" * 456, "after": "б" * 332}]
        _check("правка", diff)
    finally:
        llm_module.chat = real_chat

    assert "456" in seen["user"] and "332" in seen["user"], seen["user"]
    print("edit_demo: diff, отдаваемый Проверяющему, несёт длины было/стало")


def test_rule_no_op_yields_already():
    # Ф8 item A: правка типа rule, ничего не изменившая (документ уже
    # нормализован — как правка 3 "слипшиеся предложения" после того, как
    # правка 2 того же запроса уже прогнала typography по всему документу),
    # обязана дать verdict="already", а не "failed": требуемое состояние УЖЕ
    # в документе, и "не удалось" вводит в заблуждение не меньше ложного
    # "сделано" (инвариант 5).
    doc = Document()
    doc.add_paragraph("Уже нормальный абзац: без лишних пробелов; и без слипшихся предложений.")
    idx = index(doc)
    before = [b["text"] for b in doc_map(doc, idx)]

    def fake_navigator(outline_text, request):
        return {"kind": "global", "rule": "typography", "ids": [], "anchors": []}

    result, doc, idx = run_edit(doc, idx, "раздели слипшиеся предложения", navigator=fake_navigator)

    assert result["verdict"] == "already", result
    assert result["reason"], "у already обязана быть непустая причина"
    after = [b["text"] for b in doc_map(doc, idx)]
    assert after == before, "already не должен менять документ"
    print(f"edit_demo: rule без изменений даёт already: {result['reason']!r}")


def test_non_rule_empty_diff_stays_failed():
    # Тот же симптом (ops применились, diff пуст), но НЕ через rule — здесь
    # "already" не полагается (см. правило item A): нет оснований заключить,
    # что желаемое состояние достигнуто, только что Редактор ничего не
    # поменял по факту. set_style на тот же стиль, что уже есть, — валидная
    # операция (validate не запрещает старый==новый для стиля), которая
    # применяется, но не меняет ТЕКСТ, который сравнивает _diff.
    doc = Document()
    doc.add_paragraph("Абзац без изменений текста.")
    idx = index(doc)
    before = [b["text"] for b in doc_map(doc, idx)]

    def fake_navigator(outline_text, request):
        return {"kind": "local", "rule": None, "ids": ["p0"], "anchors": []}

    def fake_editor(fragment_text, request, feedback=None):
        return {"ops": [{"op": "set_style", "id": "p0", "style": "Normal"}]}

    def fake_checker(request, diff):
        raise AssertionError("Проверяющий не должен вызываться — diff пуст раньше")

    result, doc, idx = run_edit(
        doc, idx, "измени абзац",
        navigator=fake_navigator, editor=fake_editor, checker=fake_checker,
    )

    assert result["verdict"] == "failed", result
    assert result["reason"] == "операции применились, но текст не изменился", result
    after = [b["text"] for b in doc_map(doc, idx)]
    assert after == before
    print("edit_demo: non-rule путь с пустым diff остаётся failed (already — только для rule)")


def test_quoted_phrase_always_merges_with_navigator_hit():
    # Ф8 item C: правки 4/8/14 отказывались с «во фрагменте только один
    # случай», потому что дословная цитата «ёлочками» из текста правки
    # искалась ТОЛЬКО как запасной путь — когда Навигатор не назвал ни
    # одного id/anchor'а. Навигатор чаще называет ОДИН случай (тот, что
    # увидел первым), и запасной путь никогда не срабатывал, хотя цитата
    # реально встречается в нескольких местах документа (правка 8 — три
    # разных написания года, каждое в «ёлочках», в трёх разных блоках).
    # Теперь цитата всегда доливает совпадения, а не только когда всё
    # остальное пусто.
    doc = Document()
    doc.add_paragraph("Первое: термин встречается тут.")
    doc.add_paragraph("Второе предложение без искомого слова.")
    doc.add_paragraph("Третье: термин встречается снова тут.")
    doc.add_paragraph("Четвёртое: термин встречается и здесь тоже.")
    idx = index(doc)

    seen_fragments = []

    def fake_navigator(outline_text, request):
        # Навигатор нашёл только ОДИН случай (p0) — как измерено на живом Навигаторе.
        return {"kind": "local", "rule": None, "ids": ["p0"], "anchors": []}

    def fake_editor(fragment_text, request, feedback=None):
        seen_fragments.append(fragment_text)
        return {"ops": []}

    request = "Замени везде «термин встречается» на другую формулировку."
    run_edit(doc, idx, request, navigator=fake_navigator, editor=fake_editor)

    assert seen_fragments, "Редактор обязан быть вызван"
    fragment_text = seen_fragments[0]
    assert "p0" in fragment_text, fragment_text
    assert "p2" in fragment_text, fragment_text
    assert "p3" in fragment_text, fragment_text
    print("edit_demo: дословная цитата в «ёлочках» слита с id Навигатора, все 3 случая дошли до фрагмента Редактора")


def test_checker_exception_restores_document_and_cleans_snapshot():
    # Ф8 item D: если Проверяющий падает (оборвался транспорт), патч уже
    # применён к живому doc, снимок ещё жив. Исключение обязано долететь
    # наружу ГРОМКО (инвариант 6), но doc обязан вернуться к состоянию ДО
    # правки — ТОЙ ЖЕ ссылкой, потому что вызывающий не получает новую пару
    # (result, doc, idx) через return (управление уходит через raise), а
    # снимок обязан быть удалён, а не течь во временные файлы.
    doc = Document()
    doc.add_paragraph("Оригинальный текст абзаца до правки.")
    idx = index(doc)
    before = [b["text"] for b in doc_map(doc, idx)]

    def fake_navigator(outline_text, request):
        return {"kind": "local", "rule": None, "ids": ["p0"], "anchors": []}

    def fake_editor(fragment_text, request, feedback=None):
        return {"ops": [{"op": "replace_text", "id": "p0", "old": "Оригинальный", "new": "Изменённый"}]}

    def fake_checker(request, diff):
        raise RuntimeError("обрыв транспорта на проверке (тест)")

    captured = {}
    real_snapshot = edit_mod._snapshot

    def spy_snapshot(d):
        path = real_snapshot(d)
        captured["path"] = path
        return path

    edit_mod._snapshot = spy_snapshot
    try:
        raised = False
        try:
            run_edit(
                doc, idx, "замени 'Оригинальный' на 'Изменённый'",
                navigator=fake_navigator, editor=fake_editor, checker=fake_checker,
            )
        except RuntimeError:
            raised = True
    finally:
        edit_mod._snapshot = real_snapshot

    assert raised, "исключение Проверяющего обязано долететь наружу, а не проглотиться"
    assert "path" in captured, "снимок обязан был создаваться (ops успели примениться)"
    assert not os.path.exists(captured["path"]), "снимок обязан быть удалён даже при исключении, а не течь"

    after = [b["text"] for b in doc_map(doc, idx)]
    assert after == before, "doc (та же ссылка) обязан вернуться к состоянию ДО правки"
    print("edit_demo: исключение Проверяющего долетает наружу, doc восстановлен на месте, снимок не течёт")


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
    test_ellipsis_truncation_retries_to_full_text()
    test_check_diff_carries_lengths()
    test_rule_no_op_yields_already()
    test_non_rule_empty_diff_stays_failed()
    test_quoted_phrase_always_merges_with_navigator_hit()
    test_checker_exception_restores_document_and_cleans_snapshot()
    test_live_edit_12()
