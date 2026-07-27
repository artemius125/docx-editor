"""Приёмка edit.py: split() на настоящем файле правок, отказ/запасной поиск,
откат и правка 12 на настоящем документе — везде фейковые роли (без вызова
модели, см. CLAUDE.md: демо, закрываемое фикстурой, закрываем фикстурой).
Живая приёмка того же корпуса — examples/colbert_run.py, не *_demo.py.
"""

import os

from docx import Document

from docx_editor import edit as edit_mod
from docx_editor import patch
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
    print("edit_demo: non-rule путь с пустым diff остаётся failed (already — только через rule/fallthrough, не отсюда)")


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


def test_set_style_diff_reaches_checker():
    # Change A: _diff раньше сравнивал только текст — set_style была ей
    # невидима, Проверяющий получал пустой diff, и run_edit откатывал правку
    # с "текст не изменился", хотя стиль реально сменился (это и произошло на
    # живой правке 4 корпуса Математика). Теперь diff обязан нести стиль, и
    # такая правка обязана доходить до Проверяющего и закрываться done.
    doc = Document()
    doc.add_paragraph("Заголовок раздела")
    idx = index(doc)

    def fake_navigator(outline_text, request):
        return {"kind": "local", "rule": None, "ids": ["p0"], "anchors": []}

    def fake_editor(fragment_text, request, feedback=None):
        return {"ops": [{"op": "set_style", "id": "p0", "style": "Heading 1"}]}

    seen_diff = []

    def fake_checker(request, diff):
        seen_diff.append(diff)
        return {"ok": True, "reason": "стиль сменился на заголовок"}

    result, doc, idx = run_edit(
        doc, idx, "сделай заголовком",
        navigator=fake_navigator, editor=fake_editor, checker=fake_checker,
    )

    assert seen_diff and seen_diff[0], "diff обязан быть непустым — Проверяющий должен увидеть изменение"
    assert "Heading 1" in seen_diff[0][0].get("note", ""), seen_diff[0]
    assert result["verdict"] == "done", result
    print(f"edit_demo: set_style дошёл до Проверяющего с note={seen_diff[0][0]['note']!r}, verdict=done")


def test_diff_move_and_insert():
    # Change A p.2: move_after обязан дать РОВНО одну запись diff'а на
    # переместившийся блок; insert_after не должен порождать НИКАКИХ записей
    # с пометкой перемещения — множество id там другое (появился новый), а
    # не просто переставленное старое.
    from docx_editor.edit import _diff

    doc = Document()
    for t in ("Первый", "Второй", "Третий", "Четвёртый"):
        doc.add_paragraph(t)
    idx = index(doc)
    before = doc_map(doc, idx)

    patch.apply(doc, idx, {"op": "move_after", "id": "p0", "after": "p2"})
    move_diff = _diff(before, doc_map(doc, idx))
    move_entries = [d for d in move_diff if "перемещ" in d.get("note", "")]
    assert len(move_entries) == 1, move_diff
    assert move_entries[0]["id"] == "p0", move_diff

    doc2 = Document()
    for t in ("Первый", "Второй", "Третий", "Четвёртый"):
        doc2.add_paragraph(t)
    idx2 = index(doc2)
    before2 = doc_map(doc2, idx2)
    patch.apply(doc2, idx2, {"op": "insert_after", "id": "p1", "text": "Вставленный", "style": "Normal"})
    insert_diff = _diff(before2, doc_map(doc2, idx2))
    spurious = [d for d in insert_diff if "перемещ" in d.get("note", "")]
    assert not spurious, insert_diff
    print("edit_demo: move_after даёт ровно одну запись в diff, insert_after не даёт ложных move-записей")


def test_rule_fallthrough_to_editor_done():
    # Change C: rule, не найдя типографских дефектов, раньше был терминальным
    # "already", даже когда правка вообще не про типографику (Математика
    # edit 1 — двойной пробел и висящий пробел остались в документе, а rule
    # искал не тот класс дефектов и смолчал). Пустой diff у rule теперь не
    # терминал: код восстанавливает документ и продолжает ТЕМ ЖЕ запросом
    # обычным локальным путём. Здесь rule ожидаемо ничего не находит
    # (никаких типографских дефектов), а фейковый Редактор честно предлагает
    # рабочую операцию — итог обязан быть done, а не already.
    doc = Document()
    doc.add_paragraph("Обычный абзац без опечаток.")
    idx = index(doc)

    def fake_navigator(outline_text, request):
        return {"kind": "local", "rule": "typography", "ids": ["p0"], "anchors": []}

    def fake_editor(fragment_text, request, feedback=None):
        return {"ops": [{"op": "replace_text", "id": "p0", "old": "Обычный", "new": "Другой"}]}

    def fake_checker(request, diff):
        return {"ok": True, "reason": "ok"}

    result, doc, idx = run_edit(
        doc, idx, "замени 'Обычный' на 'Другой'",
        navigator=fake_navigator, editor=fake_editor, checker=fake_checker,
    )

    assert result["verdict"] == "done", result
    after_text = next(b["text"] for b in doc_map(doc, idx) if b["id"] == "p0")
    assert after_text == "Другой абзац без опечаток.", after_text
    print("edit_demo: rule без изменений упал в локальный путь, Редактор довёл правку до done")


def test_rule_fallthrough_editor_declines_stays_already():
    # Тот же fallthrough, но Редактор честно отказывается (пустые ops) — оба
    # пути (rule и Редактор) согласны, что менять нечего, already здесь
    # легитимен (в отличие от отказа validate/Проверяющего — item C.4).
    doc = Document()
    doc.add_paragraph("Обычный абзац без опечаток.")
    idx = index(doc)
    before = [b["text"] for b in doc_map(doc, idx)]

    def fake_navigator(outline_text, request):
        return {"kind": "local", "rule": "typography", "ids": ["p0"], "anchors": []}

    def fake_editor(fragment_text, request, feedback=None):
        return {"ops": []}

    def fake_checker(request, diff):
        raise AssertionError("Проверяющий не должен вызываться — Редактор честно отказался")

    result, doc, idx = run_edit(
        doc, idx, "замени 'Обычный' на 'Другой'",
        navigator=fake_navigator, editor=fake_editor, checker=fake_checker,
    )

    assert result["verdict"] == "already", result
    after = [b["text"] for b in doc_map(doc, idx)]
    assert after == before, "already не должен менять документ"
    print(f"edit_demo: fallthrough с честным отказом Редактора даёт already: {result['reason']!r}")


def test_edit_12_on_real_doc_fixture():
    # Раньше это был единственный живой вызов модели во всём демо — упирался в
    # реальный API при каждом прогоне run_all.py (см. CLAUDE.md: демо,
    # закрываемое фикстурой, закрываем фикстурой). Живая приёмка ровно этого
    # же корпуса и этой же правки №12 остаётся в examples/colbert_run.py (не
    # *_demo.py, run_all.py его не подхватывает) — там весь прогон из 20 правок
    # проверяет тот же инвариант "вердикт правдив" по-настоящему, живой
    # моделью. Здесь — фейковые Навигатор/Редактор/Проверяющий в форме их
    # реального JSON-ответа (как в остальных тестах файла), но реальный
    # документ, реальный текст правки №12 (через split реального файла правок)
    # и реальный p76: инвариант — verdict=done обязан означать, что p76
    # реально поменялся, а остальные блоки документа не тронуты.
    doc = Document(REAL_DOC)
    idx = index(doc)
    blocks = doc_map(doc, idx)
    p76_before = next(b["text"] for b in blocks if b["id"] == "p76")
    other_before = [b["text"] for b in blocks if b["id"] != "p76"]

    request = split(open(EDITS_FILE, encoding="utf-8").read())[11]
    old = "на порядок (в 6–10 раз)"
    new = "более чем в 6–10 раз"
    assert old in p76_before, p76_before

    def fake_navigator(outline_text, request):
        return {"kind": "local", "rule": None, "ids": ["p76"], "anchors": []}

    def fake_editor(fragment_text, request, feedback=None):
        # Правка цитирует ещё «на два порядка (в ~170 раз)», и _resolve честно
        # доливает по этой цитате приблизительное совпадение где-то ещё в
        # документе (locate) — с кластеризацией это отдельный, СВОЙ вызов
        # Редактора на СВОЁМ маленьком фрагменте. Фейк обязан быть таким же
        # избирательным, как настоящий: отвечать только когда p76 реально
        # виден в ЭТОМ фрагменте, иначе честно "нечего делать здесь".
        if "p76 " not in fragment_text:
            return {"ops": []}
        return {"ops": [{"op": "replace_text", "id": "p76", "old": old, "new": new}]}

    def fake_checker(request, diff):
        return {"ok": True, "reason": "кратность и формулировка теперь согласованы"}

    result, doc, idx = run_edit(
        doc, idx, request,
        navigator=fake_navigator, editor=fake_editor, checker=fake_checker,
    )

    assert result["verdict"] == "done", result
    p76_after = next(b["text"] for b in doc_map(doc, idx) if b["id"] == "p76")
    assert p76_after != p76_before and old not in p76_after and new in p76_after, p76_after
    other_after = [b["text"] for b in doc_map(doc, idx) if b["id"] != "p76"]
    assert other_after == other_before, "правка №12 не должна трогать остальные блоки документа"
    print(f"edit_demo: правка №12 на реальном документе (фикстура): p76 было {p76_before!r}, стало {p76_after!r}")


def test_scattered_targets_three_calls_one_verdict():
    # Спецификация кластеризации: цели в трёх разных частях документа (индексы
    # 0/5/10 из 12 абзацев — при around=1 окна p0/p5/p10 не пересекаются)
    # обязаны дать ТРИ маленьких вызова Редактора (каждый видит ровно одну
    # цель плюс соседей), но ОДИН снимок/diff/вердикт на всю правку — цикл
    # дробит шаги, не обязательство (см. _run_clusters/_finish).
    doc = Document()
    for i in range(12):
        doc.add_paragraph(f"Абзац {i}: обычный текст.")
    idx = index(doc)
    target_ids = ["p0", "p5", "p10"]

    def fake_navigator(outline_text, request):
        return {"kind": "local", "rule": None, "ids": target_ids, "anchors": []}

    calls = []

    def fake_editor(fragment_text, request, feedback=None):
        calls.append(fragment_text)
        present = [t for t in target_ids if f"{t} " in fragment_text]
        assert len(present) == 1, f"кластер обязан нести ровно одну цель: {fragment_text!r}"
        return {"ops": [{"op": "replace_text", "id": present[0], "old": "обычный", "new": "изменённый"}]}

    checker_calls = []

    def fake_checker(request, diff):
        checker_calls.append(diff)
        assert len(diff) == 3, diff
        return {"ok": True, "reason": "ok"}

    result, doc, idx = run_edit(
        doc, idx, "замени 'обычный' на 'изменённый' в трёх местах",
        navigator=fake_navigator, editor=fake_editor, checker=fake_checker,
    )

    assert len(calls) == 3, f"три разнесённых блока обязаны дать три вызова Редактора, получили {len(calls)}"
    assert len(checker_calls) == 1, "Проверяющий обязан вызываться РОВНО один раз на всю правку"
    assert result["verdict"] == "done", result
    assert result["iter"] == 3, result
    print("edit_demo: три разнесённые цели дали три маленьких вызова Редактора и один общий вердикт")


def test_adjacent_targets_one_call():
    # Клаcтеризация: два соседних абзаца (p0/p1, окна пересекаются) — ОДИН
    # кластер и, значит, ОДИН вызов Редактора, а не два.
    doc = Document()
    doc.add_paragraph("Абзац A: обычный.")
    doc.add_paragraph("Абзац B: обычный.")
    idx = index(doc)

    def fake_navigator(outline_text, request):
        return {"kind": "local", "rule": None, "ids": ["p0", "p1"], "anchors": []}

    calls = []

    def fake_editor(fragment_text, request, feedback=None):
        calls.append(fragment_text)
        return {"ops": [
            {"op": "replace_text", "id": "p0", "old": "обычный", "new": "другой"},
            {"op": "replace_text", "id": "p1", "old": "обычный", "new": "другой"},
        ]}

    def fake_checker(request, diff):
        return {"ok": True, "reason": "ok"}

    result, doc, idx = run_edit(
        doc, idx, "замени 'обычный' на 'другой' в обоих абзацах",
        navigator=fake_navigator, editor=fake_editor, checker=fake_checker,
    )

    assert len(calls) == 1, f"два соседних абзаца обязаны дать ОДИН вызов Редактора, получили {len(calls)}"
    assert result["verdict"] == "done", result
    print("edit_demo: два соседних резолвленных id слились в один кластер — один вызов Редактора")


def test_empty_ops_cluster_skipped_others_apply():
    # Кластер, где Редактор честно вернул пустые ops ("нечего делать здесь"),
    # не должен проваливать правку — другие кластеры всё ещё могут дать работу,
    # и итоговый вердикт обязан быть done.
    doc = Document()
    for i in range(8):
        doc.add_paragraph(f"Абзац {i}: обычный текст.")
    idx = index(doc)

    def fake_editor(fragment_text, request, feedback=None):
        if "p0 " in fragment_text:
            return {"ops": []}
        return {"ops": [{"op": "replace_text", "id": "p7", "old": "обычный", "new": "другой"}]}

    def fake_navigator(outline_text, request):
        return {"kind": "local", "rule": None, "ids": ["p0", "p7"], "anchors": []}

    def fake_checker(request, diff):
        return {"ok": True, "reason": "ok"}

    result, doc, idx = run_edit(
        doc, idx, "замени 'обычный' на 'другой' там, где уместно",
        navigator=fake_navigator, editor=fake_editor, checker=fake_checker,
    )

    assert result["verdict"] == "done", result
    p7_text = next(b["text"] for b in doc_map(doc, idx) if b["id"] == "p7")
    assert "другой" in p7_text, p7_text
    print("edit_demo: пустой ops одного кластера пропущен, другой кластер применился, вердикт done")


def test_invalid_cluster_aborts_whole_edit():
    # Кластер, чьи ops не проходят валидацию даже после ретрая внутри
    # _apply_ops, обязан прервать ВСЮ правку — включая откат уже применённого
    # предыдущего кластера. Атомарность остаётся на уровне всей правки, а не
    # кластера (см. спецификацию задачи: "шаги мельче, обязательство то же").
    doc = Document()
    doc.add_paragraph("Абзац A: обычный текст.")
    for i in range(1, 5):
        doc.add_paragraph(f"Абзац {i}: наполнитель.")
    doc.add_paragraph("Абзац B: обычный текст.")
    idx = index(doc)
    before = [b["text"] for b in doc_map(doc, idx)]

    def fake_navigator(outline_text, request):
        return {"kind": "local", "rule": None, "ids": ["p0", "p5"], "anchors": []}

    def fake_editor(fragment_text, request, feedback=None):
        if "p0 " in fragment_text:
            return {"ops": [{"op": "replace_text", "id": "p0", "old": "обычный", "new": "другой"}]}
        # p5-кластер: "old" которого в блоке нет — невалидно что при первой
        # попытке, что при ретрае (feedback игнорируется нарочно).
        return {"ops": [{"op": "replace_text", "id": "p5", "old": "текста, которого тут нет", "new": "х"}]}

    def fake_checker(request, diff):
        raise AssertionError("Проверяющий не должен вызываться — правка обязана прерваться раньше")

    result, doc, idx = run_edit(
        doc, idx, "замени 'обычный' на 'другой'",
        navigator=fake_navigator, editor=fake_editor, checker=fake_checker,
    )

    assert result["verdict"] == "failed", result
    after = [b["text"] for b in doc_map(doc, idx)]
    assert after == before, "невалидный кластер обязан откатить ВСЮ правку, включая уже применённый первый кластер"
    print("edit_demo: невалидный кластер после ретрая откатил всю правку, документ побайтово как был")


def test_many_scattered_targets_no_cap():
    # Владелец явно снял верхнюю границу на число целей: сколько бы
    # резолвленных id ни было, каждый непересекающийся кластер даёт свой
    # вызов Редактора — без кэпа и без отката на "один большой фрагмент".
    doc = Document()
    n_targets, total_paras = 15, 45
    for i in range(total_paras):
        doc.add_paragraph(f"Абзац {i}: наполнитель.")
    idx = index(doc)
    target_ids = [f"p{i}" for i in range(0, total_paras, 3)]
    assert len(target_ids) == n_targets

    def fake_navigator(outline_text, request):
        return {"kind": "local", "rule": None, "ids": target_ids, "anchors": []}

    calls = []

    def fake_editor(fragment_text, request, feedback=None):
        calls.append(fragment_text)
        present = [t for t in target_ids if f"{t} " in fragment_text]
        assert len(present) == 1, fragment_text
        return {"ops": [{"op": "replace_text", "id": present[0], "old": "наполнитель", "new": "правка"}]}

    def fake_checker(request, diff):
        assert len(diff) == n_targets, diff
        return {"ok": True, "reason": "ok"}

    result, doc, idx = run_edit(
        doc, idx, "замени 'наполнитель' на 'правка' везде, где нужно",
        navigator=fake_navigator, editor=fake_editor, checker=fake_checker,
    )

    assert len(calls) == n_targets, f"{n_targets} разнесённых целей без кэпа обязаны дать {n_targets} вызовов, получили {len(calls)}"
    assert result["verdict"] == "done", result
    assert result["iter"] == n_targets, result
    print(f"edit_demo: {n_targets} разнесённых целей — {len(calls)} вызовов Редактора, кэп на 12 отсутствует")


def test_out_of_lane_op_blocked_document_unchanged():
    # Находка: _apply_ops звал patch.validate с картой ВСЕГО документа, поэтому
    # операция на любой существующий id проходила, даже если Редактор видел
    # только свой фрагмент. Редактор, которому показали окрестность p0, но
    # который отвечает операцией на далёкий p9, обязан быть отклонён кодом, а
    # не применён — документ должен остаться побайтово как был, verdict != done.
    doc = Document()
    for i in range(10):
        doc.add_paragraph(f"Абзац {i}: обычный текст.")
    idx = index(doc)
    before = [b["text"] for b in doc_map(doc, idx)]

    def fake_navigator(outline_text, request):
        return {"kind": "local", "rule": None, "ids": ["p0"], "anchors": []}

    def fake_editor(fragment_text, request, feedback=None):
        assert "p9 " not in fragment_text, "p9 не должен попадать во фрагмент вокруг p0"
        return {"ops": [{"op": "replace_text", "id": "p9", "old": "обычный", "new": "чужой"}]}

    def fake_checker(request, diff):
        raise AssertionError("Проверяющий не должен вызываться — операция вне фрагмента обязана быть отклонена раньше")

    result, doc, idx = run_edit(
        doc, idx, "замени 'обычный' на 'чужой' в p0",
        navigator=fake_navigator, editor=fake_editor, checker=fake_checker,
    )

    assert result["verdict"] != "done", result
    assert "вне фрагмента" in result["reason"], result
    after = [b["text"] for b in doc_map(doc, idx)]
    assert after == before, "документ обязан остаться побайтово как был — блок вне фрагмента не должен меняться"
    print(f"edit_demo: операция на блок вне фрагмента отклонена, документ не тронут: {result['reason']!r}")


def test_out_of_lane_retry_recovers_within_fragment():
    # Тот же случай, но Редактор, получив ошибку об "вне фрагмента" как
    # feedback, на ретрае отвечает верно — в пределах СВОЕГО фрагмента.
    # Единственный существующий ретрай внутри _apply_ops обязан это спасти.
    doc = Document()
    for i in range(10):
        doc.add_paragraph(f"Абзац {i}: обычный текст.")
    idx = index(doc)

    def fake_navigator(outline_text, request):
        return {"kind": "local", "rule": None, "ids": ["p0"], "anchors": []}

    def fake_editor(fragment_text, request, feedback=None):
        if feedback is None:
            return {"ops": [{"op": "replace_text", "id": "p9", "old": "обычный", "new": "чужой"}]}
        assert "вне фрагмента" in feedback, feedback
        return {"ops": [{"op": "replace_text", "id": "p0", "old": "обычный", "new": "исправленный"}]}

    def fake_checker(request, diff):
        return {"ok": True, "reason": "ok"}

    result, doc, idx = run_edit(
        doc, idx, "замени 'обычный' на 'исправленный' в p0",
        navigator=fake_navigator, editor=fake_editor, checker=fake_checker,
    )

    assert result["verdict"] == "done", result
    p9_text = next(b["text"] for b in doc_map(doc, idx) if b["id"] == "p9")
    assert "чужой" not in p9_text, p9_text
    p0_text = next(b["text"] for b in doc_map(doc, idx) if b["id"] == "p0")
    assert "исправленный" in p0_text, p0_text
    print("edit_demo: после ретрая с feedback Редактор попал в свой фрагмент, verdict=done")


def test_replace_all_not_blocked_by_lane_guard():
    # replace_all документ-широкая ПО КОНТРАКТУ и id не несёт — гвард обязан
    # её пропускать, хотя формально она меняет блоки вне показанного фрагмента.
    doc = Document()
    for i in range(10):
        doc.add_paragraph(f"Абзац {i}: обычный текст.")
    idx = index(doc)

    def fake_navigator(outline_text, request):
        return {"kind": "local", "rule": None, "ids": ["p0"], "anchors": []}

    def fake_editor(fragment_text, request, feedback=None):
        return {"ops": [{"op": "replace_all", "old": "обычный", "new": "новый"}]}

    def fake_checker(request, diff):
        return {"ok": True, "reason": "ok"}

    result, doc, idx = run_edit(
        doc, idx, "замени 'обычный' на 'новый' везде",
        navigator=fake_navigator, editor=fake_editor, checker=fake_checker,
    )

    assert result["verdict"] == "done", result
    p9_text = next(b["text"] for b in doc_map(doc, idx) if b["id"] == "p9")
    assert "новый" in p9_text, p9_text
    print("edit_demo: replace_all не заблокирован гвардом фрагмента, применился по всему документу")


def test_op_on_block_created_in_same_batch_applies():
    # insert_after создаёт свежий id — операция на него в ТОМ ЖЕ батче обязана
    # быть разрешена, хотя этого id не было во фрагменте, который видел
    # Редактор изначально (см. пополнение fragment_ids в _apply_ops).
    doc = Document()
    doc.add_paragraph("Первый абзац.")
    doc.add_paragraph("Второй абзац.")
    idx = index(doc)

    def fake_navigator(outline_text, request):
        return {"kind": "local", "rule": None, "ids": ["p0"], "anchors": []}

    def fake_editor(fragment_text, request, feedback=None):
        return {"ops": [
            {"op": "insert_after", "id": "p0", "text": "Новый абзац.", "style": "Normal"},
            {"op": "set_style", "id": "p2", "style": "Heading 1"},
        ]}

    def fake_checker(request, diff):
        return {"ok": True, "reason": "ok"}

    result, doc, idx = run_edit(
        doc, idx, "вставь абзац после первого и сделай его заголовком",
        navigator=fake_navigator, editor=fake_editor, checker=fake_checker,
    )

    assert result["verdict"] == "done", result
    new_block = next(b for b in doc_map(doc, idx) if b["id"] == "p2")
    assert new_block["text"] == "Новый абзац.", new_block
    assert new_block["style"] == "Heading 1", new_block
    print("edit_demo: операция на блок, созданный этим же батчем, применилась, verdict=done")


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
    test_set_style_diff_reaches_checker()
    test_diff_move_and_insert()
    test_rule_fallthrough_to_editor_done()
    test_rule_fallthrough_editor_declines_stays_already()
    test_edit_12_on_real_doc_fixture()
    test_scattered_targets_three_calls_one_verdict()
    test_adjacent_targets_one_call()
    test_empty_ops_cluster_skipped_others_apply()
    test_invalid_cluster_aborts_whole_edit()
    test_many_scattered_targets_no_cap()
    test_out_of_lane_op_blocked_document_unchanged()
    test_out_of_lane_retry_recovers_within_fragment()
    test_replace_all_not_blocked_by_lane_guard()
    test_op_on_block_created_in_same_batch_applies()
