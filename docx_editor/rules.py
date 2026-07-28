"""Две детерминированные нормализации текста, без модели: typography и quotes.

Правят живой документ на месте через _ptext/_replace_span из patch.py — ту же
машинерию, что использует replace_all (раны, гиперссылки, ячейки таблиц), чтобы
не заводить второй механизм редактирования текста.
"""

import re

from docx.oxml.ns import qn

from docx_editor.patch import _ptext, _replace_span

# Пробел (включая U+00A0 и другие юникодные пробелы) перед пунктуацией.
_WS_BEFORE_PUNCT = re.compile(r"[^\S\n]+(?=[,.:;!?)\]}»])")

# Пропущенный пробел между предложениями: буква любого регистра + точка + заглавная.
# Защита от ложных срабатываний на десятичных числах (4.2, 0.4448) — не регистр
# слева, а то, что классы букв не включают цифры: цифра по обе стороны от точки
# в принципе не может совпасть с [A-Za-zА-Яа-яЁё] / [A-ZА-ЯЁ].
_GLUED_SENTENCE = re.compile(r"([A-Za-zА-Яа-яЁё])\.([A-ZА-ЯЁ])")

# Двойной (и более) пробел внутри строки и висячий пробел в конце абзаца.
# ПАРТИЯ 3: Математика 1 просит убрать оба названных места «и заодно проверить
# весь текст на такое же». Проверено замером w9: без этого правка уходит
# Редактору, тот чинит два процитированных места и не доходит до блока 50 —
# вердикт «сделано» при невыполненной правке, первая ложь с w5. Смысл здесь не
# нужен ни на грамм, значит это работа кода (инвариант 1).
_DOUBLE_SPACE = re.compile(r"[^\S\n]{2,}")
_TRAILING_SPACE = re.compile(r"[^\S\n]+$")

# Невидимые форматирующие символы: word joiner, zero-width space, BOM, мягкий перенос.
_INVISIBLE = re.compile(r"[⁠​﻿­]+")


def _paragraphs(doc):
    """Все w:p документа, включая абзацы внутри ячеек таблиц (как replace_all)."""
    return doc.element.body.iter(qn("w:p"))


def typography(doc):
    """Пять правил по всему документу: пробел перед пунктуацией, слипшиеся
    предложения, невидимые символы, двойной пробел, висячий пробел в конце
    абзаца. Правки внутри абзаца — справа налево, иначе более ранние замены
    сдвигают смещения более поздних."""
    ws = glued = invisible = doubled = trailing = 0
    for p_el in _paragraphs(doc):
        text = _ptext(p_el)
        matches = list(_WS_BEFORE_PUNCT.finditer(text))
        for m in reversed(matches):
            _replace_span(p_el, m.start(), m.end(), "")
        ws += len(matches)

        text = _ptext(p_el)
        matches = list(_GLUED_SENTENCE.finditer(text))
        for m in reversed(matches):
            _replace_span(p_el, m.start(), m.end(), f"{m.group(1)}. {m.group(2)}")
        glued += len(matches)

        # двойной пробел — до висячего: схлопнутый хвост «…слово  » станет
        # «…слово », и его добьёт следующее правило
        text = _ptext(p_el)
        matches = list(_DOUBLE_SPACE.finditer(text))
        for m in reversed(matches):
            _replace_span(p_el, m.start(), m.end(), " ")
        doubled += len(matches)

        text = _ptext(p_el)
        matches = list(_TRAILING_SPACE.finditer(text))
        for m in reversed(matches):
            _replace_span(p_el, m.start(), m.end(), "")
        trailing += len(matches)

        text = _ptext(p_el)
        matches = list(_INVISIBLE.finditer(text))
        for m in reversed(matches):
            _replace_span(p_el, m.start(), m.end(), "")
        invisible += len(matches)

    return (
        f"Типография: убрано пробелов перед пунктуацией — {ws}, "
        f"разделено слипшихся предложений — {glued}, "
        f"вырезано невидимых символов — {invisible}, "
        f"схлопнуто двойных пробелов — {doubled}, "
        f"убрано висячих пробелов в конце абзаца — {trailing}"
    )


def quotes(doc):
    """Прямые " заменяются на ёлочки: чётные по порядку в абзаце — «, нечётные — ».

    В документе вложенных кавычек внутри ёлочек нет, число прямых кавычек в
    каждом абзаце чётно, поэтому простое чередование в порядке появления даёт
    корректную расстановку открывающих/закрывающих.
    """
    replaced = 0
    for p_el in _paragraphs(doc):
        text = _ptext(p_el)
        positions = [i for i, ch in enumerate(text) if ch == '"']
        for order, pos in reversed(list(enumerate(positions))):
            char = "«" if order % 2 == 0 else "»"
            _replace_span(p_el, pos, pos + 1, char)
        replaced += len(positions)

    return f"Кавычки: заменено на ёлочки — {replaced}"
