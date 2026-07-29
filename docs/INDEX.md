[← Документация](README.md) · [Открытый план](../BUILD_PLAN.md) · [Логбук](log/README.md)

# Справочник тем

Знаешь слово — найди файл. Колонка «код» — где это живёт, «правила» — где
действующий контракт, «история» — где разбор, как оно появилось и что при
этом сломалось. Открывай ОДНУ ячейку, а не всю строку.

## Разбор документа

| Тема | Код | Правила | История |
|---|---|---|---|
| Блок, `doc_map`, оформление ранов, уровень/список/сноски/поля в блоке | `docx_editor/parse.py` | CONVENTIONS «Блок» | [Ф1–Ф4](log/phase-01-04-foundation.md); level/list — [Ф9](log/phase-08-10-acceptance-findings.md); усечение `render()` — [Ф12](log/phase-12-13-corruption-completeness.md) (Н35); счётчик сносок — [Ф19-бис](log/phase-17-21-routing.md) (Н50) |
| Поиск: подстрока, regex, нечёткий якорь, оглавление, фрагмент с соседями | `docx_editor/find.py` | — | [Ф4](log/phase-01-04-foundation.md) (Н5–Н7); `find.locate` — [Ф11](log/measurement-w4-w12.md) (Н32); `pattern`-адрес — [В5](log/driving-run-2026-07-29.md) (Н61) |
| Типографские правила (`normalize`: пробелы, слипшиеся предложения, невидимые символы, кавычки) | `docx_editor/rules.py` | CONVENTIONS «Патч», `normalize` | [Ф3](log/phase-01-04-foundation.md) |

## Патч и применение

| Тема | Код | Правила | История |
|---|---|---|---|
| Валидатор и применение всех 21 операции патча | `docx_editor/patch.py` | CONVENTIONS «Контракты», полный список операций | [Ф2](log/phase-01-04-foundation.md) (Н1–Н4); таблицы/строки/колонки, `\n`-гвард — [В2](log/driving-run-2026-07-29.md) (Н56); поля/колонтитулы — [В2-бис](log/driving-run-2026-07-29.md) (Н57–Н58) |
| Сноски Word (`footnote`, часть `footnotes.xml`) | `docx_editor/patch.py` | CONVENTIONS «Патч» | [Ф15](log/phase-14-16-research-operations.md) (Н44); попадание в промпт — [Ф17](log/phase-14-16-research-operations.md) (Н45) |
| `set_text` с сохранением начертания ранов | `docx_editor/patch.py::_op_set_text` | CONVENTIONS «локальная» ветка маршрутизации | [Ф19](log/phase-17-21-routing.md); диффом по opcodes — [В3](log/driving-run-2026-07-29.md) (Н59) |

## Цикл правки и маршрутизация

| Тема | Код | Правила | История |
|---|---|---|---|
| Навигатор → поиск → Редактор → валидатор → применение → Проверка | `docx_editor/edit.py` | CONVENTIONS «Конвейер» | [Ф5](log/phase-05-07-edit-cycle.md); безопасность отката — [Ф8](log/phase-08-10-acceptance-findings.md) (Н18) |
| Кластеризация, полоса Редактора, самоколлизии батча (`written`) | `docx_editor/edit.py` | CONVENTIONS «локальная» ветка | [Ф11](log/measurement-w4-w12.md) (Н33–Н34); Ф12 (Н36); Ф13/Ф13-бис (Н37–Н43) |
| Маршрут `rule` / `already` | `docx_editor/edit.py` | CONVENTIONS «Маршрутизация», п. 1 | [Ф8](log/phase-08-10-acceptance-findings.md) (Н15–Н16); [Ф11](log/measurement-w4-w12.md) (Н31) |
| Маршрут `unify` (лексический контракт) — **статус не решён** | `docx_editor/edit.py::_run_unify` | CONVENTIONS «Маршрутизация», п. 2 | [Ф18](log/phase-17-21-routing.md) (Н46–Н48); русское словоизменение — [Ф19-бис/В6](log/phase-17-21-routing.md) [и](log/driving-run-2026-07-29.md) (Н49, Н62, Н69) |
| Маршрут `compose` (составные правки, один вызов вместо кластеров) | `docx_editor/edit.py` | CONVENTIONS «Маршрутизация», п. 4 | [Ф20](log/phase-17-21-routing.md) |
| Проверка следа (`trace`) — код проверяет СДЕЛАННОЕ, не ярлык | `docx_editor/edit.py::_trace_error` | CONVENTIONS «Проверка следа» | [В1](log/driving-run-2026-07-29.md) (Н53, Н55); инверсия и починка — [В11](log/driving-run-2026-07-29.md) (Н66) |

## Модель и сервер

| Тема | Код | Правила | История |
|---|---|---|---|
| Один вызов модели, `.env`, JSON-режим, ретрай | `docx_editor/llm.py` | CLAUDE.md «Окружение» | [Ф5](log/phase-05-07-edit-cycle.md); контекстное переполнение — [Ф13-бис](log/phase-12-13-corruption-completeness.md) (Н43, открыто) |
| Пять эндпоинтов, поток `/edit`, сердцебиение, ретрай транспорта | `docx_editor/server.py` | CONVENTIONS «Эндпоинты», «Событие потока» | [Ф1](log/phase-01-04-foundation.md); heartbeat — [Ф10](log/phase-08-10-acceptance-findings.md) (Н21); ретрай в продукте — [В4](log/driving-run-2026-07-29.md) (Н60) |
| Журнал JSONL, `GET /logs` | `docx_editor/log.py` | CONVENTIONS «Событие потока» (поля из `renderLogRecord`) | [Ф6](log/phase-05-07-edit-cycle.md); дыра авторизации — [Ф10](log/phase-08-10-acceptance-findings.md) (Н23, открыто) |
| Веб-интерфейс (`web/index.html`) — внешний, не менять формат блока | `web/index.html` | CONVENTIONS «Блок» (зафиксирован снаружи) | [Ф1](log/phase-01-04-foundation.md), [Ф6](log/phase-05-07-edit-cycle.md); заголовки/списки в просмотрщике — [Ф9](log/phase-08-10-acceptance-findings.md) (Н19) |

## Замер и приёмка

| Тема | Где | Правила | История |
|---|---|---|---|
| Стенд: `run.py` (изолированно), `run_seq.py` (накопительно), `oracle.py`, `score.py`, `report.py` | `bench/` | CLAUDE.md «Метод: замер решает» | [Ф11](log/measurement-w4-w12.md) — стенд заведён |
| Оценщик (`oracle.py`) — маркеры по файлу, независимые от вердикта системы | `bench/oracle.py` | — | Починка `m19`/`c7`, пробел `n/a` на 21/40 — [В8](log/driving-run-2026-07-29.md) (Н63–Н65) |
| Приёмочные корпуса: ColBERT, Математика, `battery` (третий, незнакомый), `Регламент` (структурный) | `bench/fixtures/`, вне репо | CONVENTIONS «Мерило» | ColBERT/Математика — [Ф7](log/phase-05-07-edit-cycle.md); `battery` — [Ф19-бис](log/phase-17-21-routing.md) (Н49); `Регламент` — [прогон за рулём](log/driving-run-2026-07-29.md) (Н54) |
| Отчёт по эффективности (график Э по партиям w4–w15) | `bench/ЭФФЕКТИВНОСТЬ.md`/`.html` | — | [измерения w4–w12](log/measurement-w4-w12.md), [Ф18](log/phase-17-21-routing.md) |
| Разбор чужой системы SuperDoc (что взять, что нет) | `bench/superdoc-разбор.md` | — | [Ф14](log/phase-14-16-research-operations.md) |
| Сырые числа оценщика по партиям | `bench/РЕЗУЛЬТАТЫ.txt` | — | [измерения w4–w12](log/measurement-w4-w12.md) |

## Процесс

| Тема | Где |
|---|---|
| Регламент сессии, инварианты, «сначала диагноз, потом ТЗ» | [../CLAUDE.md](../CLAUDE.md) |
| Указатель находок Н1–Н70 | [log/README.md](log/README.md) |
| Живые прогоны (приёмка, показ владельцу, `battery`, прогон за рулём) | [log/README.md](log/README.md) |
| Открытая работа (что делать дальше) | [../BUILD_PLAN.md](../BUILD_PLAN.md) |
