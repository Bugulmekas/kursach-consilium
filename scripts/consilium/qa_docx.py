#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Универсальная QA-проверка .docx перед защитой.

Проверяет не смысл, а формальные признаки:
- факты о документе (страницы, секции, таблицы, рисунки, колонтитулы);
- незаполненные заглушки;
- покрытие ссылок [n] и диапазон источников;
- дубли заголовков;
- подписи таблиц и рисунков;
- красный текст;
- обязательные и запрещённые фразы.

Пример:
    python -X utf8 qa_docx.py thesis.docx --sources 20
    python -X utf8 qa_docx.py thesis.docx --config qa.json
    python -X utf8 qa_docx.py thesis.docx --require "ПРИЛОЖЕНИЕ А" --forbid "TODO"

Код возврата 1, если есть FAIL.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

from docx_tools import CITE_RE, doc_info, is_toc, iter_paragraphs, load_doc, paragraph_style

DEFAULT_PLACEHOLDERS = [
    "TODO",
    "TBD",
    "FIXME",
    "ЗАМЕНИТЬ",
    "Ваша тема",
    "???",
    "XXX",
    "[вставьте",
]

RESULTS: list[tuple[str, str, str]] = []


def add(status: str, name: str, detail: str = "") -> None:
    RESULTS.append((status, name, detail))


def run(args) -> int:
    path = Path(args.path)
    info = doc_info(path)
    entries, doc = load_doc(path)
    paras = [(paragraph_style(p), t, is_toc(p)) for _, _, p, t in iter_paragraphs(doc)]
    text = "\n".join(t for _, t, _ in paras)

    cfg = {}
    if args.config:
        cfg = json.loads(Path(args.config).read_text(encoding="utf-8"))
    sources = args.sources or cfg.get("sources")
    placeholders = cfg.get("placeholders", DEFAULT_PLACEHOLDERS)
    require = list(cfg.get("require", [])) + list(args.require or [])
    forbid = list(cfg.get("forbid", [])) + list(args.forbid or [])

    print("== Факты ==")
    print(f"Страниц: {info['pages']} | слов: {info['words']} | секций: {info['sections']}")
    print(f"Таблиц: {info['tables']} | рисунков: {info['drawings']} | titlePg: {info['title_page']}")
    print(f"Полей TOC: {info['toc_fields']}")
    for h in info["headers"]:
        print(
            f"  {h['name']}: PAGE={h['page_fields']} "
            f"NUMPAGES={h['numpages_fields']} таблиц={h['tables']}"
        )
    print()

    if info["pages"] == "?":
        add("WARN", "Счётчик страниц", "app.xml без <Pages>: сохраните документ в Word")
    else:
        add("PASS", "Счётчик страниц", f"{info['pages']} стр.")

    add("PASS" if info["sections"] > 0 else "FAIL", "Секции", str(info["sections"]))

    # заглушки
    found_ph = []
    low = text.lower()
    for ph in placeholders:
        if ph.lower() in low:
            found_ph.append(ph)
    add(
        "FAIL" if found_ph else "PASS",
        "Заглушки",
        f"найдены: {', '.join(found_ph)}" if found_ph else "не найдены",
    )

    # цитирование
    nums = set()
    for group in CITE_RE.findall(text):
        for part in group.split(","):
            part = part.strip().replace("–", "-")
            if "-" in part:
                a, b = part.split("-", 1)
                if a.strip().isdigit() and b.strip().isdigit():
                    nums.update(range(int(a), int(b) + 1))
            elif part.isdigit():
                nums.add(int(part))
    if not nums:
        add("WARN", "Ссылки [n]", "не найдены")
    elif sources:
        expected = set(range(1, int(sources) + 1))
        missing = sorted(expected - nums)
        extra = sorted(nums - expected)
        detail = f"использовано {len(nums)} из {sources}"
        if missing:
            detail += f"; не цитируются: {missing}"
        if extra:
            detail += f"; вне диапазона: {extra}"
        add("FAIL" if missing or extra else "PASS", "Цитирование", detail)
    else:
        add("PASS", "Ссылки [n]", f"использовано {len(nums)}, максимум {max(nums)}")

    # дубли заголовков
    headings = []
    for style, t, toc in paras:
        s = t.strip()
        if not s or len(s) > 140 or toc:
            continue
        is_heading = any(k in style.lower() for k in ("heading", "заголов"))
        looks_numbered = bool(
            re.match(r"^\d+(\.\d+)*\.?\s+[А-ЯA-Z]", s)
        )
        if is_heading or looks_numbered:
            headings.append(s)
    seen, dups = set(), []
    for h in headings:
        key = h.lower()
        if key in seen:
            dups.append(h)
        seen.add(key)
    add(
        "WARN" if dups else "PASS",
        "Дубли заголовков",
        ("; ".join(dups[:5]) + " (проверьте, не оглавление ли)") if dups else "нет",
    )

    # подписи таблиц
    caps = sum(1 for _, t, _ in paras if re.match(r"^\s*Таблица\s+\d+", t, re.I))
    if info["tables"] == 0:
        add("PASS", "Подписи таблиц", "таблиц нет")
    elif caps < info["tables"]:
        add("WARN", "Подписи таблиц", f"таблиц {info['tables']}, подписей {caps}")
    else:
        add("PASS", "Подписи таблиц", f"{caps}")

    # подписи рисунков
    fig_caps = sum(1 for _, t, _ in paras if re.match(r"^\s*Рисунок\s+\d+", t, re.I))
    if info["drawings"] and fig_caps < info["drawings"]:
        add("WARN", "Подписи рисунков", f"рисунков {info['drawings']}, подписей {fig_caps}")
    else:
        add("PASS", "Подписи рисунков", f"{fig_caps}")

    # красный текст
    red = len(re.findall(r'w:color w:val="(?:FF0000|C00000)"', doc))
    add("WARN" if red else "PASS", "Красный текст", f"{red} run(s)")

    # штампы
    total_page = sum(h["page_fields"] for h in info["headers"])
    total_num = sum(h["numpages_fields"] for h in info["headers"])
    if cfg.get("require_stamp") or args.require_stamp:
        add(
            "PASS" if total_page and total_num else "WARN",
            "Поля штампа",
            f"PAGE={total_page}, NUMPAGES={total_num}",
        )

    # обязательные / запрещённые фразы
    for needle in require:
        add("PASS" if needle in text else "FAIL", f"Обязательно: {needle}", "")
    for needle in forbid:
        add("PASS" if needle not in text else "FAIL", f"Запрещено: {needle}", "")

    print("== Проверки ==")
    for status, name, detail in RESULTS:
        line = f"[{status}] {name}"
        if detail:
            line += f" — {detail}"
        print(line)

    fails = sum(1 for s, _, _ in RESULTS if s == "FAIL")
    warns = sum(1 for s, _, _ in RESULTS if s == "WARN")
    passes = len(RESULTS) - fails - warns
    print()
    print(f"Итог: PASS={passes}, WARN={warns}, FAIL={fails}")
    return 1 if fails else 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="QA-проверка .docx")
    ap.add_argument("path")
    ap.add_argument("--sources", type=int, help="ожидаемое число источников")
    ap.add_argument("--config", help="JSON с настройками")
    ap.add_argument("--require", action="append", help="обязательная фраза (можно несколько)")
    ap.add_argument("--forbid", action="append", help="запрещённая фраза (можно несколько)")
    ap.add_argument("--require-stamp", action="store_true", help="требовать PAGE+NUMPAGES")
    args = ap.parse_args(argv)
    if not Path(args.path).exists():
        print(f"Файл не найден: {args.path}", file=sys.stderr)
        return 2
    return run(args)


if __name__ == "__main__":
    raise SystemExit(main())
