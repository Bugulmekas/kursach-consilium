#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Небольшой набор инструментов для .docx без внешних зависимостей.

Работает с OOXML напрямую: читает word/document.xml, находит абзацы,
заменяет текст с сохранением разметки абзаца и пишет файл обратно.

Примеры:
    python -X utf8 docx_tools.py info thesis.docx
    python -X utf8 docx_tools.py text thesis.docx --out dump.txt
    python -X utf8 docx_tools.py find thesis.docx "старая фраза"
    python -X utf8 docx_tools.py replace thesis.docx "старая фраза" "новая фраза"
    python -X utf8 docx_tools.py citations thesis.docx

Важно:
- перед заменой создаётся резервная копия;
- замена идёт на уровне абзаца: если внутри абзаца были разные
  начертания, абзац станет однородным. Для точечных правок это
  обычно приемлемо, для сложной вёрстки — проверяйте результат;
- оглавление по умолчанию не трогается (--include-toc, чтобы включить).
"""
from __future__ import annotations

import argparse
import html
import os
import re
import shutil
import sys
import zipfile
from datetime import datetime
from pathlib import Path

PARA_RE = re.compile(r"<w:p\b[^>]*>.*?</w:p>", re.S)
T_RE = re.compile(r"<w:t(?: [^>]*)?>.*?</w:t>", re.S)
STYLE_RE = re.compile(r'<w:pStyle w:val="([^"]+)"')
CITE_RE = re.compile(r"\[([0-9,\u2013\-– ]+)\]")


# --------------------------------------------------------------------------
# чтение / запись
# --------------------------------------------------------------------------
def read_entries(path: str | Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as z:
        return {name: z.read(name) for name in z.namelist()}


def write_entries(path: str | Path, entries: dict[str, bytes]) -> None:
    path = Path(path)
    tmp = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in entries.items():
            z.writestr(name, data)
    os.replace(tmp, path)


def load_doc(path: str | Path) -> tuple[dict[str, bytes], str]:
    entries = read_entries(path)
    return entries, entries["word/document.xml"].decode("utf-8")


def backup(path: str | Path) -> Path:
    path = Path(path)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    dst = path.with_name(f"{path.stem}.bak-{stamp}{path.suffix}")
    shutil.copy2(path, dst)
    return dst


# --------------------------------------------------------------------------
# разбор XML
# --------------------------------------------------------------------------
def paragraph_text(p: str) -> str:
    raw = "".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", p))
    return html.unescape(raw)


def iter_paragraphs(doc: str):
    for m in PARA_RE.finditer(doc):
        yield m.start(), m.end(), m.group(0), paragraph_text(m.group(0))


def paragraph_style(p: str) -> str:
    m = STYLE_RE.search(p)
    return m.group(1) if m else ""


def is_toc(p: str) -> bool:
    if "<w:hyperlink" in p:
        return True
    style = paragraph_style(p).lower()
    return "toc" in style or "оглавлен" in style


def escape(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def set_paragraph_text(p: str, new_text: str) -> str:
    """Заменить текст абзаца, сохранив pPr и последний run.

    Текст кладётся в последний w:t, остальные w:t обнуляются. Так
    сохраняется шрифт/цвет последнего run и не ломается структура.
    """
    spans = list(T_RE.finditer(p))
    if not spans:
        return p
    out = p
    for i, m in enumerate(reversed(spans)):
        value = escape(new_text) if i == len(spans) - 1 else ""
        out = out[: m.start()] + f'<w:t xml:space="preserve">{value}</w:t>' + out[m.end() :]
    return out


# --------------------------------------------------------------------------
# команды
# --------------------------------------------------------------------------
def doc_info(path: str | Path) -> dict:
    entries, doc = load_doc(path)
    app = entries.get("docProps/app.xml", b"").decode("utf-8", errors="ignore")
    headers = []
    for name in entries:
        if re.match(r"word/(header|footer)\d+\.xml$", name):
            x = entries[name].decode("utf-8", errors="ignore")
            headers.append(
                {
                    "name": name,
                    "page_fields": x.count("PAGE"),
                    "numpages_fields": x.count("NUMPAGES"),
                    "tables": x.count("<w:tbl>"),
                }
            )
    instr = "".join(re.findall(r"<w:instrText[^>]*>([^<]*)<", doc))
    return {
        "path": str(path),
        "size": os.path.getsize(path),
        "pages": (re.findall(r"<Pages>(\d+)</Pages>", app) or ["?"])[0],
        "words": (re.findall(r"<Words>(\d+)</Words>", app) or ["?"])[0],
        "sections": doc.count("<w:sectPr"),
        "tables": doc.count("<w:tbl>"),
        "drawings": doc.count("<w:drawing>"),
        "title_page": doc.count("<w:titlePg"),
        "toc_fields": instr.count("TOC"),
        "headers": headers,
    }


def cmd_info(args) -> int:
    info = doc_info(args.path)
    print(f"Файл: {info['path']}")
    print(f"Размер: {info['size']} байт")
    print(f"Страниц (по app.xml): {info['pages']}")
    print(f"Слов: {info['words']}")
    print(f"Секций: {info['sections']}")
    print(f"Таблиц: {info['tables']}")
    print(f"Рисунков (w:drawing): {info['drawings']}")
    print(f"titlePg (особый титул): {info['title_page']}")
    print(f"Полей TOC: {info['toc_fields']}")
    print("Колонтитулы:")
    for h in info["headers"]:
        print(
            f"  {h['name']}: PAGE={h['page_fields']} "
            f"NUMPAGES={h['numpages_fields']} таблиц={h['tables']}"
        )
    return 0


def cmd_text(args) -> int:
    entries, doc = load_doc(args.path)
    styles = {}
    if "word/styles.xml" in entries:
        sx = entries["word/styles.xml"].decode("utf-8", errors="ignore")
        for m in re.finditer(r'<w:style [^>]*w:styleId="([^"]+)"[^>]*>.*?</w:style>', sx, re.S):
            nm = re.search(r'<w:name w:val="([^"]+)"/>', m.group(0))
            if nm:
                styles[m.group(1)] = nm.group(1)

    lines = []
    for _, _, p, text in iter_paragraphs(doc):
        if not text.strip():
            continue
        sid = paragraph_style(p)
        name = styles.get(sid, sid)
        lines.append(f"[{name}] {text}" if name else text)

    out = "\n".join(lines)
    if args.out:
        Path(args.out).write_text(out, encoding="utf-8")
        print(f"Записано: {args.out} ({len(lines)} абзацев)")
    else:
        print(out)
    return 0


def cmd_find(args) -> int:
    _, doc = load_doc(args.path)
    found = 0
    for i, (_, _, p, text) in enumerate(iter_paragraphs(doc), 1):
        if args.text.lower() in text.lower():
            found += 1
            print(f"#{i} [{paragraph_style(p) or '-'}] {text[:300]}")
    print(f"Найдено: {found}")
    return 0 if found else 1


def cmd_replace(args) -> int:
    path = Path(args.path)
    if not args.text:
        print("Укажите текст для поиска.", file=sys.stderr)
        return 2
    entries, doc = load_doc(path)
    matches = []
    for start, end, p, text in iter_paragraphs(doc):
        if args.text and args.skip_toc and is_toc(p):
            continue
        if args.text and args.text.lower() not in text.lower():
            continue
        matches.append((start, end, p, text))
        if not args.all:
            break

    if not matches:
        print("Совпадений нет.")
        return 1

    if args.dry_run:
        for _, _, _, text in matches:
            print(f"БЫЛО: {text[:300]}")
        print(f"Совпадений: {len(matches)} (dry-run, файл не изменён)")
        return 0

    for start, end, p, text in reversed(matches):
        new_p = set_paragraph_text(p, args.new)
        doc = doc[:start] + new_p + doc[end:]

    if not args.no_backup:
        print(f"Резервная копия: {backup(path)}")

    entries["word/document.xml"] = doc.encode("utf-8")
    write_entries(path, entries)
    print(f"Заменено абзацев: {len(matches)}")
    for _, _, _, text in matches:
        print(f"  было: {text[:200]}")
    return 0


def cmd_citations(args) -> int:
    _, doc = load_doc(args.path)
    text = "\n".join(t for _, _, _, t in iter_paragraphs(doc))
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
        print("Ссылок вида [n] не найдено.")
        return 1
    print("Использованные номера:", ", ".join(map(str, sorted(nums))))
    print("Максимум:", max(nums))
    if args.sources:
        missing = sorted(set(range(1, args.sources + 1)) - nums)
        extra = sorted(nums - set(range(1, args.sources + 1)))
        print("Не цитируются:", missing or "нет")
        print("Вне диапазона:", extra or "нет")
        return 0 if not missing and not extra else 1
    return 0


def build_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(description="Инструменты для .docx без внешних зависимостей")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("info", help="факты о документе")
    p.add_argument("path")
    p.set_defaults(func=cmd_info)

    p = sub.add_parser("text", help="выгрузить текст с стилями")
    p.add_argument("path")
    p.add_argument("--out")
    p.set_defaults(func=cmd_text)

    p = sub.add_parser("find", help="найти текст")
    p.add_argument("path")
    p.add_argument("text")
    p.set_defaults(func=cmd_find)

    p = sub.add_parser("replace", help="заменить текст в абзаце")
    p.add_argument("path")
    p.add_argument("text", help="что искать")
    p.add_argument("new", help="на что заменить")
    p.add_argument("--all", action="store_true", help="все совпадения")
    p.add_argument("--include-toc", action="store_true", help="не пропускать оглавление")
    p.add_argument("--no-backup", action="store_true")
    p.add_argument("--dry-run", action="store_true")
    p.set_defaults(func=cmd_replace, skip_toc=True)

    p = sub.add_parser("citations", help="проверить ссылки [n]")
    p.add_argument("path")
    p.add_argument("--sources", type=int, help="сколько источников должно быть")
    p.set_defaults(func=cmd_citations)

    return ap


def main(argv=None) -> int:
    ap = build_parser()
    args = ap.parse_args(argv)
    if getattr(args, "include_toc", False):
        args.skip_toc = False
    try:
        return args.func(args)
    except KeyError as e:
        print(f"Ошибка: в документе нет части {e}", file=sys.stderr)
        return 2
    except FileNotFoundError as e:
        print(f"Ошибка: файл не найден: {e}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
