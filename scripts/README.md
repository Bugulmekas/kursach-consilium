# Скрипты

## `consilium/` — работа с docx

| Файл | Назначение |
|---|---|
| `docx_tools.py` | `info`, `text`, `find`, `replace`, `citations` — без внешних зависимостей |
| `qa_docx.py` | формальная QA-проверка: заглушки, ссылки, дубли, подписи, штампы |
| `example.qa.json` | пример конфига для `qa_docx.py` |

### Примеры

```powershell
# факты о документе
python -X utf8 scripts\consilium\docx_tools.py info thesis.docx

# выгрузить текст со стилями
python -X utf8 scripts\consilium\docx_tools.py text thesis.docx --out dump.txt

# найти фразу
python -X utf8 scripts\consilium\docx_tools.py find thesis.docx "старая фраза"

# заменить (с резервной копией; оглавление не трогается)
python -X utf8 scripts\consilium\docx_tools.py replace thesis.docx "старая фраза" "новая фраза"

# проверить цитирование
python -X utf8 scripts\consilium\docx_tools.py citations thesis.docx --sources 20

# QA
python -X utf8 scripts\consilium\qa_docx.py thesis.docx --config scripts\consilium\example.qa.json
```

### Важно

- `replace` заменяет текст на уровне абзаца: внутри абзаца форматирование станет однородным. Для сложных абзацев правьте в Word.
- `replace` по умолчанию пропускает абзацы оглавления. Если нужно иначе — `--include-toc`.
- `qa_docx.py` не меняет файл. Код возврата 1 — есть FAIL.

## `word-com/` — COM-раннеры Word

| Файл | Назначение |
|---|---|
| `open_doc.ps1` | открыть docx, опционально обновить поля/TOC, задать зум |
| `goto_page.ps1` | перейти к странице в открытом документе |
| `desktop_shot.ps1` | скриншот всего экрана |
| `zoom.ps1` | изменить масштаб |
| `close_all.ps1` | закрыть все документы и Word |
| `update_fields.ps1` | обновить поля и TOC дважды, сохранить, закрыть |

```powershell
powershell -ExecutionPolicy Bypass -File scripts\word-com\open_doc.ps1 -Path "D:\work\thesis.docx" -Update -Zoom 70
powershell -ExecutionPolicy Bypass -File scripts\word-com\goto_page.ps1 -Page 2
powershell -ExecutionPolicy Bypass -File scripts\word-com\desktop_shot.ps1 -Out "D:\work\check.png"
powershell -ExecutionPolicy Bypass -File scripts\word-com\update_fields.ps1 -Path "D:\work\thesis.docx"
```

## Требования

- Windows PowerShell 5.1+ или PowerShell 7+.
- Microsoft Word.
- Python 3.10+.

Если Windows PowerShell 5.1 ломает кириллицу в `.ps1` — сохраните файл в UTF-8 **с BOM** или запускайте через `pwsh`.
