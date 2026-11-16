# Рецепты Word: MCP, COM, OOXML

## 1. MCP word-live

### Найти и заменить

```text
word_live_find_text({ search_text: "старая фраза", context_chars: 80 })
word_live_replace_text({ find_text: "старая фраза", replace_text: "новая фраза" })
```

Ограничения:

- find/replace — до ~255 символов;
- если фраза разбита на несколько run'ов, ищите короткий устойчивый фрагмент;
- после замены проверяйте `word_screen_capture`.

### Вставить текст

```text
word_live_insert_text({ text: "Текст", position: "end" })
```

`position` может быть `start`, `end` или имя закладки. Для вставки **в конкретный абзац** надёжнее python-docx при закрытом Word.

### Форматирование

```text
word_live_format_text({
  start_paragraph: 12,
  end_paragraph: 12,
  bold: true,
  color: "000000"
})
```

### Скриншот

```text
word_screen_capture({ output_path: "C:\\temp\\check.png" })
```

Затем прочитать PNG. Это единственная надёжная проверка «как выглядит на экране».

### Сохранить

```text
word_live_save({})
```

---

## 2. PowerShell COM

### Открыть и обновить поля

```powershell
$word = New-Object -ComObject Word.Application
$word.Visible = $true
$doc = $word.Documents.Open($path)
$doc.Fields.Update()
$doc.Repaginate()
$doc.TablesOfContents.Item(1).Update()
$doc.Repaginate()
$doc.TablesOfContents.Item(1).Update()
$doc.Fields.Update()
$doc.Save()
```

### Страницы и секции

```powershell
$doc.Sections.Count
$doc.ComputeStatistics(2)      # страниц
$p.Range.Information(3)        # страница абзаца
```

### Перебор коллекций

```powershell
for ($i = 1; $i -le $doc.Sections.Count; $i++) {
    $s = $doc.Sections.Item($i)
}
```

`foreach` по COM-коллекциям может молча не выполниться.

### Закрыть всё

```powershell
for ($i = $word.Documents.Count; $i -ge 1; $i--) {
    $word.Documents.Item($i).Close(0)
}
$word.Quit()
```

---

## 3. OOXML

### Прочитать текст

```python
import html, re, zipfile

z = zipfile.ZipFile(path)
doc = z.read('word/document.xml').decode('utf-8')
paras = re.findall(r'<w:p\b[^>]*>.*?</w:p>', doc, re.S)

def text(p):
    return html.unescape(
        ''.join(re.findall(r'<w:t(?: [^>]*)?>([^<]*)</w:t>', p))
    )
```

### Заменить текст в абзаце, сохранив форматирование

```python
def set_para_text(p, new_text):
    spans = list(re.finditer(r'<w:t(?: [^>]*)?>.*?</w:t>', p, re.S))
    out = p
    for i, m in enumerate(reversed(spans)):
        val = esc(new_text) if i == len(spans) - 1 else ''
        out = out[:m.start()] + f'<w:t xml:space="preserve">{val}</w:t>' + out[m.end():]
    return out
```

### Записать обратно

```python
with zipfile.ZipFile(path) as z:
    entries = {n: z.read(n) for n in z.namelist()}

entries['word/document.xml'] = new_doc.encode('utf-8')

with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
    for name, data in entries.items():
        z.writestr(name, data)
```

### Вставить абзац по шаблону

Берём `w:pPr` и `w:rPr` соседнего абзаца:

```python
ppr = re.search(r'<w:pPr>.*?</w:pPr>', neighbor, re.S).group(0)
rpr = re.search(r'<w:rPr>.*?</w:rPr>', run, re.S).group(0)

new_p = (
    '<w:p>' + ppr +
    '<w:r>' + rpr +
    '<w:t xml:space="preserve">' + esc(text) + '</w:t>' +
    '</w:r></w:p>'
)
```

### Таблица

Ячейка:

```python
'<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>' + borders + '</w:tcPr>'
'<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>'
'<w:r><w:rPr>' + rpr + '</w:rPr>'
'<w:t xml:space="preserve">' + esc(text) + '</w:t></w:r></w:p></w:tc>'
```

---

## 4. Типовые проверки

### Страницы, секции, таблицы, рисунки

```python
app = z.read('docProps/app.xml').decode('utf-8')
pages = re.findall(r'<Pages>(\d+)</Pages>', app)
sections = doc.count('<w:sectPr')
tables = doc.count('<w:tbl>')
figures = doc.count('<w:drawing>')
```

### Поля в колонтитулах

```python
for name in z.namelist():
    if re.match(r'word/(header|footer)\d+\.xml$', name):
        x = z.read(name).decode('utf-8')
        print(name, x.count('PAGE'), x.count('NUMPAGES'))
```

### Цитирование

```python
cites = re.findall(r'\[([0-9,\u2013\-– ]+)\]', text)
nums = set()
for c in cites:
    for part in c.split(','):
        part = part.strip().replace('–', '-')
        if '-' in part:
            a, b = part.split('-')
            if a.strip().isdigit() and b.strip().isdigit():
                nums.update(range(int(a), int(b) + 1))
        elif part.isdigit():
            nums.add(int(part))
```

---

## 5. Кодировки

- Python: `sys.stdout.reconfigure(encoding='utf-8')`, запуск `python -X utf8`.
- PowerShell: файлы с кириллицей — UTF-8 **с BOM**.
- Не передавать русские строки в inline-команды PowerShell: только через файлы.
