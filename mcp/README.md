# MCP-конфигурация

Файл [`opencode.mcp.json`](opencode.mcp.json) — готовый фрагмент для `opencode.json`.

## Куда вставлять

Откройте `%USERPROFILE%\.config\opencode\opencode.json` и добавьте содержимое `mcp` в корневой объект. Если ключ `mcp` уже есть — слейте объекты, не перезаписывайте существующие серверы.

## Что заменить

| Плейсхолдер | На что |
|---|---|
| `C:\\Users\\<user>\\.local\\bin\\browse.exe` | реальный путь к `browse.exe` |
| `C:\\Users\\<user>\\.local\\bin\\uvx.exe` | реальный путь к `uvx.exe` (или просто `uvx`, если он в `PATH`) |

## Вариант с обёрткой

Если ваш OpenCode (или плагин) ожидает `mcp.servers`, оберните так:

```json
{
  "mcp": {
    "servers": {
      "browse": { "type": "local", "command": ["...browse.exe", "mcp"] },
      "word-live": {
        "type": "local",
        "command": ["...uvx.exe", "--from", "word-mcp-live", "word_mcp_server.exe"]
      }
    }
  }
}
```

## Проверка

1. Перезапустить OpenCode.
2. Открыть Word с тестовым `.docx`.
3. Вызвать `word_live_get_info` — должны прийти страницы и слова.
4. Вызвать `word_screen_capture` — должен появиться PNG.

## Требования

- Windows + установленный Microsoft Word.
- `uv`/`uvx`.
- Для `browse` — браузерное расширение OpenCode.

Подробности: [`docs/02-MCP-SERVER.md`](../docs/02-MCP-SERVER.md).
