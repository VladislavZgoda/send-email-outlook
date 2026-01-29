Запуск:
1. .\.venv\Scripts\activate
2. uv run main.py

Для работы необходимо создать в папке со скриптом файл emails.json
[
  {
    "email": "test@yandex.ru",
    "files": ["test.xlsx"]
  }
]

Затем создать папку assets и положить в нее необходимые файлы для отправки.
!!! Outlook должен быть установлен и запущен.
