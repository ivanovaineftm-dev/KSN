# Excel Processor (FastAPI)

Веб-приложение для загрузки, обработки и выгрузки Excel-файлов (`.xlsx`).

## Функционал

- Загрузка `.xlsx` файла через веб-форму.
- Сохранение исходного файла на сервере в папке `uploads/`.
- Обработка файла по правилам:
  1. Оставляются только строки, где столбец **C** содержит слово `стажер`.
  2. Удаляются строки, где столбец **G** пустой.
  3. Удаляются строки, где столбец **G** содержит `февраль` (любой регистр, также с подстрокой `феврал`).
  4. Удаляются строки, где столбец **H** пустой.
- Сохранение обработанного файла в папке `processed/`.
- Автоматическая выгрузка (скачивание) обработанного файла пользователю.

## Технологии

- Backend: **FastAPI**
- Frontend: **HTML + JavaScript**
- Обработка Excel: **openpyxl + pandas**

## Локальный запуск

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --reload
```

Откройте в браузере: `http://127.0.0.1:8000`

## Структура проекта

```text
app/
  main.py
  processor.py
  static/app.js
  templates/index.html
uploads/
processed/
requirements.txt
README.md
```
