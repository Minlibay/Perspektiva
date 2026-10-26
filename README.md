# Проверка документов аудита

Web-приложение для проверки документов аудита ОС СМК по чек-листам: План АУДИТА, Сводный акт,
Акт, Заключения, Отчеты, ПКД, ПН, Листы оценки — всего 11 чек-листов. План и Сводный акт
проверяются через GigaChat API, остальные — по формулам, без модели.

## Архитектура

- **Backend**: Python + FastAPI (порт 8000)
- **Frontend**: React + Bootstrap (порт 3000 в режиме dev, или статическая сборка в `frontend/dist`)

## Возможности

✅ Загрузка пакета аудита в три поля: бланки чек-листов, проверяемые документы, источники  
✅ Автоматическое определение типа бланка и документа, ручное исправление типа  
✅ Предупреждения: документы разных организаций, чек-лист без документа, дубликаты  
✅ Заполнение бланков отметками ОК/NOK с обоснованием, пункты «проверить вручную»  
✅ Скачивание каждого заполненного чек-листа или всех ZIP-архивом  

## Установка и запуск

### 1. Backend

```bash
cd backend
pip install -r requirements.txt
python main.py
```

Backend запустится на `http://localhost:8000`

### 2. Frontend (режим разработки)

```bash
cd frontend
npm install
npm run dev
```

Frontend запустится на `http://localhost:3000`

### 3. Frontend (production сборка)

```bash
cd frontend
npm run build
```

Собранные файлы находятся в `frontend/dist/`

Для запуска только backend с раздачей статики:

```python
# В main.py добавьте раздачу статики
from fastapi.staticfiles import StaticFiles
app.mount("/", StaticFiles(directory="../frontend/dist", html=True), name="static")
```

### 4. Тесты backend

```bash
cd backend
pip install -r requirements-dev.txt
python -m pytest tests
```

Чек-листы, их правила и план подключения — [docs/ЧЕК-ЛИСТЫ.md](docs/ЧЕК-ЛИСТЫ.md).

## Использование

1. **Загрузите пакет аудита в три поля**: бланки чек-листов (поле 1), проверяемые документы
   (поле 2), источники (поле 3). Тип каждого бланка и документа определяется автоматически;
   неверный тип документа можно исправить в списке под файлом.
2. **Проверьте пакет**: система показывает предупреждения (нераспознанные файлы, чек-лист без
   документа, документы разных организаций, дубликаты) и таблицу «Что проверить».
3. **Отметьте чек-листы и запустите проверку.** План АУДИТА и Сводный акт проверяются через
   GigaChat (нужен API-ключ в «Админ»), остальные 9 — без модели, за секунды.
4. **Результаты появляются по мере готовности**; каждый заполненный бланк скачивается отдельно
   или все вместе ZIP-архивом. Подробности — [docs/ЧЕК-ЛИСТЫ.md](docs/ЧЕК-ЛИСТЫ.md).

## Структура проекта

```
D:\Perpektiva\Web\
├── backend\
│   ├── main.py                 # FastAPI приложение, План и Сводный акт (GigaChat)
│   ├── checklists.py           # Реестр и разбор бланков чек-листов
│   ├── doc_types.py            # Определение типа документа
│   ├── fields.py               # Извлечение полей из документов
│   ├── formula_checks.py       # Правила чек-листов «по формулам»
│   ├── package.py              # Разбор пакета аудита и предупреждения
│   ├── tests\                  # Тесты (pytest)
│   ├── requirements.txt        # Python зависимости
│   ├── uploads\                # Загруженные документы по сессиям
│   └── outputs\                # Заполненные чек-листы по сессиям
├── frontend\
│   ├── src\
│   │   ├── App.jsx             # Основной компонент React
│   │   ├── App.css             # Стили
│   │   └── main.jsx            # Точка входа React
│   ├── index.html
│   ├── package.json
│   ├── vite.config.js
│   └── dist/                   # Production сборка
└── README.md
```

## API Endpoints

| Метод | Endpoint | Описание |
|-------|----------|----------|
| GET | `/` | Статус API |
| POST | `/api/settings/gigachat` | Сохранить API ключ GigaChat |
| GET | `/api/settings/gigachat` | Получить API ключ |
| POST | `/api/session/new` | Новая сессия (пакет аудита) |
| POST | `/api/upload` | Загрузить файлы в поле (`block`: `checklist` / `plan` / `sources`) |
| POST | `/api/upload-from-path` | Загрузить из папки на сервере (только при `ENABLE_LOCAL_UPLOAD=1`) |
| GET | `/api/files` | Список загруженных файлов |
| DELETE | `/api/files/{filename}` | Удалить файл из сессии и из поля |
| POST | `/api/sessions/{id}/analyze` | Разбор пакета: типы файлов, доступные чек-листы, предупреждения |
| POST | `/api/sessions/{id}/run` | Запустить проверку выбранных чек-листов (в фоне) |
| GET | `/api/sessions/{id}/status` | Ход проверки и готовые результаты |
| GET | `/api/download/{session_id}/{file}` | Скачать заполненный чек-лист |
| GET | `/api/download-outputs/{session_id}` | Все заполненные чек-листы ZIP-архивом |
| GET | `/api/gigachat/diagnose` | Проверка ключа GigaChat и список моделей |
| GET | `/api/admin/sessions` | Список сессий (админка) |
| GET | `/api/admin/sessions/{id}` | Файлы и результаты сессии |
| GET | `/api/admin/sessions/{id}/zip` | Сессия ZIP-архивом |
| DELETE | `/api/admin/sessions/{id}` | Удалить сессию |

## Получение API ключа GigaChat

1. Зарегистрируйтесь на [https://developers.sber.ru/gigachat](https://developers.sber.ru/gigachat)
2. Создайте проект и получите API ключ
3. Вставьте ключ в настройках приложения
