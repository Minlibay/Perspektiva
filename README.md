# План АУДИТА - Автоматическое заполнение

Web-приложение для автоматического заполнения документа "План АУДИТА" с использованием GigaChat API.

## Архитектура

- **Backend**: Python + FastAPI (порт 8000)
- **Frontend**: React + Bootstrap (порт 3000 в режиме dev, или статическая сборка в `frontend/dist`)

## Возможности

✅ Загрузка документов из папок Пакет 2, Пакет 3 и других  
✅ Автоматическое извлечение данных через GigaChat API  
✅ Заполнение шаблона "План АУДИТА" извлечёнными данными  
✅ Предпросмотр извлечённых данных  
✅ Скачивание заполненного документа  

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

## Использование

1. **Введите API ключ GigaChat** в настройках и сохраните
2. **Загрузите документы**:
   - Нажмите "Загрузить из папки" и укажите путь (например: `D:\Perpektiva\Пакет 2`)
   - Или выберите файлы вручную
3. **Выберите шаблон** плана из списка загруженных файлов
4. **Запустите обработку** - GigaChat извлечёт данные из документов
5. **Просмотрите результат** и скачайте заполненный план

## Структура проекта

```
D:\Perpektiva\Web\
├── backend\
│   ├── main.py                 # FastAPI приложение
│   ├── requirements.txt        # Python зависимости
│   ├── uploads\                # Загруженные документы
│   └── outputs\                # Готовые заполненные планы
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
| POST | `/api/upload` | Загрузить файлы |
| POST | `/api/upload-from-path` | Загрузить из папки |
| GET | `/api/files` | Список загруженных файлов |
| DELETE | `/api/files/{filename}` | Удалить файл |
| POST | `/api/process` | Обработать документы через GigaChat |
| GET | `/api/download/{filename}` | Скачать результат |

## Получение API ключа GigaChat

1. Зарегистрируйтесь на [https://developers.sber.ru/gigachat](https://developers.sber.ru/gigachat)
2. Создайте проект и получите API ключ
3. Вставьте ключ в настройках приложения
