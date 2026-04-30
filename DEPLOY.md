# Развёртывание (Ubuntu 24.04, Docker, HTTP по IP)

Целевой сервер: 2 CPU / 6 GB RAM, публичный IP. Доступ только по HTTP (порт 80).

## 1. Подготовка сервера

```bash
sudo apt update && sudo apt upgrade -y
sudo apt install -y docker.io docker-compose-plugin git
sudo systemctl enable --now docker
sudo usermod -aG docker $USER   # затем перелогиниться
```

Открыть в фаерволе только 80:
```bash
sudo ufw allow OpenSSH
sudo ufw allow 80/tcp
sudo ufw enable
```

## 2. Залить код

Скопировать каталог `Web/` на сервер целиком (например, `/opt/audit-app/`):
```bash
scp -r ./Web user@SERVER_IP:/opt/audit-app
```
или через `rsync -av --exclude node_modules --exclude __pycache__ --exclude uploads --exclude outputs ./Web/ user@SERVER_IP:/opt/audit-app/`.

## 3. Ключ GigaChat

На сервере положить корректный `backend/gigachat_settings.json` вида:
```json
{ "api_key": "ВАШ_КЛЮЧ", "model": "GigaChat" }
```
Файл монтируется в контейнер как volume — его правки переживают пересборку и рестарт.

## 4. Сборка и запуск

```bash
cd /opt/audit-app
docker compose build
docker compose up -d
docker compose logs -f --tail=100
```

Приложение доступно на `http://SERVER_IP/`. Все API-запросы фронт делает same-origin через nginx (`/api/...`).

## 5. Обновление

```bash
cd /opt/audit-app
git pull            # либо rsync новых файлов
docker compose build
docker compose up -d
```

## 6. Останов / диагностика

```bash
docker compose down              # остановить
docker compose logs backend      # логи бэка
docker compose exec backend bash # шелл внутри бэка
docker stats                     # потребление памяти/CPU
```

## Что отключено/изменено для прод-деплоя

- `/api/upload-from-path` отключён по умолчанию (env `ENABLE_LOCAL_UPLOAD=0`). На сервере произвольные локальные пути не имеют смысла, и эндпоинт открыт на path-traversal. Включать только осознанно.
- Фронт обращается на относительный путь (`/api/...`), а не `localhost:8000`. CORS поэтому не задействуется.
- `uploads/` и `outputs/` живут только внутри контейнера и сбрасываются при `docker compose down`. Если потребуется сохранять — добавить volumes в `docker-compose.yml`.
- `verify_ssl_certs=False` для GigaChat оставлено как было (требование интеграции).

## Замечания по ресурсам (2 CPU / 6 GB)

- `easyocr` тянет PyTorch CPU — образ бэка ~2–3 GB, первый запуск OCR скачивает модели (~100 MB) в `~/.EasyOCR` внутри контейнера.
- При обработке крупных PDF возможен пик памяти. Лимит бэка в compose: 4 GB. При OOM — уменьшить размер пакета или добавить swap (`sudo fallocate -l 4G /swapfile && sudo mkswap /swapfile && sudo swapon /swapfile`).
- Запущен один воркер uvicorn — параллельная обработка не предусмотрена (см. фиксированное имя выходного файла в `main.py`).
