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
{ "api_key": "ВАШ_КЛЮЧ", "model": "GigaChat", "scope": "GIGACHAT_API_CORP" }
```
Файл монтируется в контейнер как volume — его правки переживают пересборку и рестарт.

**`scope` зависит от типа аккаунта, на который выпущен ключ:**
- `GIGACHAT_API_PERS` — физическое лицо;
- `GIGACHAT_API_B2B` — юр.лицо / ИП, предоплатные пакеты токенов;
- `GIGACHAT_API_CORP` — юр.лицо / ИП, оплата по факту (pay-as-you-go).

При смене аккаунта с физлица на организацию нужно и заменить `api_key` на ключ из
нового проекта организации, и выставить корректный `scope` (иначе OAuth вернёт HTTP 401).
Если поле `scope` не задано — используется значение по умолчанию `GIGACHAT_API_CORP`
(константа `DEFAULT_GIGACHAT_SCOPE` в `main.py`). Scope также можно выбрать в UI
(раздел «Настройки GigaChat API» → «Тип аккаунта»).

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

## 5.1. Если сборка падает на `pip install` (таймаут pypi.org)

Симптом:
```
ReadTimeoutError("HTTPSConnectionPool(host='pypi.org', port=443): Read timed out. (read timeout=15)")
```

Если шаг `apt-get install` при этом прошёл — сеть на сервере в целом работает,
проблема именно в доступе к pypi.org (медленный канал или блокировка).

**Шаг 1. Проверить доступность с сервера:**
```bash
curl -sS -o /dev/null -w '%{http_code} %{time_total}s\n' --max-time 30 https://pypi.org/simple/fastapi/
```
- `200` и время < 5 с — канал живой, помогает просто увеличенный таймаут (уже в Dockerfile: 120 с, 10 повторов). Повторите `docker compose build`.
- таймаут / другой код — переходите к шагу 2.

**Шаг 2. Собрать через зеркало.** Индекс вынесен в build-arg:
```bash
docker compose build --build-arg PIP_INDEX_URL=https://<зеркало>/simple backend
```
Зеркало сначала проверьте тем же `curl` с сервера — рабочее должно отдавать `200`
на `https://<зеркало>/simple/fastapi/`. Для зеркала без валидного TLS-сертификата
добавьте `--build-arg PIP_TRUSTED_HOST=<хост>`.

**Шаг 3. Если внешний доступ закрыт полностью** — собрать образ на машине с интернетом
и перенести:
```bash
docker compose build backend           # там, где интернет есть
docker save audit-backend | gzip > backend.tar.gz
scp backend.tar.gz user@SERVER_IP:/opt/audit-app/
# на сервере:
gunzip -c backend.tar.gz | docker load
docker compose up -d
```

Слой с `pip install` кэшируется: пока `requirements.txt` не меняется, повторные
сборки не ходят в сеть.

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
