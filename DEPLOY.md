# Развёртывание на VPS (Ubuntu 24.04, Docker)

Приложение — два контейнера:
- **backend** — FastAPI, разбор документов, OCR (Tesseract), обращения к GigaChat;
- **frontend** — nginx: отдаёт интерфейс и проксирует `/api/...` на backend.

Сайт открывается по адресу `http://IP_СЕРВЕРА:8080/`. Порт можно сменить, а сайт — закрыть паролем
или перевести на домен с HTTPS (разделы 4–6).

**Сервер:** от 2 CPU и 4 GB RAM (лучше 6 GB), от 20 GB диска (образы ~3 GB, остальное — загруженные
пакеты), Ubuntu 22.04 или 24.04 и доступ по SSH.

Все команды ниже выполняются на сервере, если не сказано иное.

## 0. Код должен быть в GitHub

Сервер забирает код из репозитория `github.com/Minlibay/Perspektiva`. Перед первым развёртыванием и
перед каждым обновлением закоммитьте и запушьте изменения со своей машины. Если они ещё не слиты в
`main`, на сервере понадобится имя ветки.

## 1. Подготовка сервера

```bash
sudo apt update && sudo apt upgrade -y
sudo apt install -y docker.io docker-compose-v2 git
sudo systemctl enable --now docker
sudo usermod -aG docker $USER
```

Перелогиньтесь (выйдите из SSH и зайдите снова), затем проверьте:
```bash
docker compose version
```

Если пакета `docker-compose-v2` нет (старый Ubuntu), поставьте Docker официальным скриптом:
`curl -fsSL https://get.docker.com | sudo sh`.

**Swap** — страховка от нехватки памяти при OCR больших сканов:
```bash
sudo fallocate -l 4G /swapfile
sudo chmod 600 /swapfile
sudo mkswap /swapfile
sudo swapon /swapfile
echo '/swapfile none swap sw 0 0' | sudo tee -a /etc/fstab
```

**Фаервол:**
```bash
sudo ufw allow OpenSSH
sudo ufw allow 80/tcp
sudo ufw allow 443/tcp
sudo ufw allow 8080/tcp
sudo ufw enable
```
Docker открывает порты контейнеров в обход ufw: даже без правила `8080` будет доступен из интернета.
Поэтому ограничивать доступ к сайту нужно паролем или списком IP в nginx (раздел 5), а не фаерволом.

## 2. Код на сервер

Если репозиторий публичный:
```bash
sudo mkdir -p /opt/audit-app && sudo chown $USER: /opt/audit-app
git clone https://github.com/Minlibay/Perspektiva.git /opt/audit-app
```

Если приватный — заведите для сервера ключ только на чтение (deploy key):
```bash
ssh-keygen -t ed25519 -C "audit-vps" -f ~/.ssh/id_ed25519 -N ""
cat ~/.ssh/id_ed25519.pub
```
Скопируйте выведенную строку в GitHub: репозиторий → Settings → Deploy keys → Add deploy key
(галочку «Allow write access» не ставить). Затем:
```bash
sudo mkdir -p /opt/audit-app && sudo chown $USER: /opt/audit-app
git clone git@github.com:Minlibay/Perspektiva.git /opt/audit-app
```

Если нужные изменения ещё не в `main`:
```bash
cd /opt/audit-app
git checkout ИМЯ_ВЕТКИ
```

## 3. Настройки и каталоги данных

```bash
cd /opt/audit-app
[ -f backend/gigachat_settings.json ] || echo '{}' > backend/gigachat_settings.json
mkdir -p backend/uploads backend/outputs
```

**Файл `gigachat_settings.json` обязательно создать до первого запуска.** Иначе Docker создаст на
его месте каталог, и ключ GigaChat не сохранится. Если так уже случилось:
`docker compose down && sudo rm -r backend/gigachat_settings.json`, затем команды выше.

Ключ GigaChat удобнее ввести после запуска в интерфейсе: «🔧 Админ» → «Настройки GigaChat API»
(ключ, тип аккаунта, модель) → «Сохранить» → «🔍 Диагностика». Можно и сразу в файле:
```json
{ "api_key": "ВАШ_КЛЮЧ", "model": "GigaChat", "scope": "GIGACHAT_API_CORP" }
```

**`scope` зависит от типа аккаунта, на который выпущен ключ:**
- `GIGACHAT_API_PERS` — физическое лицо;
- `GIGACHAT_API_B2B` — юр.лицо / ИП, предоплатные пакеты токенов;
- `GIGACHAT_API_CORP` — юр.лицо / ИП, оплата по факту (pay-as-you-go).

При смене аккаунта с физлица на организацию нужно и заменить `api_key` на ключ из нового проекта
организации, и выставить корректный `scope` (иначе OAuth вернёт HTTP 401). Без `scope`
используется `GIGACHAT_API_CORP` (константа `DEFAULT_GIGACHAT_SCOPE` в `main.py`).

GigaChat нужен только для чек-листов «План АУДИТА» и «Сводный акт». Остальные 9 работают без ключа.

## 4. Сборка и запуск

```bash
cd /opt/audit-app
docker compose build
docker compose up -d
docker compose ps
```

Первая сборка занимает 5–15 минут. Python-зависимости ставятся из `backend/vendor/` без обращения к
pypi.org — в логе сборки видно `>>> Офлайн-установка из vendor/`. Если сборка падает на скачивании
образов или пакетов — разделы 9–10.

В `docker compose ps` оба контейнера должны быть в состоянии `Up`. Проверка:
1. Откройте `http://IP_СЕРВЕРА:8080/` — должна открыться «Проверка документов аудита».
2. «🔧 Админ» → введите ключ GigaChat → «Сохранить» → «🔍 Диагностика»: должен прийти список моделей.
3. «← К рабочему режиму» → загрузите бланк чек-листа в поле 1: у файла появится тип «ЧК: …».

Логи backend: `docker compose logs -f --tail=100 backend`.

**Сайт на порту 80 вместо 8080** (порт должен быть свободен — проверка: `sudo ss -ltnp | grep ':80 '`):
```bash
cd /opt/audit-app
echo 'APP_PORT=80' > .env
docker compose up -d
```
Файл `.env` не хранится в git, поэтому `git pull` его не затрёт.

## 5. Закрыть доступ паролем (важно)

Входа по логину в приложении нет. Любой, кто знает адрес, может открыть «Админ», увидеть все сессии
с документами клиентов, скачать их и прочитать ключ GigaChat. На публичном сервере закройте сайт
паролем — это делается в nginx без правки файлов из git:

```bash
cd /opt/audit-app
# пароль для входа: логин auditor, пароль команда спросит дважды
printf 'auditor:%s\n' "$(openssl passwd -apr1)" > .htpasswd
# копия frontend/nginx.conf с двумя строками auth_basic
sed 's|^    server_name _;|&\n    auth_basic "Audit";\n    auth_basic_user_file /etc/nginx/.htpasswd;|' \
    frontend/nginx.conf > nginx-auth.conf
grep -c auth_basic nginx-auth.conf   # должно вывести 2
cat > docker-compose.override.yml <<'EOF'
services:
  frontend:
    volumes:
      - ./nginx-auth.conf:/etc/nginx/conf.d/default.conf:ro
      - ./.htpasswd:/etc/nginx/.htpasswd:ro
EOF
docker compose up -d
```

`docker-compose.override.yml` Docker подхватывает сам, все три файла в git не попадают. Ещё
пользователь: `printf 'имя:%s\n' "$(openssl passwd -apr1)" >> .htpasswd`, затем
`docker compose restart frontend`. Если поменяется `frontend/nginx.conf` в репозитории, пересоздайте
`nginx-auth.conf` той же командой `sed`.

Вместо пароля можно пускать только с адресов офиса: в `nginx-auth.conf` замените две строки
`auth_basic…` на `allow 203.0.113.10;` (по строке на адрес) и `deny all;`.

По HTTP пароль передаётся почти открытым текстом. Если есть домен — переведите сайт на HTTPS (раздел 6).

## 6. Домен и HTTPS (по желанию)

Нужна DNS-запись типа A: `audit.example.ru` → IP сервера. Сертификат Let's Encrypt выпустит и будет
продлевать Caddy, установленный прямо на сервере:

```bash
cd /opt/audit-app
echo 'APP_PORT=127.0.0.1:8080' > .env   # контейнер виден только Caddy, не из интернета
docker compose up -d
sudo apt install -y caddy
sudo tee /etc/caddy/Caddyfile <<'EOF'
audit.example.ru {
    reverse_proxy 127.0.0.1:8080
}
EOF
sudo systemctl reload caddy
```

Сайт откроется по `https://audit.example.ru/`. Порты 80 и 443 должны быть свободны: их займёт Caddy.
Пароль из раздела 5 продолжает работать. Если пакета `caddy` нет в репозитории — установка по
инструкции https://caddyserver.com/docs/install.

## 7. Обновление

Со своей машины запушьте изменения в GitHub, затем на сервере:
```bash
cd /opt/audit-app
git pull
docker compose build
docker compose up -d
docker image prune -f      # удалить старые образы
```

Обновляйте, когда никто не запустил проверку: рестарт прерывает идущую проверку. Загруженные пакеты и
готовые результаты лежат на диске сервера (раздел 8) и переживают пересборку.

## 8. Данные, место на диске, резервная копия

- Пакеты и результаты — в `backend/uploads/<сессия>/` и `backend/outputs/<сессия>/` на сервере
  (смонтированы в контейнер). Сами они не удаляются: старые сессии удаляйте в «Админ» → 🗑.
- Сколько занято: `du -sh backend/uploads backend/outputs`, `df -h`, `docker system df`.
  Кэш сборки Docker чистится командой `docker builder prune -f`.
- Резервная копия:
  ```bash
  cd /opt/audit-app
  tar czf ~/audit-backup-$(date +%F).tgz backend/uploads backend/outputs backend/gigachat_settings.json
  ```

## 9. Диагностика

```bash
docker compose ps                         # оба контейнера должны быть Up
docker compose logs --tail=200 backend    # ошибки бэка
docker compose logs --tail=200 frontend   # ошибки nginx
docker compose exec backend bash          # шелл внутри бэка
docker stats                              # память/CPU
docker compose down                       # остановить всё
```

- **Сайт не открывается.** На сервере выполните `curl -I http://localhost:8080/` (или порт из `.env`). Если в ответе `200`
  или `401`, приложение работает, а порт закрыт снаружи. Проверьте фаервол в панели хостера: у многих
  VPS он свой, отдельно от ufw.
- **`502 Bad Gateway` при загрузке или проверке.** Backend не запущен или упал — смотрите
  `docker compose logs backend`.
- **План или Сводный акт завершаются с ошибкой GigaChat.** «Админ» → «🔍 Диагностика». HTTP 401 —
  неверный ключ или `scope` (раздел 3). Остальные чек-листы проверяются и без GigaChat.
- **Не скачиваются образы** (`python:3.11-slim`, `node:20-alpine`, `nginx:1.27-alpine`) — Docker Hub
  недоступен с сервера. Подключите зеркало и пересоберите:
  ```bash
  echo '{ "registry-mirrors": ["https://mirror.gcr.io"] }' | sudo tee /etc/docker/daemon.json
  sudo systemctl restart docker
  ```

## 10. Сборка падает на `pip install` (таймаут pypi.org)

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

**Шаг 3. Офлайн-сборка из `backend/vendor/` (основной вариант, когда pypi недоступен).**

В репозитории лежит каталог `backend/vendor/` с заранее скачанными `.whl` под Linux/py3.11.
Если он непустой, Dockerfile ставит зависимости **только из него**, в сеть не обращаясь:

```bash
git pull
docker compose build backend
docker compose up -d
```

В логе сборки при этом видно `>>> Офлайн-установка из vendor/`.

Как обновить `vendor/` (делается на машине с интернетом, при смене `requirements.txt`):

```bash
cd backend
docker run --rm -v "$PWD:/w" -w /w python:3.11-slim pip download -r requirements.txt -d vendor
```

Скачивание внутри контейнера гарантирует правильные Linux-сборки. Если Docker
недоступен, тот же результат обычным pip — с явным указанием платформы:

```bash
pip download -r requirements.txt -d vendor --only-binary=:all: \
  --platform manylinux2014_x86_64 --platform manylinux_2_17_x86_64 \
  --platform manylinux_2_28_x86_64 --platform any --python-version 3.11
```

**Важно:** нужны именно `.whl`. Файлы `.tar.gz` (исходники) офлайн не ставятся —
для сборки они тянут build-backend из сети. И нельзя качать без `--platform` на
Windows/macOS: приедут сборки под чужую ОС, на сервере они не заработают.
Проверить каталог: `ls vendor/*.whl | grep -viE 'manylinux|none-any'` — вывод должен быть пустым.

**Шаг 4. Крайний вариант** — собрать образ там, где интернет есть, и перенести целиком:
```bash
docker compose build backend
docker save audit-backend | gzip > backend.tar.gz
scp backend.tar.gz user@SERVER_IP:/opt/audit-app/
# на сервере:
gunzip -c backend.tar.gz | docker load
docker compose up -d
```

Слой с `pip install` кэшируется: пока `requirements.txt` не меняется, повторные
сборки не ходят в сеть.

## Особенности прод-развёртывания

- `/api/upload-from-path` отключён (env `ENABLE_LOCAL_UPLOAD=0`): на сервере произвольные локальные
  пути не имеют смысла, а эндпоинт открыт для path-traversal. Включать только осознанно.
- Фронт обращается к API по относительному пути (`/api/...`) через тот же nginx, CORS не задействован.
- Бланки чек-листов на сервер класть не нужно: пользователь загружает их в поле 1 вместе с пакетом.
- `verify_ssl_certs=False` для GigaChat оставлено как было (требование интеграции).
- OCR — Tesseract (`rus+eng`), стоит в образе backend. Сканы распознаются при разборе пакета; первый
  разбор пакета со сканами может занять минуты. Тексты кэшируются в `uploads/<сессия>/_texts.json`,
  повторный разбор OCR не запускает. Качество и скорость OCR настраиваются переменными `OCR_DPI`
  (по умолчанию 200) и `OCR_LANG`.
- Жёсткого лимита памяти у backend нет (`mem_reservation: 1g`), пики на крупных PDF уходят в swap.
- Запущен один воркер uvicorn. Проверка идёт в фоновом потоке, у каждой сессии свой статус и своя
  папка результатов, поэтому несколько пользователей могут проверять пакеты одновременно. Статус
  живёт в памяти процесса: после рестарта идущая проверка прерывается, а готовые результаты
  берутся из `_meta.json` сессии.
