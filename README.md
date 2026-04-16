# Система задач (MBC)

Фронтенд (статика + `script.js`) и **Node.js API** с **PostgreSQL**: единое состояние приложения в JSONB, пользователи в таблице `users`, JWT, миграции SQL.

## Локальный запуск

1. Установите [Node.js 20+](https://nodejs.org/).
2. Скопируйте `.env.example` в `.env`, укажите `DATABASE_URL` (локальный Postgres или облако).
3. В production задайте **`JWT_SECRET`** (длинная случайная строка) и **`ADMIN_PHONE` / `ADMIN_PASSWORD`** — при первом старте создаётся первый администратор в БД.
4. Установка и старт:

```bash
npm install
npm start
```

Откройте `http://localhost:3000`. Миграции выполняются автоматически при запуске.

## Деплой в Railway + GitHub

1. Залейте репозиторий на **GitHub** (корень проекта — здесь же `package.json` и `server/`).
2. В [Railway](https://railway.app): **New Project** → **Deploy from GitHub** → выберите репозиторий (если пишет *Repo not found* — заново подключите GitHub в настройках аккаунта Railway и выдайте доступ к репозиторию `Alisher1994/tasks`).
3. **PostgreSQL обязателен:** в том же проекте нажмите **+ New** → **Database** → **PostgreSQL**. Дождитесь, пока появится отдельный сервис «Postgres».
4. **Связать БД с приложением (иначе будет ошибка `DATABASE_URL`):**
   - Откройте сервис **вашего приложения** (не Postgres) → вкладка **Variables**.
   - **Add variable** → **Add reference** (или **Reference Variable**).
   - Выберите сервис **PostgreSQL** → переменную **`DATABASE_URL`** → сохраните.  
     В списке переменных приложения должна появиться `DATABASE_URL` со значением вида `postgresql://...`.
5. В тех же **Variables** приложения добавьте вручную:
   - **`JWT_SECRET`** — длинная случайная строка (обязательно в production, иначе сервер не стартует);
   - **`ADMIN_PHONE`**, **`ADMIN_PASSWORD`** — первый администратор при первом запуске;
   - **`CLOUDINARY_CLOUD_NAME`**, **`CLOUDINARY_API_KEY`**, **`CLOUDINARY_API_SECRET`** — постоянное хранилище фото задач (рекомендуется, чтобы медиа не терялись при деплоях);
   - при необходимости: `TELEGRAM_BOT_TOKEN`, `TELEGRAM_WEBHOOK_SECRET`, `PUBLIC_APP_URL` (если домен не подставляется автоматически при вызове регистрации webhook).
6. Команда старта: **`npm start`** (уже в `package.json`). **PORT** задавать не нужно — задаёт Railway.

Предупреждение npm `Use --omit=dev` на деплое можно игнорировать — на работу приложения не влияет.

### Если в логах снова «Ошибка: задайте переменную DATABASE_URL»

- Убедитесь, что в **Variables именно сервиса с Node** есть **`DATABASE_URL`** (через reference из Postgres), а не только у сервиса базы.
- После добавления переменной нажмите **Redeploy** у сервиса приложения.

После деплоя откройте выданный домен Railway. Вход — телефон/пароль из `ADMIN_*` (первый раз), затем можно добавлять пользователей через API (см. ниже).

## API (кратко)

| Метод | Путь | Описание |
|--------|------|----------|
| `GET` | `/health` | Проверка сервера и БД |
| `POST` | `/api/auth/login` | Вход, JSON `{ phone, password }` → `{ token, displayName }` |
| `GET` | `/api/auth/me` | Текущий пользователь (Bearer) |
| `GET` | `/api/data` | Снимок данных приложения (Bearer) |
| `PUT` | `/api/data` | Сохранение снимка (Bearer), тело `{ data: { ... } }` |
| `POST` | `/api/telegram/set-webhook` | Регистрация webhook бота на текущем домене (Bearer); после сохранения токена в приложении |
| `POST` | `/api/media/upload` | Загрузка медиа в Cloudinary (Bearer), тело `{ dataUrl, fileName }` |
| `GET` | `/api/admin/users` | Список пользователей (**admin**) |
| `POST` | `/api/admin/users` | Создать пользователя (**admin**), `{ phone, password, displayName?, role? }` |
| `DELETE` | `/api/admin/users/:id` | Удалить пользователя (**admin**) |

Telegram: после входа в приложение нажмите **«Сохранить токен»** в «Прочие настройки» → Telegram — сервер вызовет `setWebhook` на `https://<ваш-домен>/api/telegram/webhook`. Команда `/start` в боте привязывает **реальный** Telegram user id к сотруднику в справочнике (и приветствие в чате). Надёжная привязка: ссылка `https://t.me/<бот>?start=e_<ID>` (ID из колонки сотрудника). Если задан `TELEGRAM_WEBHOOK_SECRET`, тот же секрет уходит в Telegram и проверяется заголовком входящих запросов.

## Архитектура данных

- **`app_state`**: одна строка `id=1`, поле `payload` (JSON) — секции задач, справочники, настройки отображения и т.д.
- **`users`**: логины с bcrypt, роли `admin` / `user`.
- **`report_shares`**: временные ссылки на отчёты.

Полноценная нормализация «каждая задача — строка SQL» не используется: модель осознанно документная (один JSON на приложение), что упрощает синхронизацию с текущим клиентом.
