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
2. В [Railway](https://railway.app): **New Project** → **Deploy from GitHub** → выберите репозиторий.
3. Добавьте плагин **PostgreSQL**; в сервисе приложения подключите переменную **`DATABASE_URL`** (Railway подставит из плагина).
4. В **Variables** сервиса приложения задайте:
   - `JWT_SECRET` — обязательно в production;
   - `ADMIN_PHONE`, `ADMIN_PASSWORD` — для создания первого администратора при первом деплое;
   - при необходимости: `TELEGRAM_BOT_TOKEN`, `TELEGRAM_WEBHOOK_SECRET`.
5. Команда старта: **`npm start`** (уже указана в `package.json`).

После деплоя откройте выданный домен Railway. Вход — телефон/пароль из `ADMIN_*` (первый раз), затем можно добавлять пользователей через API (см. ниже).

## API (кратко)

| Метод | Путь | Описание |
|--------|------|----------|
| `GET` | `/health` | Проверка сервера и БД |
| `POST` | `/api/auth/login` | Вход, JSON `{ phone, password }` → `{ token, displayName }` |
| `GET` | `/api/auth/me` | Текущий пользователь (Bearer) |
| `GET` | `/api/data` | Снимок данных приложения (Bearer) |
| `PUT` | `/api/data` | Сохранение снимка (Bearer), тело `{ data: { ... } }` |
| `GET` | `/api/admin/users` | Список пользователей (**admin**) |
| `POST` | `/api/admin/users` | Создать пользователя (**admin**), `{ phone, password, displayName?, role? }` |
| `DELETE` | `/api/admin/users/:id` | Удалить пользователя (**admin**) |

Telegram: `POST /api/telegram/webhook` — URL в настройках бота на ваш домен.

## Архитектура данных

- **`app_state`**: одна строка `id=1`, поле `payload` (JSON) — секции задач, справочники, настройки отображения и т.д.
- **`users`**: логины с bcrypt, роли `admin` / `user`.
- **`report_shares`**: временные ссылки на отчёты.

Полноценная нормализация «каждая задача — строка SQL» не используется: модель осознанно документная (один JSON на приложение), что упрощает синхронизацию с текущим клиентом.
