# Techno Perspektyva Bot 2.0

Оновлений Telegram-бот для Cloudflare Workers + KV + Google Sheets.

## Що вже додано
- покрокова форма заявки
- підтвердження заявки перед записом у таблицю
- редагування полів перед підтвердженням
- антидубль заявки
- базовий anti-flood
- статуси: Нова / Передзвонити / В роботі / Виконана / Скасована
- автоматичне повідомлення клієнту при зміні статусу
- команди для адміна: /new_requests, /today, /stats
- логування помилок у лист `BotLogs`
- сумісність зі старими записами в існуючій таблиці
- автоматичне розширення заголовків таблиці без видалення старих даних

## Структура
- `src/index.js` — основний код воркера
- `wrangler.jsonc` — конфіг Cloudflare Worker
- `dev.vars.example` — локальні змінні
- `bot-worker-tenj-4170b24bae14.json` — service account JSON

## Поточні аркуші Google Sheets
- основний: `Applications`
- логи: `BotLogs`

## Що зміниться в таблиці
Існуючі колонки не ламаються.

Базові колонки залишаються ті самі:
- request_id
- created_at
- source
- telegram_id
- username
- name
- phone
- email
- address
- description
- call_time
- status

Додатково бот може автоматично дозаписати праворуч нові колонки:
- updated_at
- last_admin_action

Старі рядки залишаться на місці.

## Як задеплоїти
1. Завантаж цей проєкт у Cloudflare Worker
2. Переконайся, що KV namespace `STATE_KV` існує
3. У Variables/Secrets мають бути:
   - BOT_TOKEN
   - GOOGLE_SHEET_ID
   - ADMIN_IDS
   - WORKSHEET_NAME
   - LOGS_WORKSHEET_NAME
   - TIMEZONE
   - SETUP_KEY
   - TELEGRAM_WEBHOOK_SECRET
   - GOOGLE_SERVICE_ACCOUNT_JSON
4. Після деплою відкрий:
   - `https://YOUR-WORKER.workers.dev/setup?key=techno-perspektyva-setup-2026`

## Нові команди
- `/start`
- `/my_requests`
- `/new_requests`
- `/today`
- `/stats`
- `/cancel`

## Важливо
- якщо деплоїш через GitHub, не публікуй секрети у відкритий репозиторій
- service account повинен мати доступ редактора до таблиці
- бот сам створить аркуш `BotLogs`, якщо його ще немає
