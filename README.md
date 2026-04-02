# Асистент — Cloudflare Worker + Telegram + Google Sheets

Готовий Telegram-бот для прийому заявок через Cloudflare Workers, KV і Google Sheets.

## Що виправлено

- webhook більше не падає, якщо Google Sheets налаштований криво
- повідомлення адміністраторам відправляється окремо від запису в таблицю
- якщо Sheets не працює, адмін все одно отримає заявку
- якщо не вдалося і в Sheets, і в Telegram admin notify, заявка не блокується назавжди dedupe-ключем
- `GOOGLE_SERVICE_ACCOUNT_JSON` підтримує як raw JSON, так і base64 JSON
- додано `/health` для швидкої перевірки конфігурації
- приведено файли проєкту до нормальної структури

## Можливості

- прийом заявок українською
- ім'я
- телефон
- email (опціонально)
- адреса
- опис задачі
- зручний час для дзвінка (опціонально)
- запис у Google Sheets
- повідомлення адміністраторам
- статуси: Нова / В роботі / Виконана
- команда `/my_requests`
- збереження стану діалогу в KV

## Структура

- `index.js` — код Worker
- `wrangler.jsonc` — конфіг Cloudflare Worker
- `package.json` — npm scripts
- `dev.vars.example` — приклад локальних змінних
- `bot-worker-tenj-4170b24bae14.json` — твій test service account JSON

## Що треба в Cloudflare

### Variables

- `BOT_TOKEN`
- `GOOGLE_SHEET_ID`
- `ADMIN_IDS`
- `WORKSHEET_NAME`
- `TIMEZONE`
- `SETUP_KEY`
- `TELEGRAM_WEBHOOK_SECRET`

### Secret

- `GOOGLE_SERVICE_ACCOUNT_JSON`

### KV binding

- `STATE_KV`

## Як підготувати Google Sheets

1. Створи Google Sheet
2. Візьми ID таблиці з URL
3. Поділись таблицею з email service account як **Editor**
4. Встав `GOOGLE_SHEET_ID` у Cloudflare
5. Встав service account JSON у secret `GOOGLE_SERVICE_ACCOUNT_JSON`

## Локальний запуск

```bash
npm install
npm run dev
```

## Деплой

```bash
npm install
npm run deploy
```

## Після деплою

Відкрий:

```text
https://YOUR-WORKER.workers.dev/setup?key=YOUR_SETUP_KEY
```

## Перевірка

- `GET /` — simple status
- `GET /health` — чи є базові конфіги
- `GET /setup?key=...` — встановити webhook і команди
- `POST /webhook` — Telegram webhook

## Важливо

Якщо адмінам не приходять повідомлення, переконайся, що:

- `ADMIN_IDS` заданий через кому без пробілів або з пробілами між значеннями — це ок
- бот уже має відкритий чат з кожним адміном
- Telegram не блокує відправку через `Forbidden: bot can't initiate conversation`

Якщо не пише в Google Sheets, перевір:

- `GOOGLE_SHEET_ID`
- чи пошерена таблиця на service account email
- чи `GOOGLE_SERVICE_ACCOUNT_JSON` переданий як валідний JSON або base64 JSON
