# Асистент  — Cloudflare Workers

Готовий шаблон Telegram-бота для прийому заявок через Cloudflare Workers + KV + Google Sheets.

## Можливості
- прийом заявок українською
- ім'я
- телефон
- email (опціонально)
- адреса
- опис задачі
- зручний час для дзвінка (опціонально)
- запис у Google Sheets
- повідомлення двом адміністраторам
- статуси: Нова / В роботі / Виконана
- команда /my_requests
- збереження стану діалогу в KV

## Що треба додати в Cloudflare
### Variables
- BOT_TOKEN
- GOOGLE_SHEET_ID
- ADMIN_IDS
- WORKSHEET_NAME
- TIMEZONE
- SETUP_KEY
- TELEGRAM_WEBHOOK_SECRET

### Secret
- GOOGLE_SERVICE_ACCOUNT_JSON

## Як підготувати таблицю
1. Створи нову Google Sheets таблицю
2. Скопіюй її Sheet ID
3. Поділись таблицею з email service account як Редактор

## Ендпоїнти
- / — health check
- /setup?key=... — встановлення webhook і команд
- /webhook — Telegram webhook

## Деплой без терміналу
1. Завантаж проєкт у GitHub
2. У Cloudflare відкрий Workers & Pages
3. Create application
4. Import a repository
5. Обери GitHub репозиторій
6. Створи KV namespace і встав його ID у wrangler.jsonc
7. Додай Variables і Secrets
8. Після деплою відкрий:
   https://YOUR-WORKER.workers.dev/setup?key=ТВОЄ_ЗНАЧЕННЯ_SETUP_KEY

## Важливо
Не заливай токени і Google key у GitHub.
