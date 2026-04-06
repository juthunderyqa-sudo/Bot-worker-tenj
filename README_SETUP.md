# Haircut Booking Bot — Setup

## Що вже готово
- пароль при першому вході
- реєстрація клієнта
- запис на вільні слоти з 09:00 до 18:00
- статуси: Очікує підтвердження / Підтверджено / Відхилено / Скасовано
- "Мої заявки"
- скасування бронювання клієнтом
- адмінські списки pending / today / stats / find_phone

## Структура
- `src/index.js` — основна логіка бота
- `wrangler.jsonc` — конфіг Cloudflare Worker
- `dev.vars.example` — приклад локальних змінних
- `ARCHITECTURE.md` — схема архітектури

## Що треба заповнити
У `wrangler.jsonc`:
- `kv_namespaces[0].id`
- `GOOGLE_SHEET_ID`
- `ADMIN_IDS`
- `SETUP_KEY`
- `TELEGRAM_WEBHOOK_SECRET`

У secrets Cloudflare:
- `BOT_TOKEN`
- `GOOGLE_SERVICE_ACCOUNT_JSON`
- `ACCESS_PASSWORD`

## Локальний запуск
1. Скопіюйте `dev.vars.example` у `dev.vars`
2. Заповніть значення
3. Запустіть:
   - `npm install`
   - `npm run dev`

## Деплой
1. `npm install`
2. `wrangler secret put BOT_TOKEN`
3. `wrangler secret put GOOGLE_SERVICE_ACCOUNT_JSON`
4. `wrangler secret put ACCESS_PASSWORD`
5. `wrangler deploy`
6. Відкрийте:
   - `https://YOUR-WORKER.workers.dev/setup?key=YOUR_SETUP_KEY`

## Після setup
Бот сам створить листи:
- `Users`
- `Bookings`
- `BotLogs`

## Якщо пароль не питає
- перевірте, що `ACCESS_PASSWORD` доданий саме як secret
- перевірте, що задеплоєний новий код
- якщо користувач уже зареєстрований, видаліть його рядок із листа `Users`
