const DEFAULT_WORKSHEET_NAME = "Applications";
const DEFAULT_TIMEZONE = "Europe/Kyiv";
const SKIP_VALUE = "Не вказано";
const STATUS_NEW = "Нова";
const STATUS_IN_PROGRESS = "В роботі";
const STATUS_DONE = "Виконана";

const FORM_STEPS = {
  WAITING_NAME: "waiting_name",
  WAITING_PHONE: "waiting_phone",
  WAITING_EMAIL: "waiting_email",
  WAITING_ADDRESS: "waiting_address",
  WAITING_DESCRIPTION: "waiting_description",
  WAITING_CALL_TIME: "waiting_call_time"
};

const SHEET_HEADERS = [
  "request_id",
  "created_at",
  "source",
  "telegram_id",
  "username",
  "name",
  "phone",
  "email",
  "address",
  "description",
  "call_time",
  "status"
];

const TEXT = {
  START: "Вітаємо в боті <b>Асистент з виконання електромонтажних робіт</b>.\n\nОберіть дію нижче.",
  START_APPLICATION: "Починаємо оформлення заявки.\n\nВведіть, будь ласка, ваше ім'я.",
  ASK_PHONE: "Введіть номер телефону у форматі <b>380XXXXXXXXX</b>.",
  ASK_EMAIL: "Введіть email або натисніть <b>Пропустити</b>.",
  ASK_ADDRESS: "Вкажіть адресу або населений пункт, де потрібні роботи.",
  ASK_DESCRIPTION: "Коротко опишіть, які роботи потрібно виконати.",
  ASK_CALL_TIME: "Вкажіть зручний час для дзвінка або натисніть <b>Пропустити</b>.",
  INVALID_NAME: "Ім'я має містити щонайменше 2 символи. Спробуйте ще раз.",
  INVALID_PHONE: "Номер має бути строго у форматі <b>380XXXXXXXXX</b>. Спробуйте ще раз.",
  INVALID_EMAIL: "Email виглядає некоректно. Спробуйте ще раз або натисніть <b>Пропустити</b>.",
  INVALID_ADDRESS: "Будь ласка, вкажіть коректну адресу або населений пункт.",
  INVALID_DESCRIPTION: "Опис має містити щонайменше 5 символів.",
  CANCEL: "Поточну дію скасовано.",
  SUCCESS: "Дякуємо. Вашу заявку збережено.\n\nМи зв'яжемося з вами найближчим часом.",
  SUCCESS_SAVED_WITHOUT_NOTIFY: "Дякуємо. Вашу заявку збережено.\n\nЄ проблема з повідомленням адміністраторів, але заявка не загубилась.",
  SUCCESS_NOTIFIED_WITHOUT_SAVE: "Дякуємо. Заявку отримано й адміністраторів повідомлено.\n\nЄ проблема із записом у таблицю, її треба перевірити в налаштуваннях Google Sheets.",
  SUCCESS_ACCEPTED_WITH_ERRORS: "Заявку отримано, але під час збереження/сповіщення сталася технічна помилка. Спробуйте ще раз пізніше або перевірте логи Worker.",
  NO_REQUESTS: "У вас поки немає заявок.",
  ADMIN_ONLY: "Ця дія доступна тільки адміністраторам.",
  STATUS_UPDATED: "Статус заявки оновлено.",
  UNKNOWN_ACTION: "Не вдалося обробити дію.",
  FALLBACK: "Оберіть одну з доступних дій нижче."
};

export default {
  async fetch(request, env, ctx) {
    try {
      const url = new URL(request.url);

      if (request.method === "GET" && url.pathname === "/") {
        return jsonResponse({ ok: true, service: "telegram-application-bot", status: "running" });
      }

      if (request.method === "GET" && url.pathname === "/setup") {
        return await handleSetup(request, env);
      }

      if (request.method === "GET" && url.pathname === "/health") {
        return await handleHealth(env);
      }

      if (request.method === "POST" && url.pathname === "/webhook") {
        return await handleWebhook(request, env, ctx);
      }

      return new Response("Not Found", { status: 404 });
    } catch (error) {
      console.error("fetch error:", error instanceof Error ? error.stack || error.message : String(error));
      return jsonResponse({ ok: false, error: error instanceof Error ? error.message : String(error) }, 500);
    }
  }
};

async function handleSetup(request, env) {
  const url = new URL(request.url);
  const key = url.searchParams.get("key");

  if (!env.SETUP_KEY || key !== env.SETUP_KEY) {
    return jsonResponse({ ok: false, error: "Unauthorized" }, 401);
  }

  validateTelegramEnv(env);

  const origin = `${url.protocol}//${url.host}`;
  const webhookUrl = `${origin}/webhook`;

  const webhookResult = await telegramApi(env, "setWebhook", {
    url: webhookUrl,
    secret_token: env.TELEGRAM_WEBHOOK_SECRET,
    allowed_updates: ["message", "callback_query"]
  });

  const commandsResult = await telegramApi(env, "setMyCommands", {
    commands: [
      { command: "start", description: "Почати роботу" },
      { command: "my_requests", description: "Мої заявки" },
      { command: "cancel", description: "Скасувати поточну дію" }
    ]
  });

  return jsonResponse({ ok: true, webhookUrl, webhookResult, commandsResult });
}

async function handleHealth(env) {
  const health = {
    ok: true,
    telegram: true,
    kv: Boolean(env.STATE_KV),
    adminsConfigured: parseAdminIds(env.ADMIN_IDS).length > 0,
    googleConfigured: hasGoogleConfig(env)
  };

  if (!health.kv || !health.adminsConfigured) {
    health.ok = false;
  }

  return jsonResponse(health, health.ok ? 200 : 500);
}

async function handleWebhook(request, env, ctx) {
  validateRuntimeEnv(env);

  if (
    env.TELEGRAM_WEBHOOK_SECRET &&
    request.headers.get("X-Telegram-Bot-Api-Secret-Token") !== env.TELEGRAM_WEBHOOK_SECRET
  ) {
    return jsonResponse({ ok: false, error: "Forbidden" }, 403);
  }

  const update = await request.json();
  const updateId = update.update_id;

  if (typeof updateId === "number") {
    const duplicate = await isDuplicateUpdate(env, updateId);
    if (duplicate) {
      return jsonResponse({ ok: true, duplicate: true });
    }
    await markUpdateProcessed(env, updateId);
  }

  ctx.waitUntil(processUpdate(update, env));
  return jsonResponse({ ok: true });
}

async function processUpdate(update, env) {
  try {
    if (update.message) {
      await handleMessage(update.message, env);
    } else if (update.callback_query) {
      await handleCallbackQuery(update.callback_query, env);
    }
  } catch (error) {
    console.error("processUpdate error:", error instanceof Error ? error.stack || error.message : String(error));
  }
}

async function isDuplicateUpdate(env, updateId) {
  const key = `update:${updateId}`;
  const value = await env.STATE_KV.get(key);
  return value === "1";
}

async function markUpdateProcessed(env, updateId) {
  const key = `update:${updateId}`;
  await env.STATE_KV.put(key, "1", { expirationTtl: 60 * 30 });
}

async function handleMessage(message, env) {
  const chatId = message.chat?.id;
  const text = (message.text || "").trim();

  if (!chatId) return;

  if (text === "/start" || text === "Головне меню") {
    await clearState(env, chatId);
    await sendMessage(env, chatId, TEXT.START, { reply_markup: mainMenuKeyboard() });
    return;
  }

  if (text === "/cancel") {
    await clearState(env, chatId);
    await sendMessage(env, chatId, TEXT.CANCEL, { reply_markup: mainMenuKeyboard() });
    return;
  }

  if (text === "/my_requests" || text === "Мої заявки") {
    await clearState(env, chatId);
    await sendUserRequests(env, chatId);
    return;
  }

  if (text === "Створити заявку") {
    await saveState(env, chatId, { step: FORM_STEPS.WAITING_NAME, data: {} });
    await sendMessage(env, chatId, TEXT.START_APPLICATION, { reply_markup: cancelKeyboard() });
    return;
  }

  const state = await getState(env, chatId);

  if (!state) {
    await sendMessage(env, chatId, TEXT.FALLBACK, { reply_markup: mainMenuKeyboard() });
    return;
  }

  await processFormStep(message, state, env);
}

async function processFormStep(message, state, env) {
  const chatId = message.chat.id;
  const text = (message.text || "").trim();
  const data = state.data || {};

  switch (state.step) {
    case FORM_STEPS.WAITING_NAME:
      if (!validateName(text)) {
        await sendMessage(env, chatId, TEXT.INVALID_NAME, { reply_markup: cancelKeyboard() });
        return;
      }
      data.name = text;
      await saveState(env, chatId, { step: FORM_STEPS.WAITING_PHONE, data });
      await sendMessage(env, chatId, TEXT.ASK_PHONE, { reply_markup: cancelKeyboard() });
      return;

    case FORM_STEPS.WAITING_PHONE: {
      const normalizedPhone = normalizePhone(text);
      if (!validatePhone(normalizedPhone)) {
        await sendMessage(env, chatId, TEXT.INVALID_PHONE, { reply_markup: cancelKeyboard() });
        return;
      }
      data.phone = normalizedPhone;
      await saveState(env, chatId, { step: FORM_STEPS.WAITING_EMAIL, data });
      await sendMessage(env, chatId, TEXT.ASK_EMAIL, { reply_markup: skipKeyboard() });
      return;
    }

    case FORM_STEPS.WAITING_EMAIL:
      if (text === "Пропустити") {
        data.email = SKIP_VALUE;
      } else {
        if (!validateEmail(text)) {
          await sendMessage(env, chatId, TEXT.INVALID_EMAIL, { reply_markup: skipKeyboard() });
          return;
        }
        data.email = text;
      }
      await saveState(env, chatId, { step: FORM_STEPS.WAITING_ADDRESS, data });
      await sendMessage(env, chatId, TEXT.ASK_ADDRESS, { reply_markup: cancelKeyboard() });
      return;

    case FORM_STEPS.WAITING_ADDRESS:
      if (!validateAddress(text)) {
        await sendMessage(env, chatId, TEXT.INVALID_ADDRESS, { reply_markup: cancelKeyboard() });
        return;
      }
      data.address = text;
      await saveState(env, chatId, { step: FORM_STEPS.WAITING_DESCRIPTION, data });
      await sendMessage(env, chatId, TEXT.ASK_DESCRIPTION, { reply_markup: cancelKeyboard() });
      return;

    case FORM_STEPS.WAITING_DESCRIPTION:
      if (!validateDescription(text)) {
        await sendMessage(env, chatId, TEXT.INVALID_DESCRIPTION, { reply_markup: cancelKeyboard() });
        return;
      }
      data.description = text;
      await saveState(env, chatId, { step: FORM_STEPS.WAITING_CALL_TIME, data });
      await sendMessage(env, chatId, TEXT.ASK_CALL_TIME, { reply_markup: skipKeyboard() });
      return;

    case FORM_STEPS.WAITING_CALL_TIME: {
      data.call_time = text === "Пропустити" || !text ? SKIP_VALUE : text;
      const payload = buildApplicationPayload(message.from, data, env);

      const dedupeKey = `request-submit:${chatId}:${hashPayload(payload)}`;
      const alreadySubmitted = await env.STATE_KV.get(dedupeKey);
      if (alreadySubmitted === "1") {
        await clearState(env, chatId);
        await sendMessage(env, chatId, TEXT.SUCCESS, { reply_markup: afterSubmitKeyboard() });
        return;
      }

      await env.STATE_KV.put(dedupeKey, "1", { expirationTtl: 60 * 10 });

      let savedToSheet = false;
      let adminsNotified = false;
      let sheetError = null;
      let notifyError = null;

      if (hasGoogleConfig(env)) {
        try {
          await appendApplication(env, payload);
          savedToSheet = true;
        } catch (error) {
          sheetError = error;
          console.error("appendApplication error:", error instanceof Error ? error.stack || error.message : String(error));
        }
      } else {
        sheetError = new Error("Google Sheets is not configured");
        console.error("appendApplication skipped: Google Sheets is not configured");
      }

      try {
        await notifyAdmins(env, payload, sheetError);
        adminsNotified = true;
      } catch (error) {
        notifyError = error;
        console.error("notifyAdmins error:", error instanceof Error ? error.stack || error.message : String(error));
      }

      if (!savedToSheet && !adminsNotified) {
        await env.STATE_KV.delete(dedupeKey);
      }

      await clearState(env, chatId);

      if (savedToSheet && adminsNotified) {
        await sendMessage(env, chatId, TEXT.SUCCESS, { reply_markup: afterSubmitKeyboard() });
        return;
      }

      if (savedToSheet && !adminsNotified) {
        await sendMessage(env, chatId, TEXT.SUCCESS_SAVED_WITHOUT_NOTIFY, { reply_markup: afterSubmitKeyboard() });
        return;
      }

      if (!savedToSheet && adminsNotified) {
        await sendMessage(env, chatId, TEXT.SUCCESS_NOTIFIED_WITHOUT_SAVE, { reply_markup: afterSubmitKeyboard() });
        return;
      }

      await sendMessage(env, chatId, TEXT.SUCCESS_ACCEPTED_WITH_ERRORS, { reply_markup: afterSubmitKeyboard() });
      return;
    }

    default:
      await clearState(env, chatId);
      await sendMessage(env, chatId, TEXT.FALLBACK, { reply_markup: mainMenuKeyboard() });
  }
}

async function handleCallbackQuery(callbackQuery, env) {
  const data = callbackQuery.data || "";
  const fromId = callbackQuery.from?.id;
  const message = callbackQuery.message;

  if (!message?.chat?.id) return;

  if (!isAdmin(env, fromId)) {
    await answerCallbackQuery(env, callbackQuery.id, TEXT.ADMIN_ONLY, true);
    return;
  }

  if (!data.startsWith("status:")) {
    await answerCallbackQuery(env, callbackQuery.id, TEXT.UNKNOWN_ACTION, true);
    return;
  }

  const parts = data.split(":");
  if (parts.length !== 3) {
    await answerCallbackQuery(env, callbackQuery.id, TEXT.UNKNOWN_ACTION, true);
    return;
  }

  const requestId = parts[1];
  const statusCode = parts[2];

  let statusValue = null;
  if (statusCode === "in_progress") statusValue = STATUS_IN_PROGRESS;
  if (statusCode === "done") statusValue = STATUS_DONE;

  if (!statusValue) {
    await answerCallbackQuery(env, callbackQuery.id, TEXT.UNKNOWN_ACTION, true);
    return;
  }

  if (!hasGoogleConfig(env)) {
    await answerCallbackQuery(env, callbackQuery.id, "Google Sheets не налаштований.", true);
    return;
  }

  const updated = await updateApplicationStatus(env, requestId, statusValue);
  if (!updated) {
    await answerCallbackQuery(env, callbackQuery.id, "Не вдалося змінити статус.", true);
    return;
  }

  const currentText = message.text || message.caption || "";
  const newText = replaceStatusInText(currentText, statusValue);

  const payload = {
    chat_id: message.chat.id,
    message_id: message.message_id,
    text: newText,
    parse_mode: "HTML"
  };

  if (statusValue !== STATUS_DONE) {
    payload.reply_markup = adminStatusKeyboard(requestId);
  }

  await telegramApi(env, "editMessageText", payload);
  await answerCallbackQuery(env, callbackQuery.id, TEXT.STATUS_UPDATED, false);
}

function buildApplicationPayload(user, formData, env) {
  const requestId = generateRequestId();
  const timezone = env.TIMEZONE || DEFAULT_TIMEZONE;
  const createdAt = new Intl.DateTimeFormat("uk-UA", {
    timeZone: timezone,
    dateStyle: "short",
    timeStyle: "medium"
  }).format(new Date());

  return {
    request_id: requestId,
    created_at: createdAt,
    source: "telegram",
    telegram_id: String(user?.id || ""),
    username: user?.username ? `@${user.username}` : "",
    name: formData.name || "",
    phone: formData.phone || "",
    email: formData.email || SKIP_VALUE,
    address: formData.address || "",
    description: formData.description || "",
    call_time: formData.call_time || SKIP_VALUE,
    status: STATUS_NEW
  };
}

async function notifyAdmins(env, payload, sheetError = null) {
  const adminIds = parseAdminIds(env.ADMIN_IDS);
  if (!adminIds.length) {
    throw new Error("ADMIN_IDS is empty");
  }

  const text = formatAdminMessage(payload, sheetError);
  let delivered = 0;

  for (const adminId of adminIds) {
    try {
      await sendMessage(env, adminId, text, {
        reply_markup: adminStatusKeyboard(payload.request_id)
      });
      delivered += 1;
    } catch (error) {
      console.error("notifyAdmins delivery error:", adminId, error instanceof Error ? error.stack || error.message : String(error));
    }
  }

  if (delivered === 0) {
    throw new Error("Failed to notify all admins");
  }
}

async function sendUserRequests(env, chatId) {
  if (!hasGoogleConfig(env)) {
    await sendMessage(env, chatId, "Перегляд заявок тимчасово недоступний: Google Sheets не налаштований.", {
      reply_markup: mainMenuKeyboard()
    });
    return;
  }

  const rows = await getApplications(env);
  const userRows = rows.filter((row) => String(row.telegram_id || "") === String(chatId));

  if (!userRows.length) {
    await sendMessage(env, chatId, TEXT.NO_REQUESTS, { reply_markup: mainMenuKeyboard() });
    return;
  }

  const chunks = [];
  for (const row of userRows) {
    chunks.push([
      `🆔 <b>${escapeHtml(row.request_id || "-")}</b>`,
      `📅 ${escapeHtml(row.created_at || "-")}`,
      `📞 ${escapeHtml(row.phone || "-")}`,
      `📍 ${escapeHtml(row.address || "-")}`,
      `📝 ${escapeHtml(row.description || "-")}`,
      `📌 Статус: <b>${escapeHtml(row.status || "-")}</b>`
    ].join("\n"));
  }

  await sendMessage(env, chatId, `<b>Ваші заявки:</b>\n\n${chunks.join("\n\n──────────\n\n")}`, {
    reply_markup: mainMenuKeyboard()
  });
}

function formatAdminMessage(payload, sheetError = null) {
  const lines = [
    "<b>Нова заявка</b>",
    `🆔 <b>${escapeHtml(payload.request_id)}</b>`,
    `📅 <b>Створено:</b> ${escapeHtml(payload.created_at)}`,
    `👤 <b>Ім'я:</b> ${escapeHtml(payload.name)}`,
    `📞 <b>Телефон:</b> ${escapeHtml(payload.phone)}`,
    `✉️ <b>Email:</b> ${escapeHtml(payload.email)}`,
    `📍 <b>Адреса:</b> ${escapeHtml(payload.address)}`,
    `📝 <b>Опис:</b> ${escapeHtml(payload.description)}`,
    `🕒 <b>Зручний час:</b> ${escapeHtml(payload.call_time)}`,
    `💬 <b>Telegram ID:</b> ${escapeHtml(payload.telegram_id)}`,
    `🔗 <b>Username:</b> ${escapeHtml(payload.username || "—")}`,
    `📌 <b>Статус:</b> ${escapeHtml(payload.status)}`
  ];

  if (sheetError) {
    const errorText = sheetError instanceof Error ? sheetError.message : String(sheetError);
    lines.push(`⚠️ <b>Sheets error:</b> ${escapeHtml(errorText)}`);
  }

  return lines.join("\n");
}

function replaceStatusInText(text, newStatus) {
  const lines = text.split("\n");
  let replaced = false;

  const updated = lines.map((line) => {
    if (line.startsWith("📌 <b>Статус:</b>")) {
      replaced = true;
      return `📌 <b>Статус:</b> ${escapeHtml(newStatus)}`;
    }
    return line;
  });

  if (!replaced) {
    updated.push(`📌 <b>Статус:</b> ${escapeHtml(newStatus)}`);
  }

  return updated.join("\n");
}

function mainMenuKeyboard() {
  return {
    keyboard: [[{ text: "Створити заявку" }], [{ text: "Мої заявки" }]],
    resize_keyboard: true
  };
}

function cancelKeyboard() {
  return { keyboard: [[{ text: "/cancel" }]], resize_keyboard: true };
}

function skipKeyboard() {
  return { keyboard: [[{ text: "Пропустити" }], [{ text: "/cancel" }]], resize_keyboard: true };
}

function afterSubmitKeyboard() {
  return {
    keyboard: [[{ text: "Створити заявку" }], [{ text: "Мої заявки" }], [{ text: "Головне меню" }]],
    resize_keyboard: true
  };
}

function adminStatusKeyboard(requestId) {
  return {
    inline_keyboard: [[
      { text: "✅ В роботу", callback_data: `status:${requestId}:in_progress` },
      { text: "🏁 Виконана", callback_data: `status:${requestId}:done` }
    ]]
  };
}

async function sendMessage(env, chatId, text, extra = {}) {
  return await telegramApi(env, "sendMessage", {
    chat_id: chatId,
    text,
    parse_mode: "HTML",
    ...extra
  });
}

async function answerCallbackQuery(env, callbackQueryId, text, showAlert = false) {
  return await telegramApi(env, "answerCallbackQuery", {
    callback_query_id: callbackQueryId,
    text,
    show_alert: showAlert
  });
}

async function telegramApi(env, method, payload) {
  const response = await fetch(`https://api.telegram.org/bot${env.BOT_TOKEN}/${method}`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(payload)
  });

  const result = await response.json();
  if (!response.ok || !result.ok) {
    throw new Error(`Telegram API error in ${method}: ${JSON.stringify(result)}`);
  }

  return result;
}

async function getState(env, chatId) {
  const raw = await env.STATE_KV.get(`state:${chatId}`);
  return raw ? JSON.parse(raw) : null;
}

async function saveState(env, chatId, state) {
  await env.STATE_KV.put(`state:${chatId}`, JSON.stringify(state), {
    expirationTtl: 60 * 60 * 24
  });
}

async function clearState(env, chatId) {
  await env.STATE_KV.delete(`state:${chatId}`);
}

function validateName(value) {
  return value && value.trim().length >= 2;
}

function normalizePhone(value) {
  return (value || "").replace(/[^\d]/g, "");
}

function validatePhone(value) {
  return /^380\d{9}$/.test(value || "");
}

function validateEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value || "");
}

function validateAddress(value) {
  return value && value.trim().length >= 3;
}

function validateDescription(value) {
  return value && value.trim().length >= 5;
}

function parseAdminIds(raw) {
  return String(raw || "").split(",").map((item) => item.trim()).filter(Boolean);
}

function isAdmin(env, userId) {
  return parseAdminIds(env.ADMIN_IDS).includes(String(userId));
}

function generateRequestId() {
  const rand = Math.random().toString(36).slice(2, 7).toUpperCase();
  return `REQ-${Date.now()}-${rand}`;
}

function hashPayload(payload) {
  const base = [
    payload.telegram_id,
    payload.name,
    payload.phone,
    payload.email,
    payload.address,
    payload.description,
    payload.call_time
  ].join("|");

  let hash = 0;
  for (let i = 0; i < base.length; i++) {
    hash = (hash * 31 + base.charCodeAt(i)) >>> 0;
  }
  return String(hash);
}

function validateTelegramEnv(env) {
  const required = ["BOT_TOKEN", "ADMIN_IDS", "STATE_KV"];
  for (const key of required) {
    if (!env[key]) {
      throw new Error(`Missing environment variable: ${key}`);
    }
  }
}

function validateRuntimeEnv(env) {
  validateTelegramEnv(env);
}

function hasGoogleConfig(env) {
  return Boolean(env.GOOGLE_SHEET_ID && env.GOOGLE_SERVICE_ACCOUNT_JSON);
}

function validateGoogleEnv(env) {
  const required = ["GOOGLE_SHEET_ID", "GOOGLE_SERVICE_ACCOUNT_JSON"];
  for (const key of required) {
    if (!env[key]) {
      throw new Error(`Missing environment variable: ${key}`);
    }
  }
}

/* ------------------------------ Google Sheets ------------------------------ */

async function appendApplication(env, application) {
  validateGoogleEnv(env);
  await ensureWorksheet(env);
  const row = SHEET_HEADERS.map((header) => application[header] ?? "");
  await sheetsAppend(env, `${getWorksheetName(env)}!A:L`, [row]);
}

async function getApplications(env) {
  validateGoogleEnv(env);
  await ensureWorksheet(env);
  const values = await sheetsGetValues(env, `${getWorksheetName(env)}!A:L`);
  if (!values.length) return [];

  const headers = values[0];
  const rows = values.slice(1);

  return rows.map((row) => {
    const item = {};
    headers.forEach((header, index) => {
      item[header] = row[index] ?? "";
    });
    return item;
  });
}

async function updateApplicationStatus(env, requestId, newStatus) {
  validateGoogleEnv(env);
  await ensureWorksheet(env);
  const values = await sheetsGetValues(env, `${getWorksheetName(env)}!A:L`);
  if (values.length <= 1) return false;

  const headers = values[0];
  const requestIdIndex = headers.indexOf("request_id");
  const statusIndex = headers.indexOf("status");

  if (requestIdIndex === -1 || statusIndex === -1) return false;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if ((row[requestIdIndex] || "") === requestId) {
      const rowNumber = i + 1;
      const columnLetter = indexToColumn(statusIndex);
      await sheetsUpdateValues(
        env,
        `${getWorksheetName(env)}!${columnLetter}${rowNumber}`,
        [[newStatus]]
      );
      return true;
    }
  }

  return false;
}

async function ensureWorksheet(env) {
  const metadata = await sheetsGetMetadata(env);
  const sheetTitle = getWorksheetName(env);
  const found = metadata.sheets?.find(
    (sheet) => sheet.properties?.title === sheetTitle
  );

  if (!found) {
    await sheetsBatchUpdate(env, {
      requests: [
        {
          addSheet: {
            properties: {
              title: sheetTitle
            }
          }
        }
      ]
    });
  }

  const values = await sheetsGetValues(env, `${sheetTitle}!A1:L1`);
  if (!values.length || !values[0] || values[0].length === 0) {
    await sheetsUpdateValues(env, `${sheetTitle}!A1:L1`, [SHEET_HEADERS]);
  }
}

function getWorksheetName(env) {
  return env.WORKSHEET_NAME || DEFAULT_WORKSHEET_NAME;
}

async function sheetsGetMetadata(env) {
  return await googleApi(
    env,
    `https://sheets.googleapis.com/v4/spreadsheets/${env.GOOGLE_SHEET_ID}`
  );
}

async function sheetsBatchUpdate(env, body) {
  return await googleApi(
    env,
    `https://sheets.googleapis.com/v4/spreadsheets/${env.GOOGLE_SHEET_ID}:batchUpdate`,
    {
      method: "POST",
      body
    }
  );
}

async function sheetsGetValues(env, range) {
  const data = await googleApi(
    env,
    `https://sheets.googleapis.com/v4/spreadsheets/${env.GOOGLE_SHEET_ID}/values/${encodeURIComponent(range)}`
  );

  return data.values || [];
}

async function sheetsUpdateValues(env, range, values) {
  return await googleApi(
    env,
    `https://sheets.googleapis.com/v4/spreadsheets/${env.GOOGLE_SHEET_ID}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`,
    {
      method: "PUT",
      body: {
        range,
        majorDimension: "ROWS",
        values
      }
    }
  );
}

async function sheetsAppend(env, range, values) {
  return await googleApi(
    env,
    `https://sheets.googleapis.com/v4/spreadsheets/${env.GOOGLE_SHEET_ID}/values/${encodeURIComponent(range)}:append?valueInputOption=USER_ENTERED`,
    {
      method: "POST",
      body: {
        range,
        majorDimension: "ROWS",
        values
      }
    }
  );
}

async function googleApi(env, url, options = {}) {
  const token = await getGoogleAccessToken(env);
  const response = await fetch(url, {
    method: options.method || "GET",
    headers: {
      authorization: `Bearer ${token}`,
      "content-type": "application/json"
    },
    body: options.body ? JSON.stringify(options.body) : undefined
  });

  const text = await response.text();
  let data = {};
  try {
    data = text ? JSON.parse(text) : {};
  } catch {
    data = { raw: text };
  }

  if (!response.ok) {
    throw new Error(`Google API error: ${response.status} ${JSON.stringify(data)}`);
  }

  return data;
}

let cachedGoogleToken = null;

async function getGoogleAccessToken(env) {
  validateGoogleEnv(env);

  if (cachedGoogleToken && cachedGoogleToken.expiresAt > Date.now() + 60_000) {
    return cachedGoogleToken.token;
  }

  const serviceAccount = parseServiceAccount(env.GOOGLE_SERVICE_ACCOUNT_JSON);
  const now = Math.floor(Date.now() / 1000);

  const jwtHeader = {
    alg: "RS256",
    typ: "JWT"
  };

  const jwtClaimSet = {
    iss: serviceAccount.client_email,
    scope: "https://www.googleapis.com/auth/spreadsheets",
    aud: serviceAccount.token_uri,
    exp: now + 3600,
    iat: now
  };

  const unsignedToken = `${base64UrlEncode(JSON.stringify(jwtHeader))}.${base64UrlEncode(
    JSON.stringify(jwtClaimSet)
  )}`;

  const signature = await signJwt(unsignedToken, serviceAccount.private_key);
  const assertion = `${unsignedToken}.${signature}`;

  const response = await fetch(serviceAccount.token_uri, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion
    })
  });

  const tokenData = await response.json();
  if (!response.ok) {
    throw new Error(`Google OAuth error: ${JSON.stringify(tokenData)}`);
  }

  cachedGoogleToken = {
    token: tokenData.access_token,
    expiresAt: Date.now() + (tokenData.expires_in || 3600) * 1000
  };

  return cachedGoogleToken.token;
}

function parseServiceAccount(rawValue) {
  const value = String(rawValue || "").trim();

  try {
    return JSON.parse(value);
  } catch {
    try {
      const decoded = atob(value);
      return JSON.parse(decoded);
    } catch {
      throw new Error("GOOGLE_SERVICE_ACCOUNT_JSON must be valid JSON or base64-encoded JSON");
    }
  }
}

async function signJwt(unsignedToken, pemPrivateKey) {
  const keyData = pemToArrayBuffer(pemPrivateKey);

  const cryptoKey = await crypto.subtle.importKey(
    "pkcs8",
    keyData,
    {
      name: "RSASSA-PKCS1-v1_5",
      hash: "SHA-256"
    },
    false,
    ["sign"]
  );

  const signature = await crypto.subtle.sign(
    "RSASSA-PKCS1-v1_5",
    cryptoKey,
    new TextEncoder().encode(unsignedToken)
  );

  return arrayBufferToBase64Url(signature);
}

function pemToArrayBuffer(pem) {
  const base64 = pem
    .replace("-----BEGIN PRIVATE KEY-----", "")
    .replace("-----END PRIVATE KEY-----", "")
    .replace(/\s+/g, "");

  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);

  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }

  return bytes.buffer;
}

function base64UrlEncode(value) {
  const bytes = new TextEncoder().encode(value);
  return arrayBufferToBase64Url(bytes);
}

function arrayBufferToBase64Url(buffer) {
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  let binary = "";

  for (const byte of bytes) {
    binary += String.fromCharCode(byte);
  }

  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

function indexToColumn(index) {
  let column = "";
  let num = index + 1;

  while (num > 0) {
    const remainder = (num - 1) % 26;
    column = String.fromCharCode(65 + remainder) + column;
    num = Math.floor((num - 1) / 26);
  }

  return column;
}

function jsonResponse(data, status = 200) {
  return new Response(JSON.stringify(data, null, 2), {
    status,
    headers: {
      "content-type": "application/json; charset=utf-8"
    }
  });
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}
