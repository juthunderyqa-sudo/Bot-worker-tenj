const DEFAULT_WORKSHEET_NAME = "Applications";
const DEFAULT_LOGS_WORKSHEET_NAME = "BotLogs";
const DEFAULT_TIMEZONE = "Europe/Kyiv";
const SKIP_VALUE = "Не вказано";
const STATUS_NEW = "Нова";
const STATUS_CALLBACK = "Передзвонити";
const STATUS_IN_PROGRESS = "В роботі";
const STATUS_DONE = "Виконана";
const STATUS_CANCELLED = "Скасована";
const MAX_DESCRIPTION_LENGTH = 1000;
const MAX_PHOTOS = 5;

const FORM_STEPS = {
  WAITING_NAME: "waiting_name",
  WAITING_PHONE: "waiting_phone",
  WAITING_EMAIL: "waiting_email",
  WAITING_ADDRESS: "waiting_address",
  WAITING_DESCRIPTION: "waiting_description",
  WAITING_CALL_TIME: "waiting_call_time",
  WAITING_PHOTOS: "waiting_photos",
  WAITING_CONFIRMATION: "waiting_confirmation"
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

const OPTIONAL_EXTRA_HEADERS = [
  "updated_at",
  "last_admin_action",
  "location_lat",
  "location_lon",
  "location_label",
  "photo_file_ids",
  "photo_count"
];

const LOG_HEADERS = ["created_at", "level", "scope", "message", "details"];

const TEXT = {
  START: "Вітаємо в боті <b>Асистент з виконання електромонтажних робіт</b>.\n\nОберіть дію нижче.",
  START_APPLICATION: "Починаємо оформлення заявки.\n\nВведіть, будь ласка, ваше ім'я.",
  ASK_PHONE: "Введіть номер телефону у форматі <b>380XXXXXXXXX</b>.",
  ASK_EMAIL: "Введіть email або натисніть <b>Пропустити</b>.",
  ASK_ADDRESS: "Надішліть адресу текстом. Геолокація — лише за бажанням, через окрему кнопку нижче.",
  ASK_DESCRIPTION: "Коротко опишіть, які роботи потрібно виконати.",
  ASK_CALL_TIME: "Вкажіть зручний час для дзвінка або натисніть <b>Пропустити</b>.",
  ASK_PHOTOS: "Надішліть 1-5 фото об'єкта/проблеми. Коли завершите — натисніть <b>Готово з фото</b> або <b>Пропустити</b>.",
  ASK_CONFIRM: "Перевірте заявку нижче.\n\nЯкщо все правильно — натисніть <b>Підтвердити заявку</b>.",
  INVALID_NAME: "Ім'я має містити щонайменше 2 символи. Спробуйте ще раз.",
  INVALID_PHONE: "Номер має бути строго у форматі <b>380XXXXXXXXX</b>. Спробуйте ще раз.",
  INVALID_EMAIL: "Email виглядає некоректно. Спробуйте ще раз або натисніть <b>Пропустити</b>.",
  INVALID_ADDRESS: "Будь ласка, вкажіть адресу текстом. Геолокацію можна надсилати лише за бажанням.",
  INVALID_DESCRIPTION: "Опис має містити щонайменше 5 символів.",
  TOO_LONG_DESCRIPTION: "Опис занадто довгий. Скоротіть його до 1000 символів.",
  INVALID_PHOTO_STEP: "Тут очікується фото. Надішліть фото, натисніть <b>Готово з фото</b> або <b>Пропустити</b>.",
  PHOTO_SAVED: "Фото додано до заявки.",
  PHOTO_LIMIT: "До заявки можна додати максимум 5 фото.",
  PHOTO_DONE_EMPTY: "Спочатку надішліть хоча б одне фото або натисніть <b>Пропустити</b>.",
  CANCEL: "Поточну дію скасовано.",
  SUCCESS: "Дякуємо. Вашу заявку збережено.\n\nМи зв'яжемося з вами найближчим часом.",
  NO_REQUESTS: "У вас поки немає заявок.",
  ADMIN_ONLY: "Ця дія доступна тільки адміністраторам.",
  STATUS_UPDATED: "Статус заявки оновлено.",
  UNKNOWN_ACTION: "Не вдалося обробити дію.",
  FALLBACK: "Оберіть одну з доступних дій нижче.",
  FLOOD_LIMIT: "Забагато запитів за короткий час. Спробуйте ще раз трохи пізніше.",
  ACTIVE_FORM_EXISTS: "У вас уже є незавершена заявка. Або завершіть її, або натисніть /cancel.",
  NOTHING_TO_CONFIRM: "Немає заявки для підтвердження.",
  USER_NOTIFIED: "Клієнта повідомлено про зміну статусу.",
  USER_NOTIFY_FAILED: "Статус змінено, але клієнту повідомлення не відправилося.",
  TODAY_EMPTY: "Сьогодні нових заявок ще не було.",
  NEW_EMPTY: "Немає заявок зі статусом 'Нова'.",
  STATS_EMPTY: "Поки що недостатньо даних для статистики.",
  EDIT_PROMPT: "Що хочете змінити у заявці?",
  FIND_PHONE_USAGE: "Використання: <b>/find_phone 380XXXXXXXXX</b>",
  FIND_PHONE_EMPTY: "За цим номером заявок не знайдено."
};

const CALLBACKS = {
  STATUS_PREFIX: "status:",
  EDIT_PREFIX: "edit:",
  CONFIRM_SUBMIT: "confirm:submit",
  CONFIRM_CANCEL: "confirm:cancel"
};

const EDIT_FIELDS = {
  name: FORM_STEPS.WAITING_NAME,
  phone: FORM_STEPS.WAITING_PHONE,
  email: FORM_STEPS.WAITING_EMAIL,
  address: FORM_STEPS.WAITING_ADDRESS,
  description: FORM_STEPS.WAITING_DESCRIPTION,
  call_time: FORM_STEPS.WAITING_CALL_TIME,
  photos: FORM_STEPS.WAITING_PHOTOS
};

export default {
  async fetch(request, env, ctx) {
    try {
      const url = new URL(request.url);

      if (request.method === "GET" && url.pathname === "/") {
        return jsonResponse({ ok: true, service: "techno-perspektyva-bot", status: "running" });
      }

      if (request.method === "GET" && url.pathname === "/setup") {
        return await handleSetup(request, env);
      }

      if (request.method === "POST" && url.pathname === "/webhook") {
        return await handleWebhook(request, env, ctx);
      }

      return new Response("Not Found", { status: 404 });
    } catch (error) {
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

  validateEnv(env);

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
      { command: "new_requests", description: "Нові заявки (адмін)" },
      { command: "today", description: "Заявки за сьогодні (адмін)" },
      { command: "stats", description: "Статистика (адмін)" },
      { command: "find_phone", description: "Пошук за номером (адмін)" },
      { command: "cancel", description: "Скасувати поточну дію" }
    ]
  });

  return jsonResponse({ ok: true, webhookUrl, webhookResult, commandsResult });
}

async function handleWebhook(request, env, ctx) {
  validateEnv(env);

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
    await logError(env, "processUpdate", error, { update });
  }
}

async function isDuplicateUpdate(env, updateId) {
  const key = `update:${updateId}`;
  const value = await env.STATE_KV.get(key);
  return value === "1";
}

async function markUpdateProcessed(env, updateId) {
  await env.STATE_KV.put(`update:${updateId}`, "1", { expirationTtl: 60 * 30 });
}

async function handleMessage(message, env) {
  const chatId = message.chat?.id;
  const text = (message.text || "").trim();

  if (!chatId) return;

  if (!(await checkRateLimit(env, chatId, message.from?.id))) {
    await sendMessage(env, chatId, TEXT.FLOOD_LIMIT, { reply_markup: mainMenuKeyboard(isAdmin(env, message.from?.id)) });
    return;
  }

  if (text === "/start" || text === "Головне меню") {
    await clearState(env, chatId);
    await sendMessage(env, chatId, TEXT.START, { reply_markup: mainMenuKeyboard(isAdmin(env, message.from?.id)) });
    return;
  }

  if (text === "/cancel") {
    await clearState(env, chatId);
    await sendMessage(env, chatId, TEXT.CANCEL, { reply_markup: mainMenuKeyboard(isAdmin(env, message.from?.id)) });
    return;
  }

  if (text === "/my_requests" || text === "Мої заявки") {
    await clearState(env, chatId);
    await sendUserRequests(env, chatId, message.from?.id);
    return;
  }

  if (text === "/new_requests") {
    if (!isAdmin(env, message.from?.id)) {
      await sendMessage(env, chatId, TEXT.ADMIN_ONLY, { reply_markup: mainMenuKeyboard(false) });
      return;
    }
    await sendAdminRequestsByStatus(env, chatId, STATUS_NEW, TEXT.NEW_EMPTY);
    return;
  }

  if (text === "/today") {
    if (!isAdmin(env, message.from?.id)) {
      await sendMessage(env, chatId, TEXT.ADMIN_ONLY, { reply_markup: mainMenuKeyboard(false) });
      return;
    }
    await sendTodayRequests(env, chatId);
    return;
  }

  if (text === "/stats") {
    if (!isAdmin(env, message.from?.id)) {
      await sendMessage(env, chatId, TEXT.ADMIN_ONLY, { reply_markup: mainMenuKeyboard(false) });
      return;
    }
    await sendStats(env, chatId);
    return;
  }

  if (text.startsWith("/find_phone")) {
    if (!isAdmin(env, message.from?.id)) {
      await sendMessage(env, chatId, TEXT.ADMIN_ONLY, { reply_markup: mainMenuKeyboard(false) });
      return;
    }
    await findRequestsByPhone(env, chatId, text);
    return;
  }

  if (text === "Створити заявку") {
    const state = await getState(env, chatId);
    if (state && !state.isEditing) {
      await sendMessage(env, chatId, TEXT.ACTIVE_FORM_EXISTS, { reply_markup: cancelKeyboard() });
      return;
    }
    await saveState(env, chatId, newState(FORM_STEPS.WAITING_NAME, {}, false));
    await sendMessage(env, chatId, TEXT.START_APPLICATION, { reply_markup: cancelKeyboard() });
    return;
  }

  const state = await getState(env, chatId);
  if (!state) {
    await sendMessage(env, chatId, TEXT.FALLBACK, { reply_markup: mainMenuKeyboard(isAdmin(env, message.from?.id)) });
    return;
  }

  await processFormStep(message, state, env);
}

async function processFormStep(message, state, env) {
  const chatId = message.chat.id;
  const text = (message.text || "").trim();
  const data = normalizeDraft(state.data || {});

  switch (state.step) {
    case FORM_STEPS.WAITING_NAME:
      if (!validateName(text)) {
        await sendMessage(env, chatId, TEXT.INVALID_NAME, { reply_markup: cancelKeyboard() });
        return;
      }
      data.name = text;
      await moveToNextStep(env, chatId, state, data, FORM_STEPS.WAITING_PHONE, TEXT.ASK_PHONE, cancelKeyboard());
      return;

    case FORM_STEPS.WAITING_PHONE: {
      const normalizedPhone = normalizePhone(text);
      if (!validatePhone(normalizedPhone)) {
        await sendMessage(env, chatId, TEXT.INVALID_PHONE, { reply_markup: cancelKeyboard() });
        return;
      }
      data.phone = normalizedPhone;
      await moveToNextStep(env, chatId, state, data, FORM_STEPS.WAITING_EMAIL, TEXT.ASK_EMAIL, skipKeyboard());
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
      await moveToNextStep(env, chatId, state, data, FORM_STEPS.WAITING_ADDRESS, TEXT.ASK_ADDRESS, addressKeyboard());
      return;

    case FORM_STEPS.WAITING_ADDRESS:
      if (message.location) {
        data.location_lat = String(message.location.latitude);
        data.location_lon = String(message.location.longitude);
        data.location_label = `https://maps.google.com/?q=${message.location.latitude},${message.location.longitude}`;
        if (!data.address || data.address === SKIP_VALUE || data.address === "Геолокація") {
          data.address = "Геолокація";
        }
      } else {
        if (text === 'Продовжити без геолокації') {
          await sendMessage(env, chatId, 'Вкажіть адресу текстом одним повідомленням.', { reply_markup: addressKeyboard() });
          return;
        }
        if (!validateAddress(text)) {
          await sendMessage(env, chatId, TEXT.INVALID_ADDRESS, { reply_markup: addressKeyboard() });
          return;
        }
        data.address = text;
      }
      await moveToNextStep(env, chatId, state, data, FORM_STEPS.WAITING_DESCRIPTION, TEXT.ASK_DESCRIPTION, cancelKeyboard());
      return;

    case FORM_STEPS.WAITING_DESCRIPTION:
      if (!validateDescription(text)) {
        await sendMessage(env, chatId, TEXT.INVALID_DESCRIPTION, { reply_markup: cancelKeyboard() });
        return;
      }
      if (text.length > MAX_DESCRIPTION_LENGTH) {
        await sendMessage(env, chatId, TEXT.TOO_LONG_DESCRIPTION, { reply_markup: cancelKeyboard() });
        return;
      }
      data.description = text;
      await moveToNextStep(env, chatId, state, data, FORM_STEPS.WAITING_CALL_TIME, TEXT.ASK_CALL_TIME, skipKeyboard());
      return;

    case FORM_STEPS.WAITING_CALL_TIME:
      data.call_time = text === "Пропустити" || !text ? SKIP_VALUE : text;
      await moveToNextStep(env, chatId, state, data, FORM_STEPS.WAITING_PHOTOS, TEXT.ASK_PHOTOS, photoKeyboard(true));
      return;

    case FORM_STEPS.WAITING_PHOTOS:
      await handlePhotoStep(message, state, data, env);
      return;

    case FORM_STEPS.WAITING_CONFIRMATION:
      await sendMessage(env, chatId, TEXT.NOTHING_TO_CONFIRM, { reply_markup: confirmKeyboard() });
      return;

    default:
      await clearState(env, chatId);
      await sendMessage(env, chatId, TEXT.FALLBACK, { reply_markup: mainMenuKeyboard(isAdmin(env, message.from?.id)) });
  }
}

async function handlePhotoStep(message, state, data, env) {
  const chatId = message.chat.id;
  const text = (message.text || "").trim();
  const photos = Array.isArray(data.photo_file_ids) ? data.photo_file_ids : [];

  if (text === "Пропустити") {
    data.photo_file_ids = [];
    data.photo_count = "0";
    await sendConfirmation(env, chatId, data);
    return;
  }

  if (text === "Готово з фото") {
    if (!photos.length) {
      await sendMessage(env, chatId, TEXT.PHOTO_DONE_EMPTY, { reply_markup: photoKeyboard(true) });
      return;
    }
    data.photo_count = String(photos.length);
    await sendConfirmation(env, chatId, data);
    return;
  }

  if (Array.isArray(message.photo) && message.photo.length) {
    if (photos.length >= MAX_PHOTOS) {
      await sendMessage(env, chatId, TEXT.PHOTO_LIMIT, { reply_markup: photoKeyboard(true) });
      return;
    }

    const best = message.photo[message.photo.length - 1];
    if (!photos.includes(best.file_id)) {
      photos.push(best.file_id);
    }

    data.photo_file_ids = photos;
    data.photo_count = String(photos.length);
    await saveState(env, chatId, newState(FORM_STEPS.WAITING_PHOTOS, data, state.isEditing));
    await sendMessage(
      env,
      chatId,
      `${TEXT.PHOTO_SAVED} (${photos.length}/${MAX_PHOTOS})`,
      { reply_markup: photoKeyboard(true) }
    );
    return;
  }

  await sendMessage(env, chatId, TEXT.INVALID_PHOTO_STEP, { reply_markup: photoKeyboard(true) });
}

async function moveToNextStep(env, chatId, state, data, nextStep, nextText, keyboard) {
  if (state.isEditing) {
    await sendConfirmation(env, chatId, data);
    return;
  }
  await saveState(env, chatId, newState(nextStep, data, false));
  await sendMessage(env, chatId, nextText, { reply_markup: keyboard });
}

async function sendConfirmation(env, chatId, data) {
  await saveState(env, chatId, newState(FORM_STEPS.WAITING_CONFIRMATION, data, false));
  await sendMessage(env, chatId, `${TEXT.ASK_CONFIRM}\n\n${formatClientPreview(data)}`, {
    reply_markup: confirmKeyboard()
  });
}

async function handleCallbackQuery(callbackQuery, env) {
  const data = callbackQuery.data || "";
  const fromId = callbackQuery.from?.id;
  const message = callbackQuery.message;

  if (!message?.chat?.id) return;

  try {
    if (data === CALLBACKS.CONFIRM_SUBMIT) {
      await handleConfirmSubmit(callbackQuery, env);
      return;
    }

    if (data === CALLBACKS.CONFIRM_CANCEL) {
      await clearState(message.chat.id ? env : env, message.chat.id);
      await answerCallbackQuery(env, callbackQuery.id, TEXT.CANCEL, false);
      await sendMessage(env, message.chat.id, TEXT.CANCEL, { reply_markup: mainMenuKeyboard(isAdmin(env, fromId)) });
      return;
    }

    if (data.startsWith(CALLBACKS.EDIT_PREFIX)) {
      await handleEditCallback(callbackQuery, env);
      return;
    }

    if (!data.startsWith(CALLBACKS.STATUS_PREFIX)) {
      await answerCallbackQuery(env, callbackQuery.id, TEXT.UNKNOWN_ACTION, true);
      return;
    }

    if (!isAdmin(env, fromId)) {
      await answerCallbackQuery(env, callbackQuery.id, TEXT.ADMIN_ONLY, true);
      return;
    }

    const parts = data.split(":");
    if (parts.length !== 3) {
      await answerCallbackQuery(env, callbackQuery.id, TEXT.UNKNOWN_ACTION, true);
      return;
    }

    const requestId = parts[1];
    const statusValue = mapStatusCode(parts[2]);
    if (!statusValue) {
      await answerCallbackQuery(env, callbackQuery.id, TEXT.UNKNOWN_ACTION, true);
      return;
    }

    const updated = await updateApplicationStatus(env, requestId, statusValue, callbackQuery.from);
    if (!updated?.ok) {
      await answerCallbackQuery(env, callbackQuery.id, "Не вдалося змінити статус.", true);
      return;
    }

    const currentText = message.text || message.caption || "";
    const newText = replaceStatusInText(currentText, statusValue, updated.updatedAt, callbackQuery.from);
    const payload = {
      chat_id: message.chat.id,
      message_id: message.message_id,
      text: newText,
      parse_mode: "HTML"
    };

    if (![STATUS_DONE, STATUS_CANCELLED].includes(statusValue)) {
      payload.reply_markup = adminStatusKeyboard(requestId);
    }

    await telegramApi(env, "editMessageText", payload);

    const notifyResult = await notifyClientAboutStatus(env, updated.row, statusValue);
    await answerCallbackQuery(
      env,
      callbackQuery.id,
      notifyResult ? `${TEXT.STATUS_UPDATED}. ${TEXT.USER_NOTIFIED}` : `${TEXT.STATUS_UPDATED}. ${TEXT.USER_NOTIFY_FAILED}`,
      false
    );
  } catch (error) {
    await logError(env, "handleCallbackQuery", error, { callbackQuery });
    await answerCallbackQuery(env, callbackQuery.id, TEXT.UNKNOWN_ACTION, true);
  }
}

async function handleConfirmSubmit(callbackQuery, env) {
  const chatId = callbackQuery.message?.chat?.id;
  if (!chatId) return;

  const state = await getState(env, chatId);
  if (!state || state.step !== FORM_STEPS.WAITING_CONFIRMATION) {
    await answerCallbackQuery(env, callbackQuery.id, TEXT.NOTHING_TO_CONFIRM, true);
    return;
  }

  const payload = buildApplicationPayload(callbackQuery.from, normalizeDraft(state.data || {}), env);
  const dedupeKey = `request-submit:${chatId}:${hashPayload(payload)}`;
  const alreadySubmitted = await env.STATE_KV.get(dedupeKey);

  if (alreadySubmitted === "1") {
    await clearState(env, chatId);
    await answerCallbackQuery(env, callbackQuery.id, TEXT.SUCCESS, false);
    await sendMessage(env, chatId, TEXT.SUCCESS, { reply_markup: afterSubmitKeyboard(isAdmin(env, callbackQuery.from?.id)) });
    return;
  }

  await env.STATE_KV.put(dedupeKey, "1", { expirationTtl: 60 * 10 });
  await appendApplication(env, payload);
  await notifyAdmins(env, payload);
  await clearState(env, chatId);

  await answerCallbackQuery(env, callbackQuery.id, "Заявку підтверджено і збережено в Google Sheets.", false);
  await sendMessage(env, chatId, TEXT.SUCCESS, { reply_markup: afterSubmitKeyboard(isAdmin(env, callbackQuery.from?.id)) });
}

async function handleEditCallback(callbackQuery, env) {
  const chatId = callbackQuery.message?.chat?.id;
  const target = callbackQuery.data.split(":")[1];

  if (!chatId || !EDIT_FIELDS[target]) {
    await answerCallbackQuery(env, callbackQuery.id, TEXT.UNKNOWN_ACTION, true);
    return;
  }

  const state = await getState(env, chatId);
  if (!state || !state.data) {
    await answerCallbackQuery(env, callbackQuery.id, TEXT.NOTHING_TO_CONFIRM, true);
    return;
  }

  const nextStep = EDIT_FIELDS[target];
  const data = normalizeDraft(state.data);
  if (nextStep === FORM_STEPS.WAITING_PHOTOS) {
    data.photo_file_ids = [];
    data.photo_count = "0";
  }
  await saveState(env, chatId, newState(nextStep, data, true));
  await answerCallbackQuery(env, callbackQuery.id, "Редагування відкрито.", false);

  const askMap = {
    [FORM_STEPS.WAITING_NAME]: TEXT.START_APPLICATION,
    [FORM_STEPS.WAITING_PHONE]: TEXT.ASK_PHONE,
    [FORM_STEPS.WAITING_EMAIL]: TEXT.ASK_EMAIL,
    [FORM_STEPS.WAITING_ADDRESS]: TEXT.ASK_ADDRESS,
    [FORM_STEPS.WAITING_DESCRIPTION]: TEXT.ASK_DESCRIPTION,
    [FORM_STEPS.WAITING_CALL_TIME]: TEXT.ASK_CALL_TIME,
    [FORM_STEPS.WAITING_PHOTOS]: TEXT.ASK_PHOTOS
  };

  const keyboardMap = {
    [FORM_STEPS.WAITING_NAME]: cancelKeyboard(),
    [FORM_STEPS.WAITING_PHONE]: cancelKeyboard(),
    [FORM_STEPS.WAITING_EMAIL]: skipKeyboard(),
    [FORM_STEPS.WAITING_ADDRESS]: addressKeyboard(),
    [FORM_STEPS.WAITING_DESCRIPTION]: cancelKeyboard(),
    [FORM_STEPS.WAITING_CALL_TIME]: skipKeyboard(),
    [FORM_STEPS.WAITING_PHOTOS]: photoKeyboard(true)
  };

  await sendMessage(env, chatId, `${TEXT.EDIT_PROMPT}\n\n${askMap[nextStep]}`, {
    reply_markup: keyboardMap[nextStep]
  });
}

function newState(step, data = {}, isEditing = false) {
  return { step, data: normalizeDraft(data), isEditing };
}

function normalizeDraft(data) {
  const photoIds = Array.isArray(data.photo_file_ids)
    ? data.photo_file_ids.filter(Boolean)
    : typeof data.photo_file_ids === "string" && data.photo_file_ids !== SKIP_VALUE && data.photo_file_ids
      ? data.photo_file_ids.split(",").map((item) => item.trim()).filter(Boolean)
      : [];

  return {
    name: data.name || "",
    phone: data.phone || "",
    email: data.email || "",
    address: data.address || "",
    description: data.description || "",
    call_time: data.call_time || "",
    location_lat: data.location_lat || "",
    location_lon: data.location_lon || "",
    location_label: data.location_label || "",
    photo_file_ids: photoIds,
    photo_count: String(data.photo_count || photoIds.length || 0)
  };
}

function mapStatusCode(statusCode) {
  if (statusCode === "callback") return STATUS_CALLBACK;
  if (statusCode === "in_progress") return STATUS_IN_PROGRESS;
  if (statusCode === "done") return STATUS_DONE;
  if (statusCode === "cancelled") return STATUS_CANCELLED;
  return null;
}

function buildApplicationPayload(user, formData, env) {
  const createdAt = formatDateTime(env, new Date());
  const photoIds = Array.isArray(formData.photo_file_ids) ? formData.photo_file_ids.filter(Boolean) : [];

  return {
    request_id: generateRequestId(),
    created_at: createdAt,
    source: "telegram",
    telegram_id: String(user?.id || ""),
    username: user?.username ? `@${user.username}` : [user?.first_name, user?.last_name].filter(Boolean).join(" "),
    name: formData.name || "",
    phone: formData.phone || "",
    email: formData.email || SKIP_VALUE,
    address: formData.address || SKIP_VALUE,
    description: formData.description || "",
    call_time: formData.call_time || SKIP_VALUE,
    status: STATUS_NEW,
    updated_at: createdAt,
    last_admin_action: "",
    location_lat: formData.location_lat || "",
    location_lon: formData.location_lon || "",
    location_label: formData.location_label || "",
    photo_file_ids: photoIds.join(","),
    photo_count: String(photoIds.length)
  };
}

async function notifyAdmins(env, payload) {
  const adminIds = parseAdminIds(env.ADMIN_IDS);
  const text = formatAdminMessage(payload);

  for (const adminId of adminIds) {
    try {
      await sendMessage(env, adminId, text, {
        reply_markup: adminStatusKeyboard(payload.request_id)
      });
      if (payload.photo_file_ids) {
        const photoIds = String(payload.photo_file_ids).split(",").map((item) => item.trim()).filter(Boolean);
        for (const fileId of photoIds) {
          await sendPhoto(env, adminId, fileId, `📷 Фото до заявки <b>${escapeHtml(payload.request_id)}</b>`);
        }
      }
    } catch (error) {
      await logError(env, "notifyAdmins", error, { adminId, requestId: payload.request_id });
    }
  }
}

async function notifyClientAboutStatus(env, row, statusValue) {
  const chatId = Number(row?.telegram_id || 0);
  if (!chatId) return false;

  const text = [
    `<b>Оновлення по заявці ${escapeHtml(row.request_id || "")}</b>`,
    `📌 Новий статус: <b>${escapeHtml(statusValue)}</b>`
  ].join("\n");

  try {
    await sendMessage(env, chatId, text, { reply_markup: afterSubmitKeyboard(false) });
    return true;
  } catch (error) {
    await logError(env, "notifyClientAboutStatus", error, { chatId, requestId: row?.request_id, statusValue });
    return false;
  }
}

async function sendUserRequests(env, chatId, userId) {
  const rows = await getApplications(env);
  const userRows = rows.filter((row) => String(row.telegram_id || "") === String(chatId));

  if (!userRows.length) {
    await sendMessage(env, chatId, TEXT.NO_REQUESTS, { reply_markup: mainMenuKeyboard(isAdmin(env, userId)) });
    return;
  }

  await sendChunkedText(
    env,
    chatId,
    `<b>Ваші заявки:</b>\n\n${userRows.sort(compareRowsByCreatedAtDesc).slice(0, 15).map(formatRequestSummary).join("\n\n──────────\n\n")}`,
    mainMenuKeyboard(isAdmin(env, userId))
  );
}

async function sendAdminRequestsByStatus(env, chatId, status, emptyText) {
  const rows = await getApplications(env);
  const filtered = rows.filter((row) => String(row.status || "") === status).sort(compareRowsByCreatedAtDesc).slice(0, 15);

  if (!filtered.length) {
    await sendMessage(env, chatId, emptyText, { reply_markup: mainMenuKeyboard(true) });
    return;
  }

  const text = `<b>Заявки зі статусом ${escapeHtml(status)}:</b>\n\n${filtered.map(formatRequestSummary).join("\n\n──────────\n\n")}`;
  await sendChunkedText(env, chatId, text, mainMenuKeyboard(true));
}

async function sendTodayRequests(env, chatId) {
  const rows = await getApplications(env);
  const todayPrefix = formatDatePrefix(env, new Date());
  const filtered = rows
    .filter((row) => String(row.created_at || "").startsWith(todayPrefix))
    .sort(compareRowsByCreatedAtDesc)
    .slice(0, 15);

  if (!filtered.length) {
    await sendMessage(env, chatId, TEXT.TODAY_EMPTY, { reply_markup: mainMenuKeyboard(true) });
    return;
  }

  const text = `<b>Заявки за сьогодні (${filtered.length}):</b>\n\n${filtered.map(formatRequestSummary).join("\n\n──────────\n\n")}`;
  await sendChunkedText(env, chatId, text, mainMenuKeyboard(true));
}

async function sendStats(env, chatId) {
  const rows = await getApplications(env);
  if (!rows.length) {
    await sendMessage(env, chatId, TEXT.STATS_EMPTY, { reply_markup: mainMenuKeyboard(true) });
    return;
  }

  const counters = { total: rows.length, new: 0, callback: 0, inProgress: 0, done: 0, cancelled: 0, photos: 0, geos: 0 };
  for (const row of rows) {
    if (row.status === STATUS_NEW) counters.new += 1;
    if (row.status === STATUS_CALLBACK) counters.callback += 1;
    if (row.status === STATUS_IN_PROGRESS) counters.inProgress += 1;
    if (row.status === STATUS_DONE) counters.done += 1;
    if (row.status === STATUS_CANCELLED) counters.cancelled += 1;
    if (Number(row.photo_count || 0) > 0) counters.photos += 1;
    if (row.location_lat && row.location_lon) counters.geos += 1;
  }

  const todayPrefix = formatDatePrefix(env, new Date());
  const todayCount = rows.filter((row) => String(row.created_at || "").startsWith(todayPrefix)).length;

  const text = [
    "<b>Статистика заявок</b>",
    `Усього: <b>${counters.total}</b>`,
    `Сьогодні: <b>${todayCount}</b>`,
    `Нова: <b>${counters.new}</b>`,
    `Передзвонити: <b>${counters.callback}</b>`,
    `В роботі: <b>${counters.inProgress}</b>`,
    `Виконана: <b>${counters.done}</b>`,
    `Скасована: <b>${counters.cancelled}</b>`,
    `З фото: <b>${counters.photos}</b>`,
    `З геолокацією: <b>${counters.geos}</b>`
  ].join("\n");

  await sendMessage(env, chatId, text, { reply_markup: mainMenuKeyboard(true) });
}

async function findRequestsByPhone(env, chatId, commandText) {
  const phone = normalizePhone(commandText.replace("/find_phone", "").trim());
  if (!phone) {
    await sendMessage(env, chatId, TEXT.FIND_PHONE_USAGE, { reply_markup: mainMenuKeyboard(true) });
    return;
  }

  const rows = await getApplications(env);
  const matched = rows.filter((row) => normalizePhone(row.phone || "") === phone).sort(compareRowsByCreatedAtDesc).slice(0, 15);
  if (!matched.length) {
    await sendMessage(env, chatId, TEXT.FIND_PHONE_EMPTY, { reply_markup: mainMenuKeyboard(true) });
    return;
  }

  const text = `<b>Заявки за номером ${escapeHtml(phone)}:</b>\n\n${matched.map(formatRequestSummary).join("\n\n──────────\n\n")}`;
  await sendChunkedText(env, chatId, text, mainMenuKeyboard(true));
}

function formatAdminMessage(payload) {
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
    `📌 <b>Статус:</b> ${escapeHtml(payload.status)}`,
    `📷 <b>Фото:</b> ${escapeHtml(payload.photo_count || "0")}`
  ];

  if (payload.location_lat && payload.location_lon) {
    lines.push(`🗺 <b>Геолокація:</b> ${escapeHtml(payload.location_label || `${payload.location_lat},${payload.location_lon}`)}`);
  }
  if (payload.updated_at) {
    lines.push(`🛠 <b>Оновлено:</b> ${escapeHtml(payload.updated_at)}`);
  }
  if (payload.last_admin_action) {
    lines.push(`👨‍🔧 <b>Остання дія:</b> ${escapeHtml(payload.last_admin_action)}`);
  }

  return lines.join("\n");
}

function formatClientPreview(data) {
  const photoCount = Array.isArray(data.photo_file_ids) ? data.photo_file_ids.length : Number(data.photo_count || 0);
  const lines = [
    `👤 <b>Ім'я:</b> ${escapeHtml(data.name || "—")}`,
    `📞 <b>Телефон:</b> ${escapeHtml(data.phone || "—")}`,
    `✉️ <b>Email:</b> ${escapeHtml(data.email || SKIP_VALUE)}`,
    `📍 <b>Адреса:</b> ${escapeHtml(data.address || "—")}`,
    `📝 <b>Опис:</b> ${escapeHtml(data.description || "—")}`,
    `🕒 <b>Зручний час:</b> ${escapeHtml(data.call_time || SKIP_VALUE)}`,
    `📷 <b>Фото:</b> ${escapeHtml(String(photoCount || 0))}`
  ];
  if (data.location_lat && data.location_lon) {
    lines.push(`🗺 <b>Геолокація:</b> ${escapeHtml(data.location_label || `${data.location_lat},${data.location_lon}`)}`);
  }
  return lines.join("\n");
}

function formatRequestSummary(row) {
  const parts = [
    `🆔 <b>${escapeHtml(row.request_id || "-")}</b>`,
    `📅 ${escapeHtml(row.created_at || "-")}`,
    `👤 ${escapeHtml(row.name || "-")}`,
    `📞 ${escapeHtml(row.phone || "-")}`,
    `📍 ${escapeHtml(row.address || "-")}`,
    `📝 ${escapeHtml(trimText(row.description || "-", 180))}`,
    `📌 Статус: <b>${escapeHtml(row.status || "-")}</b>`
  ];
  if (Number(row.photo_count || 0) > 0) parts.push(`📷 Фото: <b>${escapeHtml(String(row.photo_count || 0))}</b>`);
  if (row.location_lat && row.location_lon) parts.push(`🗺 Гео: <b>так</b>`);
  return parts.join("\n");
}

function replaceStatusInText(text, newStatus, updatedAt, adminUser) {
  const lines = text.split("\n");
  let statusReplaced = false;
  let updatedReplaced = false;
  let adminReplaced = false;
  const adminLabel = formatAdminLabel(adminUser);

  const updated = lines.map((line) => {
    if (line.startsWith("📌 <b>Статус:</b>")) {
      statusReplaced = true;
      return `📌 <b>Статус:</b> ${escapeHtml(newStatus)}`;
    }
    if (line.startsWith("🛠 <b>Оновлено:</b>")) {
      updatedReplaced = true;
      return `🛠 <b>Оновлено:</b> ${escapeHtml(updatedAt || "—")}`;
    }
    if (line.startsWith("👨‍🔧 <b>Остання дія:</b>")) {
      adminReplaced = true;
      return `👨‍🔧 <b>Остання дія:</b> ${escapeHtml(adminLabel)}`;
    }
    return line;
  });

  if (!statusReplaced) updated.push(`📌 <b>Статус:</b> ${escapeHtml(newStatus)}`);
  if (!updatedReplaced) updated.push(`🛠 <b>Оновлено:</b> ${escapeHtml(updatedAt || "—")}`);
  if (!adminReplaced) updated.push(`👨‍🔧 <b>Остання дія:</b> ${escapeHtml(adminLabel)}`);
  return updated.join("\n");
}

function mainMenuKeyboard(isAdminUser = false) {
  const keyboard = [[{ text: "Створити заявку" }], [{ text: "Мої заявки" }]];
  if (isAdminUser) {
    keyboard.push([{ text: "/new_requests" }, { text: "/today" }]);
    keyboard.push([{ text: "/stats" }]);
  }
  return { keyboard, resize_keyboard: true };
}

function cancelKeyboard() {
  return { keyboard: [[{ text: "/cancel" }]], resize_keyboard: true };
}

function skipKeyboard() {
  return { keyboard: [[{ text: "Пропустити" }], [{ text: "/cancel" }]], resize_keyboard: true };
}

function addressKeyboard() {
  return {
    keyboard: [
      [{ text: "📍 Надіслати геолокацію", request_location: true }],
      [{ text: "Продовжити без геолокації" }],
      [{ text: "/cancel" }]
    ],
    resize_keyboard: true,
    one_time_keyboard: true
  };
}

function photoKeyboard(includeSkip) {
  const rows = [[{ text: "Готово з фото" }]];
  if (includeSkip) rows.push([{ text: "Пропустити" }]);
  rows.push([{ text: "/cancel" }]);
  return { keyboard: rows, resize_keyboard: true };
}

function afterSubmitKeyboard(isAdminUser = false) {
  const keyboard = [[{ text: "Створити заявку" }], [{ text: "Мої заявки" }], [{ text: "Головне меню" }]];
  if (isAdminUser) keyboard.push([{ text: "/new_requests" }, { text: "/today" }]);
  return { keyboard, resize_keyboard: true };
}

function adminStatusKeyboard(requestId) {
  return {
    inline_keyboard: [
      [
        { text: "📞 Передзвонити", callback_data: `status:${requestId}:callback` },
        { text: "✅ В роботу", callback_data: `status:${requestId}:in_progress` }
      ],
      [
        { text: "🏁 Виконана", callback_data: `status:${requestId}:done` },
        { text: "⛔ Скасована", callback_data: `status:${requestId}:cancelled` }
      ]
    ]
  };
}

function confirmKeyboard() {
  return {
    inline_keyboard: [
      [{ text: "✅ Підтвердити заявку", callback_data: CALLBACKS.CONFIRM_SUBMIT }],
      [
        { text: "✏️ Ім'я", callback_data: `${CALLBACKS.EDIT_PREFIX}name` },
        { text: "✏️ Телефон", callback_data: `${CALLBACKS.EDIT_PREFIX}phone` }
      ],
      [
        { text: "✏️ Email", callback_data: `${CALLBACKS.EDIT_PREFIX}email` },
        { text: "✏️ Адреса", callback_data: `${CALLBACKS.EDIT_PREFIX}address` }
      ],
      [
        { text: "✏️ Опис", callback_data: `${CALLBACKS.EDIT_PREFIX}description` },
        { text: "✏️ Час", callback_data: `${CALLBACKS.EDIT_PREFIX}call_time` }
      ],
      [{ text: "✏️ Фото", callback_data: `${CALLBACKS.EDIT_PREFIX}photos` }],
      [{ text: "❌ Скасувати", callback_data: CALLBACKS.CONFIRM_CANCEL }]
    ]
  };
}

async function sendChunkedText(env, chatId, text, replyMarkup) {
  const chunks = splitLongText(text, 3800);
  for (let i = 0; i < chunks.length; i++) {
    await sendMessage(env, chatId, chunks[i], i === chunks.length - 1 ? { reply_markup: replyMarkup } : {});
  }
}

async function sendMessage(env, chatId, text, extra = {}) {
  return await telegramApi(env, "sendMessage", {
    chat_id: chatId,
    text,
    parse_mode: "HTML",
    disable_web_page_preview: true,
    ...extra
  });
}

async function sendPhoto(env, chatId, fileId, caption = "") {
  return await telegramApi(env, "sendPhoto", {
    chat_id: chatId,
    photo: fileId,
    caption,
    parse_mode: "HTML"
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
  await env.STATE_KV.put(`state:${chatId}`, JSON.stringify(state), { expirationTtl: 60 * 60 * 24 });
}

async function clearState(env, chatId) {
  await env.STATE_KV.delete(`state:${chatId}`);
}

async function checkRateLimit(env, chatId, userId) {
  const key = `rate:${chatId}:${userId || "0"}`;
  const raw = await env.STATE_KV.get(key);
  const now = Date.now();
  const bucket = raw ? JSON.parse(raw) : { count: 0, start: now };

  if (now - bucket.start > 15_000) {
    bucket.count = 0;
    bucket.start = now;
  }

  bucket.count += 1;
  await env.STATE_KV.put(key, JSON.stringify(bucket), { expirationTtl: 60 });
  return bucket.count <= 15;
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
    payload.call_time,
    payload.photo_file_ids
  ].join("|");

  let hash = 0;
  for (let i = 0; i < base.length; i++) {
    hash = (hash * 31 + base.charCodeAt(i)) >>> 0;
  }
  return String(hash);
}

function validateEnv(env) {
  const required = ["BOT_TOKEN", "GOOGLE_SHEET_ID", "GOOGLE_SERVICE_ACCOUNT_JSON", "ADMIN_IDS", "STATE_KV"];
  for (const key of required) {
    if (!env[key]) {
      throw new Error(`Missing environment variable: ${key}`);
    }
  }
}

async function appendApplication(env, application) {
  const context = await ensureWorksheet(env);
  const row = context.headers.map((header) => application[header] ?? "");
  await withGoogleRetry(() => sheetsAppend(env, `${getWorksheetName(env)}!A:${indexToColumn(context.headers.length - 1)}`, [row]));
}

async function getApplications(env) {
  const context = await ensureWorksheet(env);
  const values = await sheetsGetValues(env, `${getWorksheetName(env)}!A:${indexToColumn(context.headers.length - 1)}`);
  if (!values.length) return [];

  const headers = values[0];
  return values.slice(1).map((row) => {
    const item = {};
    headers.forEach((header, index) => {
      item[header] = row[index] ?? "";
    });
    return item;
  });
}

async function updateApplicationStatus(env, requestId, newStatus, adminUser) {
  const context = await ensureWorksheet(env);
  const headers = context.headers;
  const values = await sheetsGetValues(env, `${getWorksheetName(env)}!A:${indexToColumn(headers.length - 1)}`);
  if (values.length <= 1) return { ok: false };

  const requestIdIndex = headers.indexOf("request_id");
  const statusIndex = headers.indexOf("status");
  const updatedAtIndex = headers.indexOf("updated_at");
  const adminActionIndex = headers.indexOf("last_admin_action");
  if (requestIdIndex === -1 || statusIndex === -1) return { ok: false };

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if ((row[requestIdIndex] || "") !== requestId) continue;

    const rowNumber = i + 1;
    const updatedAt = formatDateTime(env, new Date());
    const rowObject = {};
    headers.forEach((header, index) => {
      rowObject[header] = row[index] ?? "";
    });
    rowObject.status = newStatus;
    rowObject.updated_at = updatedAt;
    rowObject.last_admin_action = formatAdminLabel(adminUser);

    const updates = [
      withGoogleRetry(() => sheetsUpdateValues(env, `${getWorksheetName(env)}!${indexToColumn(statusIndex)}${rowNumber}`, [[newStatus]]))
    ];
    if (updatedAtIndex !== -1) {
      updates.push(withGoogleRetry(() => sheetsUpdateValues(env, `${getWorksheetName(env)}!${indexToColumn(updatedAtIndex)}${rowNumber}`, [[updatedAt]])));
    }
    if (adminActionIndex !== -1) {
      updates.push(withGoogleRetry(() => sheetsUpdateValues(env, `${getWorksheetName(env)}!${indexToColumn(adminActionIndex)}${rowNumber}`, [[formatAdminLabel(adminUser)]])));
    }
    await Promise.all(updates);
    return { ok: true, row: rowObject, updatedAt };
  }

  return { ok: false };
}

async function ensureWorksheet(env) {
  const metadata = await sheetsGetMetadata(env);
  const sheetTitle = getWorksheetName(env);
  let found = metadata.sheets?.find((sheet) => sheet.properties?.title === sheetTitle);

  if (!found) {
    await sheetsBatchUpdate(env, { requests: [{ addSheet: { properties: { title: sheetTitle } } }] });
    const refreshed = await sheetsGetMetadata(env);
    found = refreshed.sheets?.find((sheet) => sheet.properties?.title === sheetTitle);
  }

  const totalHeaders = [...SHEET_HEADERS, ...OPTIONAL_EXTRA_HEADERS];
  const currentHeaders = await sheetsGetValues(env, `${sheetTitle}!A1:${indexToColumn(totalHeaders.length - 1)}1`);
  const firstRow = currentHeaders[0] || [];

  if (!firstRow.length) {
    await sheetsUpdateValues(env, `${sheetTitle}!A1:${indexToColumn(totalHeaders.length - 1)}1`, [totalHeaders]);
    return { headers: totalHeaders, sheetId: found?.properties?.sheetId };
  }

  const mergedHeaders = firstRow.slice();
  for (const header of totalHeaders) {
    if (!mergedHeaders.includes(header)) {
      mergedHeaders.push(header);
    }
  }

  if (mergedHeaders.length !== firstRow.length) {
    await sheetsUpdateValues(env, `${sheetTitle}!A1:${indexToColumn(mergedHeaders.length - 1)}1`, [mergedHeaders]);
  }

  return { headers: mergedHeaders, sheetId: found?.properties?.sheetId };
}

function getWorksheetName(env) {
  return env.WORKSHEET_NAME || DEFAULT_WORKSHEET_NAME;
}

function getLogsWorksheetName(env) {
  return env.LOGS_WORKSHEET_NAME || DEFAULT_LOGS_WORKSHEET_NAME;
}

async function ensureLogsWorksheet(env) {
  const metadata = await sheetsGetMetadata(env);
  const sheetTitle = getLogsWorksheetName(env);
  const found = metadata.sheets?.find((sheet) => sheet.properties?.title === sheetTitle);
  if (!found) {
    await sheetsBatchUpdate(env, { requests: [{ addSheet: { properties: { title: sheetTitle } } }] });
  }

  const values = await sheetsGetValues(env, `${sheetTitle}!A1:E1`);
  if (!values.length || !values[0] || !values[0].length) {
    await sheetsUpdateValues(env, `${sheetTitle}!A1:E1`, [LOG_HEADERS]);
  }
}

async function sheetsGetMetadata(env) {
  return await googleApi(env, `https://sheets.googleapis.com/v4/spreadsheets/${env.GOOGLE_SHEET_ID}`);
}

async function sheetsBatchUpdate(env, body) {
  return await googleApi(env, `https://sheets.googleapis.com/v4/spreadsheets/${env.GOOGLE_SHEET_ID}:batchUpdate`, {
    method: "POST",
    body
  });
}

async function sheetsGetValues(env, range) {
  const data = await googleApi(env, `https://sheets.googleapis.com/v4/spreadsheets/${env.GOOGLE_SHEET_ID}/values/${encodeURIComponent(range)}`);
  return data.values || [];
}

async function sheetsUpdateValues(env, range, values) {
  return await googleApi(
    env,
    `https://sheets.googleapis.com/v4/spreadsheets/${env.GOOGLE_SHEET_ID}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`,
    { method: "PUT", body: { range, majorDimension: "ROWS", values } }
  );
}

async function sheetsAppend(env, range, values) {
  return await googleApi(
    env,
    `https://sheets.googleapis.com/v4/spreadsheets/${env.GOOGLE_SHEET_ID}/values/${encodeURIComponent(range)}:append?valueInputOption=USER_ENTERED`,
    { method: "POST", body: { range, majorDimension: "ROWS", values } }
  );
}


async function withGoogleRetry(fn, retries = 2) {
  let lastError;
  for (let attempt = 0; attempt <= retries; attempt++) {
    try {
      return await fn();
    } catch (error) {
      lastError = error;
      if (attempt === retries) break;
      await new Promise((resolve) => setTimeout(resolve, 250 * (attempt + 1)));
    }
  }
  throw lastError;
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
  if (cachedGoogleToken && cachedGoogleToken.expiresAt > Date.now() + 60_000) {
    return cachedGoogleToken.token;
  }

  const serviceAccount = JSON.parse(env.GOOGLE_SERVICE_ACCOUNT_JSON);
  const now = Math.floor(Date.now() / 1000);
  const jwtHeader = { alg: "RS256", typ: "JWT" };
  const jwtClaimSet = {
    iss: serviceAccount.client_email,
    scope: "https://www.googleapis.com/auth/spreadsheets",
    aud: serviceAccount.token_uri,
    exp: now + 3600,
    iat: now
  };

  const unsignedToken = `${base64UrlEncode(JSON.stringify(jwtHeader))}.${base64UrlEncode(JSON.stringify(jwtClaimSet))}`;
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

async function signJwt(unsignedToken, pemPrivateKey) {
  const keyData = pemToArrayBuffer(pemPrivateKey);
  const cryptoKey = await crypto.subtle.importKey(
    "pkcs8",
    keyData,
    { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" },
    false,
    ["sign"]
  );

  const signature = await crypto.subtle.sign("RSASSA-PKCS1-v1_5", cryptoKey, new TextEncoder().encode(unsignedToken));
  return arrayBufferToBase64Url(signature);
}

function pemToArrayBuffer(pem) {
  const base64 = pem.replace("-----BEGIN PRIVATE KEY-----", "").replace("-----END PRIVATE KEY-----", "").replace(/\s+/g, "");
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}

function base64UrlEncode(value) {
  const bytes = new TextEncoder().encode(value);
  return arrayBufferToBase64Url(bytes);
}

function arrayBufferToBase64Url(buffer) {
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
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
    headers: { "content-type": "application/json; charset=utf-8" }
  });
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function trimText(value, maxLength) {
  const text = String(value || "");
  return text.length > maxLength ? `${text.slice(0, maxLength - 1)}…` : text;
}

function splitLongText(text, maxLength) {
  if (text.length <= maxLength) return [text];
  const parts = [];
  let rest = text;
  while (rest.length > maxLength) {
    let splitAt = rest.lastIndexOf("\n\n", maxLength);
    if (splitAt < 1000) splitAt = rest.lastIndexOf("\n", maxLength);
    if (splitAt < 500) splitAt = maxLength;
    parts.push(rest.slice(0, splitAt));
    rest = rest.slice(splitAt).trimStart();
  }
  if (rest) parts.push(rest);
  return parts;
}

function compareRowsByCreatedAtDesc(a, b) {
  return parseDateTimeValue(b.created_at) - parseDateTimeValue(a.created_at);
}

function parseDateTimeValue(value) {
  const text = String(value || "").trim();
  const match = text.match(/^(\d{2})\.(\d{2})\.(\d{4}),?\s+(\d{2}):(\d{2}):(\d{2})$/);
  if (match) {
    const [, dd, mm, yyyy, hh, mi, ss] = match;
    return new Date(`${yyyy}-${mm}-${dd}T${hh}:${mi}:${ss}Z`).getTime();
  }
  return 0;
}

function formatDateTime(env, date) {
  const timezone = env.TIMEZONE || DEFAULT_TIMEZONE;
  return new Intl.DateTimeFormat("uk-UA", {
    timeZone: timezone,
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false
  }).format(date);
}

function formatDatePrefix(env, date) {
  const timezone = env.TIMEZONE || DEFAULT_TIMEZONE;
  return new Intl.DateTimeFormat("uk-UA", {
    timeZone: timezone,
    day: "2-digit",
    month: "2-digit",
    year: "numeric"
  }).format(date);
}

function formatAdminLabel(adminUser) {
  if (!adminUser) return "Адміністратор";
  if (adminUser.username) return `@${adminUser.username}`;
  const fullName = [adminUser.first_name, adminUser.last_name].filter(Boolean).join(" ").trim();
  return fullName || String(adminUser.id || "Адміністратор");
}

async function logError(env, scope, error, details = null) {
  try {
    console.error(scope, error instanceof Error ? error.message : String(error), details || "");
    if (!env.GOOGLE_SHEET_ID || !env.GOOGLE_SERVICE_ACCOUNT_JSON) return;
    await ensureLogsWorksheet(env);
    const row = [
      formatDateTime(env, new Date()),
      "error",
      scope,
      error instanceof Error ? error.message : String(error),
      details ? trimText(JSON.stringify(details), 3000) : ""
    ];
    await sheetsAppend(env, `${getLogsWorksheetName(env)}!A:E`, [row]);
  } catch (nestedError) {
    console.error("logError failed", nestedError instanceof Error ? nestedError.message : String(nestedError));
  }
}
