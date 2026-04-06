const DEFAULT_TIMEZONE = "Europe/Kyiv";
const DEFAULT_BOOKINGS_SHEET = "Bookings";
const DEFAULT_USERS_SHEET = "Users";
const DEFAULT_LOGS_SHEET = "BotLogs";
const DEFAULT_SLOT_MINUTES = 60;
const DEFAULT_WORK_START_HOUR = 9;
const DEFAULT_WORK_END_HOUR = 18;
const BOOKING_LOOKAHEAD_DAYS = 14;
const MAX_COMMENT_LENGTH = 500;

const STATUS_PENDING = "Очікує підтвердження";
const STATUS_APPROVED = "Підтверджено";
const STATUS_REJECTED = "Відхилено";
const STATUS_CANCELLED = "Скасовано";

const USER_HEADERS = [
  "telegram_id",
  "username",
  "full_name",
  "phone",
  "registered_at",
  "last_login_at",
  "is_registered"
];

const BOOKING_HEADERS = [
  "booking_id",
  "created_at",
  "updated_at",
  "telegram_id",
  "username",
  "full_name",
  "phone",
  "booking_date",
  "booking_time",
  "slot_key",
  "comment",
  "status",
  "admin_action_by",
  "approved_at",
  "rejected_at",
  "cancelled_at",
  "cancelled_by"
];

const LOG_HEADERS = ["created_at", "level", "scope", "message", "details"];

const FORM_STEPS = {
  WAITING_PASSWORD: "waiting_password",
  WAITING_REG_NAME: "waiting_reg_name",
  WAITING_REG_PHONE: "waiting_reg_phone",
  WAITING_BOOKING_DATE: "waiting_booking_date",
  WAITING_BOOKING_TIME: "waiting_booking_time",
  WAITING_BOOKING_COMMENT: "waiting_booking_comment",
  WAITING_BOOKING_CONFIRM: "waiting_booking_confirm"
};

const CALLBACKS = {
  DATE_PREFIX: "date:",
  TIME_PREFIX: "time:",
  CONFIRM_BOOKING: "booking:confirm",
  CANCEL_BOOKING_FLOW: "booking:cancel_flow",
  MY_BOOKINGS: "menu:my_bookings",
  START_BOOKING: "menu:start_booking",
  CANCEL_PREFIX: "cancel:",
  ADMIN_APPROVE_PREFIX: "admin:approve:",
  ADMIN_REJECT_PREFIX: "admin:reject:",
  ADMIN_PENDING: "admin:list:pending",
  ADMIN_TODAY: "admin:list:today",
  BACK_MAIN: "menu:main"
};

const TEXT = {
  START_GUEST: "Вітаю. Для доступу до бота введіть пароль.",
  PASSWORD_INVALID: "Невірний пароль. Спробуйте ще раз.",
  ASK_REG_NAME: "Доступ дозволено. Введіть ваше ім'я та прізвище.",
  INVALID_NAME: "Вкажіть ім'я щонайменше з 2 символів.",
  ASK_REG_PHONE: "Введіть номер телефону у форматі <b>380XXXXXXXXX</b>.",
  INVALID_PHONE: "Номер має бути у форматі <b>380XXXXXXXXX</b>.",
  REGISTRATION_DONE: "Реєстрацію завершено. Тепер ви можете записатися на стрижку.",
  START_REGISTERED: "Вітаю. Оберіть дію в меню нижче.",
  NOT_REGISTERED: "Спочатку потрібно пройти реєстрацію через пароль.",
  START_BOOKING: "Оберіть дату для запису.",
  NO_FREE_SLOTS: "На цю дату вільних слотів немає. Оберіть іншу дату.",
  ASK_TIME: "Оберіть вільний час.",
  ASK_COMMENT: "За потреби напишіть коментар до запису або натисніть <b>Пропустити</b>.",
  COMMENT_TOO_LONG: "Коментар задовгий. До 500 символів.",
  NOTHING_TO_CONFIRM: "Немає бронювання для підтвердження.",
  BOOKING_SAVED: "Заявку на запис створено. Вона з'явиться в активних лише після підтвердження адміністратором.",
  BOOKING_CANCELLED_FLOW: "Створення запису скасовано.",
  NO_BOOKINGS: "У вас поки немає заявок.",
  BOOKING_CANCELLED: "Бронювання скасовано.",
  BOOKING_CANCEL_DENIED: "Це бронювання вже не можна скасувати.",
  ADMIN_ONLY: "Ця дія доступна лише адміну.",
  STATUS_UPDATED: "Статус заявки оновлено.",
  UNKNOWN_ACTION: "Не вдалося обробити дію.",
  FALLBACK: "Оберіть одну з доступних дій у меню.",
  FLOOD_LIMIT: "Забагато дій за короткий час. Спробуйте трохи пізніше.",
  ACTIVE_FLOW_EXISTS: "У вас вже є незавершена дія. Завершіть її або натисніть /cancel.",
  TODAY_EMPTY: "На сьогодні записів немає.",
  PENDING_EMPTY: "Немає заявок, що очікують підтвердження.",
  STATS_EMPTY: "Ще немає даних для статистики.",
  CANCEL: "Поточну дію скасовано.",
  INVALID_DATE: "Не вдалося визначити дату. Спробуйте ще раз.",
  INVALID_TIME: "Не вдалося визначити час. Спробуйте ще раз.",
  SLOT_ALREADY_TAKEN: "Цей слот уже зайнятий або очікує підтвердження. Оберіть інший час.",
  SKIP: "Пропустити",
  MY_BOOKINGS_TITLE: "<b>Мої заявки</b>",
  ADMIN_PENDING_TITLE: "<b>Нові заявки</b>",
  ADMIN_TODAY_TITLE: "<b>Записи на сьогодні</b>",
  STATS_TITLE: "<b>Статистика записів</b>",
  FIND_PHONE_USAGE: "Використання: <b>/find_phone 380XXXXXXXXX</b>",
  FIND_PHONE_EMPTY: "За цим номером нічого не знайдено."
};

export default {
  async fetch(request, env, ctx) {
    try {
      const url = new URL(request.url);

      if (request.method === "GET" && url.pathname === "/") {
        return jsonResponse({ ok: true, service: "hair-booking-bot", status: "running" });
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
  await ensureAllWorksheets(env);

  const origin = `${url.protocol}//${url.host}`;
  const webhookUrl = `${origin}/webhook`;

  const webhookResult = await telegramApi(env, "setWebhook", {
    url: webhookUrl,
    secret_token: env.TELEGRAM_WEBHOOK_SECRET,
    allowed_updates: ["message", "callback_query"]
  });

  const commandsResult = await telegramApi(env, "setMyCommands", {
    commands: [
      { command: "start", description: "Запустити бота" },
      { command: "book", description: "Створити заявку на запис" },
      { command: "my_bookings", description: "Мої заявки" },
      { command: "pending", description: "Нові заявки (адмін)" },
      { command: "today", description: "Список на сьогодні (адмін)" },
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

async function handleMessage(message, env) {
  const chatId = message.chat?.id;
  const userId = message.from?.id;
  const text = (message.text || "").trim();

  if (!chatId) return;

  if (!(await checkRateLimit(env, chatId, userId))) {
    await sendMessage(env, chatId, TEXT.FLOOD_LIMIT, { reply_markup: mainMenuKeyboard(false, false) });
    return;
  }

  const registration = await getUserRegistration(env, userId);
  const admin = isAdmin(env, userId);
  const registered = Boolean(registration?.is_registered === "true");

  if (text === "/cancel") {
    await clearState(env, chatId);
    await sendMessage(env, chatId, TEXT.CANCEL, { reply_markup: mainMenuKeyboard(registered, admin) });
    return;
  }

  if (text === "/start" || text === "Головне меню") {
    await clearState(env, chatId);
    if (!registered) {
      await saveState(env, chatId, { step: FORM_STEPS.WAITING_PASSWORD, data: {} });
      await sendMessage(env, chatId, TEXT.START_GUEST, { reply_markup: cancelKeyboard() });
      return;
    }
    await touchUserLastLogin(env, userId);
    await sendMessage(env, chatId, TEXT.START_REGISTERED, { reply_markup: mainMenuKeyboard(true, admin) });
    return;
  }

  if (!registered) {
    const state = await getState(env, chatId);
    if (!state) {
      await saveState(env, chatId, { step: FORM_STEPS.WAITING_PASSWORD, data: {} });
      await sendMessage(env, chatId, TEXT.START_GUEST, { reply_markup: cancelKeyboard() });
      return;
    }
    await processGuestFlow(message, state, env);
    return;
  }

  if (text === "/book" || text === "Записатися") {
    const state = await getState(env, chatId);
    if (state) {
      await sendMessage(env, chatId, TEXT.ACTIVE_FLOW_EXISTS, { reply_markup: cancelKeyboard() });
      return;
    }
    await startBookingFlow(env, chatId, userId, admin);
    return;
  }

  if (text === "/my_bookings" || text === "Мої заявки") {
    await clearState(env, chatId);
    await sendUserBookings(env, chatId, userId);
    return;
  }

  if (text === "/pending") {
    if (!admin) {
      await sendMessage(env, chatId, TEXT.ADMIN_ONLY, { reply_markup: mainMenuKeyboard(true, false) });
      return;
    }
    await clearState(env, chatId);
    await sendPendingBookings(env, chatId);
    return;
  }

  if (text === "/today") {
    if (!admin) {
      await sendMessage(env, chatId, TEXT.ADMIN_ONLY, { reply_markup: mainMenuKeyboard(true, false) });
      return;
    }
    await clearState(env, chatId);
    await sendTodayBookings(env, chatId);
    return;
  }

  if (text === "/stats") {
    if (!admin) {
      await sendMessage(env, chatId, TEXT.ADMIN_ONLY, { reply_markup: mainMenuKeyboard(true, false) });
      return;
    }
    await clearState(env, chatId);
    await sendStats(env, chatId);
    return;
  }

  if (text.startsWith("/find_phone")) {
    if (!admin) {
      await sendMessage(env, chatId, TEXT.ADMIN_ONLY, { reply_markup: mainMenuKeyboard(true, false) });
      return;
    }
    await clearState(env, chatId);
    await findBookingsByPhone(env, chatId, text);
    return;
  }

  const state = await getState(env, chatId);
  if (!state) {
    await sendMessage(env, chatId, TEXT.FALLBACK, { reply_markup: mainMenuKeyboard(true, admin) });
    return;
  }

  await processRegisteredFlow(message, state, env, registration);
}

async function processGuestFlow(message, state, env) {
  const chatId = message.chat.id;
  const userId = message.from?.id;
  const text = (message.text || "").trim();

  switch (state.step) {
    case FORM_STEPS.WAITING_PASSWORD: {
      if (text !== String(env.ACCESS_PASSWORD || "")) {
        await sendMessage(env, chatId, TEXT.PASSWORD_INVALID, { reply_markup: cancelKeyboard() });
        return;
      }
      await saveState(env, chatId, { step: FORM_STEPS.WAITING_REG_NAME, data: {} });
      await sendMessage(env, chatId, TEXT.ASK_REG_NAME, { reply_markup: cancelKeyboard() });
      return;
    }

    case FORM_STEPS.WAITING_REG_NAME: {
      if (!validateName(text)) {
        await sendMessage(env, chatId, TEXT.INVALID_NAME, { reply_markup: cancelKeyboard() });
        return;
      }
      await saveState(env, chatId, { step: FORM_STEPS.WAITING_REG_PHONE, data: { full_name: text } });
      await sendMessage(env, chatId, TEXT.ASK_REG_PHONE, { reply_markup: cancelKeyboard() });
      return;
    }

    case FORM_STEPS.WAITING_REG_PHONE: {
      const phone = normalizePhone(text);
      if (!validatePhone(phone)) {
        await sendMessage(env, chatId, TEXT.INVALID_PHONE, { reply_markup: cancelKeyboard() });
        return;
      }
      const draft = state.data || {};
      await upsertUser(env, {
        telegram_id: String(userId || chatId),
        username: getTelegramUsername(message.from),
        full_name: draft.full_name || "",
        phone,
        registered_at: formatDateTime(env, new Date()),
        last_login_at: formatDateTime(env, new Date()),
        is_registered: "true"
      });
      await clearState(env, chatId);
      await sendMessage(env, chatId, TEXT.REGISTRATION_DONE, { reply_markup: mainMenuKeyboard(true, isAdmin(env, userId)) });
      return;
    }

    default:
      await clearState(env, chatId);
      await sendMessage(env, chatId, TEXT.START_GUEST, { reply_markup: cancelKeyboard() });
  }
}

async function processRegisteredFlow(message, state, env, registration) {
  const chatId = message.chat.id;
  const text = (message.text || "").trim();
  const data = state.data || {};

  switch (state.step) {
    case FORM_STEPS.WAITING_BOOKING_COMMENT: {
      const comment = text === TEXT.SKIP ? "" : text;
      if (comment.length > MAX_COMMENT_LENGTH) {
        await sendMessage(env, chatId, TEXT.COMMENT_TOO_LONG, { reply_markup: skipKeyboard() });
        return;
      }
      data.comment = comment;
      await sendBookingConfirmation(env, chatId, registration, data);
      return;
    }

    default:
      await sendMessage(env, chatId, TEXT.FALLBACK, { reply_markup: mainMenuKeyboard(true, isAdmin(env, message.from?.id)) });
  }
}

async function startBookingFlow(env, chatId, userId, admin) {
  const dates = getUpcomingDates(env, BOOKING_LOOKAHEAD_DAYS);
  await saveState(env, chatId, { step: FORM_STEPS.WAITING_BOOKING_DATE, data: {} });
  await sendMessage(env, chatId, TEXT.START_BOOKING, {
    reply_markup: bookingDatesKeyboard(dates, admin)
  });
}

async function sendBookingConfirmation(env, chatId, registration, data) {
  await saveState(env, chatId, { step: FORM_STEPS.WAITING_BOOKING_CONFIRM, data });
  const preview = [
    "<b>Підтвердіть заявку на запис</b>",
    `👤 <b>Ім'я:</b> ${escapeHtml(registration?.full_name || "—")}`,
    `📞 <b>Телефон:</b> ${escapeHtml(registration?.phone || "—")}`,
    `📅 <b>Дата:</b> ${escapeHtml(formatHumanDate(data.booking_date))}`,
    `🕒 <b>Час:</b> ${escapeHtml(data.booking_time || "—")}`,
    `💬 <b>Коментар:</b> ${escapeHtml(data.comment || "—")}`,
    `📌 <b>Статус після створення:</b> ${STATUS_PENDING}`
  ].join("\n");

  await sendMessage(env, chatId, preview, { reply_markup: bookingConfirmKeyboard() });
}

async function handleCallbackQuery(callbackQuery, env) {
  const data = callbackQuery.data || "";
  const message = callbackQuery.message;
  const chatId = message?.chat?.id;
  const userId = callbackQuery.from?.id;
  const admin = isAdmin(env, userId);

  if (!chatId) return;

  try {
    if (data === CALLBACKS.START_BOOKING) {
      const registration = await getUserRegistration(env, userId);
      if (!registration?.is_registered || registration.is_registered !== "true") {
        await answerCallbackQuery(env, callbackQuery.id, TEXT.NOT_REGISTERED, true);
        return;
      }
      await startBookingFlow(env, chatId, userId, admin);
      await answerCallbackQuery(env, callbackQuery.id, "Оберіть дату.", false);
      return;
    }

    if (data === CALLBACKS.MY_BOOKINGS) {
      await clearState(env, chatId);
      await sendUserBookings(env, chatId, userId);
      await answerCallbackQuery(env, callbackQuery.id, "Показую ваші заявки.", false);
      return;
    }

    if (data === CALLBACKS.BACK_MAIN) {
      await clearState(env, chatId);
      const registration = await getUserRegistration(env, userId);
      await sendMessage(env, chatId, TEXT.START_REGISTERED, {
        reply_markup: mainMenuKeyboard(Boolean(registration?.is_registered === "true"), admin)
      });
      await answerCallbackQuery(env, callbackQuery.id, "Головне меню.", false);
      return;
    }

    if (data.startsWith(CALLBACKS.DATE_PREFIX)) {
      await handleDateSelection(callbackQuery, env);
      return;
    }

    if (data.startsWith(CALLBACKS.TIME_PREFIX)) {
      await handleTimeSelection(callbackQuery, env);
      return;
    }

    if (data === CALLBACKS.CONFIRM_BOOKING) {
      await handleBookingConfirm(callbackQuery, env);
      return;
    }

    if (data === CALLBACKS.CANCEL_BOOKING_FLOW) {
      await clearState(env, chatId);
      await answerCallbackQuery(env, callbackQuery.id, TEXT.BOOKING_CANCELLED_FLOW, false);
      await sendMessage(env, chatId, TEXT.BOOKING_CANCELLED_FLOW, { reply_markup: mainMenuKeyboard(true, admin) });
      return;
    }

    if (data.startsWith(CALLBACKS.CANCEL_PREFIX)) {
      await handleUserCancelBooking(callbackQuery, env);
      return;
    }

    if (data === CALLBACKS.ADMIN_PENDING) {
      if (!admin) {
        await answerCallbackQuery(env, callbackQuery.id, TEXT.ADMIN_ONLY, true);
        return;
      }
      await sendPendingBookings(env, chatId);
      await answerCallbackQuery(env, callbackQuery.id, "Показую нові заявки.", false);
      return;
    }

    if (data === CALLBACKS.ADMIN_TODAY) {
      if (!admin) {
        await answerCallbackQuery(env, callbackQuery.id, TEXT.ADMIN_ONLY, true);
        return;
      }
      await sendTodayBookings(env, chatId);
      await answerCallbackQuery(env, callbackQuery.id, "Показую записи на сьогодні.", false);
      return;
    }

    if (data.startsWith(CALLBACKS.ADMIN_APPROVE_PREFIX) || data.startsWith(CALLBACKS.ADMIN_REJECT_PREFIX)) {
      await handleAdminDecision(callbackQuery, env);
      return;
    }

    await answerCallbackQuery(env, callbackQuery.id, TEXT.UNKNOWN_ACTION, true);
  } catch (error) {
    await logError(env, "handleCallbackQuery", error, { callbackQuery });
    await answerCallbackQuery(env, callbackQuery.id, TEXT.UNKNOWN_ACTION, true);
  }
}

async function handleDateSelection(callbackQuery, env) {
  const chatId = callbackQuery.message?.chat?.id;
  if (!chatId) return;

  const selectedDate = callbackQuery.data.slice(CALLBACKS.DATE_PREFIX.length);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(selectedDate)) {
    await answerCallbackQuery(env, callbackQuery.id, TEXT.INVALID_DATE, true);
    return;
  }

  const freeSlots = await getFreeSlots(env, selectedDate);
  if (!freeSlots.length) {
    await answerCallbackQuery(env, callbackQuery.id, TEXT.NO_FREE_SLOTS, true);
    return;
  }

  await saveState(env, chatId, {
    step: FORM_STEPS.WAITING_BOOKING_TIME,
    data: { booking_date: selectedDate }
  });

  await answerCallbackQuery(env, callbackQuery.id, `Дата: ${formatHumanDate(selectedDate)}`, false);
  await sendMessage(env, chatId, `${TEXT.ASK_TIME}\n\n📅 ${escapeHtml(formatHumanDate(selectedDate))}`, {
    reply_markup: bookingTimesKeyboard(selectedDate, freeSlots)
  });
}

async function handleTimeSelection(callbackQuery, env) {
  const chatId = callbackQuery.message?.chat?.id;
  if (!chatId) return;

  const payload = callbackQuery.data.slice(CALLBACKS.TIME_PREFIX.length);
  const [date, time] = payload.split("|");
  if (!date || !time) {
    await answerCallbackQuery(env, callbackQuery.id, TEXT.INVALID_TIME, true);
    return;
  }

  const freeSlots = await getFreeSlots(env, date);
  if (!freeSlots.includes(time)) {
    await answerCallbackQuery(env, callbackQuery.id, TEXT.SLOT_ALREADY_TAKEN, true);
    return;
  }

  await saveState(env, chatId, {
    step: FORM_STEPS.WAITING_BOOKING_COMMENT,
    data: { booking_date: date, booking_time: time, slot_key: `${date}|${time}` }
  });

  await answerCallbackQuery(env, callbackQuery.id, `Час: ${time}`, false);
  await sendMessage(env, chatId, `${TEXT.ASK_COMMENT}\n\n📅 ${escapeHtml(formatHumanDate(date))}\n🕒 ${escapeHtml(time)}`, {
    reply_markup: skipKeyboard()
  });
}

async function handleBookingConfirm(callbackQuery, env) {
  const chatId = callbackQuery.message?.chat?.id;
  const userId = callbackQuery.from?.id;
  if (!chatId) return;

  const state = await getState(env, chatId);
  if (!state || state.step !== FORM_STEPS.WAITING_BOOKING_CONFIRM) {
    await answerCallbackQuery(env, callbackQuery.id, TEXT.NOTHING_TO_CONFIRM, true);
    return;
  }

  const registration = await getUserRegistration(env, userId);
  if (!registration?.is_registered || registration.is_registered !== "true") {
    await answerCallbackQuery(env, callbackQuery.id, TEXT.NOT_REGISTERED, true);
    return;
  }

  const draft = state.data || {};
  const freeSlots = await getFreeSlots(env, draft.booking_date);
  if (!freeSlots.includes(draft.booking_time)) {
    await answerCallbackQuery(env, callbackQuery.id, TEXT.SLOT_ALREADY_TAKEN, true);
    await clearState(env, chatId);
    await sendMessage(env, chatId, TEXT.SLOT_ALREADY_TAKEN, { reply_markup: mainMenuKeyboard(true, isAdmin(env, userId)) });
    return;
  }

  const now = formatDateTime(env, new Date());
  const booking = {
    booking_id: generateBookingId(),
    created_at: now,
    updated_at: now,
    telegram_id: String(userId || chatId),
    username: registration.username || getTelegramUsername(callbackQuery.from),
    full_name: registration.full_name || "",
    phone: registration.phone || "",
    booking_date: draft.booking_date,
    booking_time: draft.booking_time,
    slot_key: `${draft.booking_date}|${draft.booking_time}`,
    comment: draft.comment || "",
    status: STATUS_PENDING,
    admin_action_by: "",
    approved_at: "",
    rejected_at: "",
    cancelled_at: "",
    cancelled_by: ""
  };

  await appendBooking(env, booking);
  await clearState(env, chatId);
  await notifyAdminsAboutBooking(env, booking);

  await answerCallbackQuery(env, callbackQuery.id, TEXT.BOOKING_SAVED, false);
  await sendMessage(env, chatId, TEXT.BOOKING_SAVED, { reply_markup: mainMenuKeyboard(true, isAdmin(env, userId)) });
}

async function handleUserCancelBooking(callbackQuery, env) {
  const bookingId = callbackQuery.data.slice(CALLBACKS.CANCEL_PREFIX.length);
  const userId = String(callbackQuery.from?.id || "");
  const chatId = callbackQuery.message?.chat?.id;
  if (!chatId) return;

  const booking = await findBookingById(env, bookingId);
  if (!booking || String(booking.telegram_id || "") !== userId) {
    await answerCallbackQuery(env, callbackQuery.id, TEXT.UNKNOWN_ACTION, true);
    return;
  }

  if (![STATUS_PENDING, STATUS_APPROVED].includes(booking.status || "")) {
    await answerCallbackQuery(env, callbackQuery.id, TEXT.BOOKING_CANCEL_DENIED, true);
    return;
  }

  const adminLabel = formatAdminLabel(callbackQuery.from);
  await updateBookingStatus(env, bookingId, STATUS_CANCELLED, {
    admin_action_by: adminLabel,
    cancelled_at: formatDateTime(env, new Date()),
    cancelled_by: adminLabel
  });

  await answerCallbackQuery(env, callbackQuery.id, TEXT.BOOKING_CANCELLED, false);
  await sendMessage(env, chatId, TEXT.BOOKING_CANCELLED, {
    reply_markup: mainMenuKeyboard(true, isAdmin(env, callbackQuery.from?.id))
  });
}

async function handleAdminDecision(callbackQuery, env) {
  const admin = isAdmin(env, callbackQuery.from?.id);
  if (!admin) {
    await answerCallbackQuery(env, callbackQuery.id, TEXT.ADMIN_ONLY, true);
    return;
  }

  const approve = callbackQuery.data.startsWith(CALLBACKS.ADMIN_APPROVE_PREFIX);
  const bookingId = callbackQuery.data.replace(CALLBACKS.ADMIN_APPROVE_PREFIX, "").replace(CALLBACKS.ADMIN_REJECT_PREFIX, "");
  const booking = await findBookingById(env, bookingId);
  if (!booking) {
    await answerCallbackQuery(env, callbackQuery.id, "Заявку не знайдено.", true);
    return;
  }

  if (booking.status !== STATUS_PENDING) {
    await answerCallbackQuery(env, callbackQuery.id, "Статус уже змінено раніше.", true);
    return;
  }

  const now = formatDateTime(env, new Date());
  const adminLabel = formatAdminLabel(callbackQuery.from);
  const nextStatus = approve ? STATUS_APPROVED : STATUS_REJECTED;
  const extra = approve
    ? { admin_action_by: adminLabel, approved_at: now, rejected_at: "", cancelled_at: "", cancelled_by: "" }
    : { admin_action_by: adminLabel, approved_at: "", rejected_at: now, cancelled_at: "", cancelled_by: "" };

  await updateBookingStatus(env, bookingId, nextStatus, extra);
  await notifyUserAboutDecision(env, booking, nextStatus);

  const refreshed = await findBookingById(env, bookingId);
  const text = formatAdminBooking(refreshed || { ...booking, status: nextStatus, ...extra });
  await telegramApi(env, "editMessageText", {
    chat_id: callbackQuery.message.chat.id,
    message_id: callbackQuery.message.message_id,
    text,
    parse_mode: "HTML"
  });

  await answerCallbackQuery(env, callbackQuery.id, `${TEXT.STATUS_UPDATED} ${nextStatus}`, false);
}

async function sendUserBookings(env, chatId, userId) {
  const rows = await getBookings(env);
  const items = rows
    .filter((row) => String(row.telegram_id || "") === String(userId || chatId))
    .sort(compareBookingsDesc)
    .slice(0, 20);

  if (!items.length) {
    await sendMessage(env, chatId, TEXT.NO_BOOKINGS, { reply_markup: mainMenuKeyboard(true, isAdmin(env, userId)) });
    return;
  }

  for (const row of items) {
    const canCancel = [STATUS_PENDING, STATUS_APPROVED].includes(row.status || "");
    await sendMessage(env, chatId, formatUserBooking(row), {
      reply_markup: canCancel ? cancelBookingKeyboard(row.booking_id) : mainMenuInlineKeyboard(isAdmin(env, userId))
    });
  }
}

async function sendPendingBookings(env, chatId) {
  const rows = await getBookings(env);
  const pending = rows.filter((row) => row.status === STATUS_PENDING).sort(compareBookingsDesc).slice(0, 20);

  if (!pending.length) {
    await sendMessage(env, chatId, TEXT.PENDING_EMPTY, { reply_markup: mainMenuKeyboard(true, true) });
    return;
  }

  await sendMessage(env, chatId, TEXT.ADMIN_PENDING_TITLE, { reply_markup: adminMenuInlineKeyboard() });
  for (const row of pending) {
    await sendMessage(env, chatId, formatAdminBooking(row), {
      reply_markup: adminDecisionKeyboard(row.booking_id)
    });
  }
}

async function sendTodayBookings(env, chatId) {
  const rows = await getBookings(env);
  const today = getTodayIso(env);
  const items = rows
    .filter((row) => row.booking_date === today && [STATUS_PENDING, STATUS_APPROVED].includes(row.status || ""))
    .sort(compareBookingsAsc)
    .slice(0, 30);

  if (!items.length) {
    await sendMessage(env, chatId, TEXT.TODAY_EMPTY, { reply_markup: mainMenuKeyboard(true, true) });
    return;
  }

  const header = `${TEXT.ADMIN_TODAY_TITLE}\n📅 ${escapeHtml(formatHumanDate(today))}\nКількість: <b>${items.length}</b>`;
  await sendMessage(env, chatId, header, { reply_markup: adminMenuInlineKeyboard() });
  await sendChunkedText(env, chatId, items.map(formatTodayLine).join("\n\n"), mainMenuKeyboard(true, true));
}

async function sendStats(env, chatId) {
  const rows = await getBookings(env);
  if (!rows.length) {
    await sendMessage(env, chatId, TEXT.STATS_EMPTY, { reply_markup: mainMenuKeyboard(true, true) });
    return;
  }

  const stats = {
    total: rows.length,
    pending: 0,
    approved: 0,
    rejected: 0,
    cancelled: 0
  };
  for (const row of rows) {
    if (row.status === STATUS_PENDING) stats.pending += 1;
    if (row.status === STATUS_APPROVED) stats.approved += 1;
    if (row.status === STATUS_REJECTED) stats.rejected += 1;
    if (row.status === STATUS_CANCELLED) stats.cancelled += 1;
  }

  const today = getTodayIso(env);
  const todayCount = rows.filter((row) => row.booking_date === today && [STATUS_PENDING, STATUS_APPROVED].includes(row.status || "")).length;

  const text = [
    TEXT.STATS_TITLE,
    `Усього заявок: <b>${stats.total}</b>`,
    `На сьогодні активних: <b>${todayCount}</b>`,
    `Очікує підтвердження: <b>${stats.pending}</b>`,
    `Підтверджено: <b>${stats.approved}</b>`,
    `Відхилено: <b>${stats.rejected}</b>`,
    `Скасовано: <b>${stats.cancelled}</b>`
  ].join("\n");

  await sendMessage(env, chatId, text, { reply_markup: mainMenuKeyboard(true, true) });
}

async function findBookingsByPhone(env, chatId, commandText) {
  const phone = normalizePhone(commandText.replace("/find_phone", "").trim());
  if (!phone) {
    await sendMessage(env, chatId, TEXT.FIND_PHONE_USAGE, { reply_markup: mainMenuKeyboard(true, true) });
    return;
  }

  const rows = await getBookings(env);
  const matched = rows.filter((row) => normalizePhone(row.phone || "") === phone).sort(compareBookingsDesc).slice(0, 20);
  if (!matched.length) {
    await sendMessage(env, chatId, TEXT.FIND_PHONE_EMPTY, { reply_markup: mainMenuKeyboard(true, true) });
    return;
  }

  await sendChunkedText(env, chatId, matched.map(formatAdminBooking).join("\n\n──────────\n\n"), mainMenuKeyboard(true, true));
}

function mainMenuKeyboard(isRegistered, isAdminUser = false) {
  if (!isRegistered) {
    return { keyboard: [[{ text: "/start" }], [{ text: "/cancel" }]], resize_keyboard: true };
  }
  const keyboard = [[{ text: "Записатися" }], [{ text: "Мої заявки" }], [{ text: "Головне меню" }]];
  if (isAdminUser) {
    keyboard.push([{ text: "/pending" }, { text: "/today" }]);
    keyboard.push([{ text: "/stats" }]);
  }
  return { keyboard, resize_keyboard: true };
}

function mainMenuInlineKeyboard(isAdminUser = false) {
  const rows = [
    [{ text: "➕ Записатися", callback_data: CALLBACKS.START_BOOKING }],
    [{ text: "📋 Мої заявки", callback_data: CALLBACKS.MY_BOOKINGS }],
    [{ text: "🏠 Меню", callback_data: CALLBACKS.BACK_MAIN }]
  ];
  if (isAdminUser) rows.push([{ text: "🆕 Нові заявки", callback_data: CALLBACKS.ADMIN_PENDING }]);
  return { inline_keyboard: rows };
}

function adminMenuInlineKeyboard() {
  return {
    inline_keyboard: [
      [{ text: "🆕 Нові заявки", callback_data: CALLBACKS.ADMIN_PENDING }],
      [{ text: "📅 Сьогодні", callback_data: CALLBACKS.ADMIN_TODAY }],
      [{ text: "🏠 Меню", callback_data: CALLBACKS.BACK_MAIN }]
    ]
  };
}

function cancelKeyboard() {
  return { keyboard: [[{ text: "/cancel" }]], resize_keyboard: true };
}

function skipKeyboard() {
  return { keyboard: [[{ text: TEXT.SKIP }], [{ text: "/cancel" }]], resize_keyboard: true };
}

function bookingDatesKeyboard(dates, isAdminUser = false) {
  const rows = [];
  for (let i = 0; i < dates.length; i += 2) {
    const pair = dates.slice(i, i + 2).map((item) => ({
      text: `${item.weekday} ${item.human}`,
      callback_data: `${CALLBACKS.DATE_PREFIX}${item.iso}`
    }));
    rows.push(pair);
  }
  rows.push([{ text: "❌ Скасувати", callback_data: CALLBACKS.CANCEL_BOOKING_FLOW }]);
  rows.push([{ text: "🏠 Меню", callback_data: CALLBACKS.BACK_MAIN }]);
  if (isAdminUser) rows.push([{ text: "🆕 Нові заявки", callback_data: CALLBACKS.ADMIN_PENDING }]);
  return { inline_keyboard: rows };
}

function bookingTimesKeyboard(date, times) {
  const rows = [];
  for (let i = 0; i < times.length; i += 3) {
    rows.push(
      times.slice(i, i + 3).map((time) => ({
        text: time,
        callback_data: `${CALLBACKS.TIME_PREFIX}${date}|${time}`
      }))
    );
  }
  rows.push([{ text: "⬅️ Інша дата", callback_data: CALLBACKS.START_BOOKING }]);
  rows.push([{ text: "❌ Скасувати", callback_data: CALLBACKS.CANCEL_BOOKING_FLOW }]);
  return { inline_keyboard: rows };
}

function bookingConfirmKeyboard() {
  return {
    inline_keyboard: [
      [{ text: "✅ Підтвердити запис", callback_data: CALLBACKS.CONFIRM_BOOKING }],
      [{ text: "❌ Скасувати", callback_data: CALLBACKS.CANCEL_BOOKING_FLOW }]
    ]
  };
}

function cancelBookingKeyboard(bookingId) {
  return {
    inline_keyboard: [
      [{ text: "❌ Скасувати бронювання", callback_data: `${CALLBACKS.CANCEL_PREFIX}${bookingId}` }],
      [{ text: "🏠 Меню", callback_data: CALLBACKS.BACK_MAIN }]
    ]
  };
}

function adminDecisionKeyboard(bookingId) {
  return {
    inline_keyboard: [
      [{ text: "✅ Підтвердити", callback_data: `${CALLBACKS.ADMIN_APPROVE_PREFIX}${bookingId}` }],
      [{ text: "⛔ Відхилити", callback_data: `${CALLBACKS.ADMIN_REJECT_PREFIX}${bookingId}` }]
    ]
  };
}

function formatUserBooking(row) {
  return [
    `🆔 <b>${escapeHtml(row.booking_id || "-")}</b>`,
    `📅 <b>Дата:</b> ${escapeHtml(formatHumanDate(row.booking_date || ""))}`,
    `🕒 <b>Час:</b> ${escapeHtml(row.booking_time || "-")}`,
    `📌 <b>Статус:</b> ${escapeHtml(row.status || "-")}`,
    `💬 <b>Коментар:</b> ${escapeHtml(row.comment || "—")}`,
    `🛠 <b>Оновлено:</b> ${escapeHtml(row.updated_at || row.created_at || "-")}`
  ].join("\n");
}

function formatAdminBooking(row) {
  return [
    "<b>Заявка на стрижку</b>",
    `🆔 <b>${escapeHtml(row.booking_id || "-")}</b>`,
    `👤 <b>Клієнт:</b> ${escapeHtml(row.full_name || "-")}`,
    `📞 <b>Телефон:</b> ${escapeHtml(row.phone || "-")}`,
    `🔗 <b>Username:</b> ${escapeHtml(row.username || "—")}`,
    `📅 <b>Дата:</b> ${escapeHtml(formatHumanDate(row.booking_date || ""))}`,
    `🕒 <b>Час:</b> ${escapeHtml(row.booking_time || "-")}`,
    `💬 <b>Коментар:</b> ${escapeHtml(row.comment || "—")}`,
    `📌 <b>Статус:</b> ${escapeHtml(row.status || "-")}`,
    `🛠 <b>Остання дія:</b> ${escapeHtml(row.admin_action_by || "—")}`
  ].join("\n");
}

function formatTodayLine(row) {
  return [
    `🕒 <b>${escapeHtml(row.booking_time || "-")}</b>`,
    `👤 ${escapeHtml(row.full_name || "-")}`,
    `📞 ${escapeHtml(row.phone || "-")}`,
    `📌 ${escapeHtml(row.status || "-")}`
  ].join("\n");
}

async function notifyAdminsAboutBooking(env, booking) {
  const adminIds = parseAdminIds(env.ADMIN_IDS);
  for (const adminId of adminIds) {
    try {
      await sendMessage(env, adminId, formatAdminBooking(booking), {
        reply_markup: adminDecisionKeyboard(booking.booking_id)
      });
    } catch (error) {
      await logError(env, "notifyAdminsAboutBooking", error, { adminId, bookingId: booking.booking_id });
    }
  }
}

async function notifyUserAboutDecision(env, booking, status) {
  const chatId = Number(booking.telegram_id || 0);
  if (!chatId) return;

  const text = [
    `<b>Оновлення по заявці ${escapeHtml(booking.booking_id || "")}</b>`,
    `📅 <b>Дата:</b> ${escapeHtml(formatHumanDate(booking.booking_date || ""))}`,
    `🕒 <b>Час:</b> ${escapeHtml(booking.booking_time || "-")}`,
    `📌 <b>Новий статус:</b> ${escapeHtml(status)}`
  ].join("\n");

  try {
    await sendMessage(env, chatId, text, { reply_markup: mainMenuKeyboard(true, false) });
  } catch (error) {
    await logError(env, "notifyUserAboutDecision", error, { chatId, bookingId: booking.booking_id, status });
  }
}

async function getFreeSlots(env, date) {
  const allSlots = generateDailySlots(env);
  const rows = await getBookings(env);
  const busy = new Set(
    rows
      .filter((row) => row.booking_date === date)
      .filter((row) => [STATUS_PENDING, STATUS_APPROVED].includes(row.status || ""))
      .map((row) => row.booking_time)
  );
  return allSlots.filter((slot) => !busy.has(slot));
}

function generateDailySlots(env) {
  const startHour = Number(env.WORK_START_HOUR || DEFAULT_WORK_START_HOUR);
  const endHour = Number(env.WORK_END_HOUR || DEFAULT_WORK_END_HOUR);
  const slotMinutes = Number(env.SLOT_MINUTES || DEFAULT_SLOT_MINUTES);

  const slots = [];
  for (let totalMinutes = startHour * 60; totalMinutes < endHour * 60; totalMinutes += slotMinutes) {
    const hh = String(Math.floor(totalMinutes / 60)).padStart(2, "0");
    const mm = String(totalMinutes % 60).padStart(2, "0");
    slots.push(`${hh}:${mm}`);
  }
  return slots;
}

function getUpcomingDates(env, days) {
  const dates = [];
  const now = new Date();
  for (let i = 0; i < days; i++) {
    const date = new Date(now.getTime());
    date.setDate(now.getDate() + i);
    const iso = formatDateIso(env, date);
    const human = formatHumanDate(iso);
    const weekday = new Intl.DateTimeFormat("uk-UA", {
      timeZone: env.TIMEZONE || DEFAULT_TIMEZONE,
      weekday: "short"
    }).format(date);
    dates.push({ iso, human, weekday });
  }
  return dates;
}

function formatHumanDate(isoDate) {
  if (!isoDate || !/^\d{4}-\d{2}-\d{2}$/.test(String(isoDate))) return "-";
  const [yyyy, mm, dd] = String(isoDate).split("-");
  return `${dd}.${mm}.${yyyy}`;
}

function formatDateIso(env, date) {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: env.TIMEZONE || DEFAULT_TIMEZONE,
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  }).formatToParts(date);
  const map = Object.fromEntries(parts.map((p) => [p.type, p.value]));
  return `${map.year}-${map.month}-${map.day}`;
}

function getTodayIso(env) {
  return formatDateIso(env, new Date());
}

async function isDuplicateUpdate(env, updateId) {
  const key = `update:${updateId}`;
  const value = await env.STATE_KV.get(key);
  return value === "1";
}

async function markUpdateProcessed(env, updateId) {
  await env.STATE_KV.put(`update:${updateId}`, "1", { expirationTtl: 60 * 30 });
}

async function getState(env, chatId) {
  const raw = await env.STATE_KV.get(`state:${chatId}`);
  return raw ? JSON.parse(raw) : null;
}

async function saveState(env, chatId, state) {
  await env.STATE_KV.put(`state:${chatId}`, JSON.stringify(state), { expirationTtl: 60 * 60 * 12 });
}

async function clearState(env, chatId) {
  await env.STATE_KV.delete(`state:${chatId}`);
}

async function checkRateLimit(env, chatId, userId) {
  const key = `rate:${chatId}:${userId || "0"}`;
  const raw = await env.STATE_KV.get(key);
  const now = Date.now();
  const bucket = raw ? JSON.parse(raw) : { count: 0, start: now };

  if (now - bucket.start > 15000) {
    bucket.count = 0;
    bucket.start = now;
  }

  bucket.count += 1;
  await env.STATE_KV.put(key, JSON.stringify(bucket), { expirationTtl: 60 });
  return bucket.count <= 18;
}

function validateName(value) {
  return String(value || "").trim().length >= 2;
}

function normalizePhone(value) {
  return String(value || "").replace(/[^\d]/g, "");
}

function validatePhone(value) {
  return /^380\d{9}$/.test(String(value || ""));
}

function getTelegramUsername(user) {
  if (!user) return "";
  if (user.username) return `@${user.username}`;
  return [user.first_name, user.last_name].filter(Boolean).join(" ").trim();
}

function parseAdminIds(raw) {
  return String(raw || "").split(",").map((item) => item.trim()).filter(Boolean);
}

function isAdmin(env, userId) {
  return parseAdminIds(env.ADMIN_IDS).includes(String(userId || ""));
}

function generateBookingId() {
  const rand = Math.random().toString(36).slice(2, 7).toUpperCase();
  return `BOOK-${Date.now()}-${rand}`;
}

function compareBookingsDesc(a, b) {
  return (`${b.booking_date} ${b.booking_time}`).localeCompare(`${a.booking_date} ${a.booking_time}`);
}

function compareBookingsAsc(a, b) {
  return (`${a.booking_date} ${a.booking_time}`).localeCompare(`${b.booking_date} ${b.booking_time}`);
}

function validateEnv(env) {
  const required = ["BOT_TOKEN", "GOOGLE_SHEET_ID", "GOOGLE_SERVICE_ACCOUNT_JSON", "ADMIN_IDS", "STATE_KV", "ACCESS_PASSWORD"];
  for (const key of required) {
    if (!env[key]) {
      throw new Error(`Missing environment variable: ${key}`);
    }
  }
}

async function ensureAllWorksheets(env) {
  await ensureSheetHeaders(env, getBookingsSheetName(env), BOOKING_HEADERS);
  await ensureSheetHeaders(env, getUsersSheetName(env), USER_HEADERS);
  await ensureSheetHeaders(env, getLogsSheetName(env), LOG_HEADERS);
}

function getBookingsSheetName(env) {
  return env.BOOKINGS_SHEET_NAME || DEFAULT_BOOKINGS_SHEET;
}

function getUsersSheetName(env) {
  return env.USERS_SHEET_NAME || DEFAULT_USERS_SHEET;
}

function getLogsSheetName(env) {
  return env.LOGS_SHEET_NAME || DEFAULT_LOGS_SHEET;
}

async function ensureSheetHeaders(env, sheetTitle, requiredHeaders) {
  const metadata = await sheetsGetMetadata(env);
  let found = metadata.sheets?.find((sheet) => sheet.properties?.title === sheetTitle);

  if (!found) {
    try {
      await sheetsBatchUpdate(env, { requests: [{ addSheet: { properties: { title: sheetTitle } } }] });
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      if (!msg.includes(`A sheet with the name \"${sheetTitle}\" already exists`)) {
        throw error;
      }
    }
    const refreshed = await sheetsGetMetadata(env);
    found = refreshed.sheets?.find((sheet) => sheet.properties?.title === sheetTitle);
  }

  const lastColumn = indexToColumn(requiredHeaders.length - 1);
  const current = await sheetsGetValues(env, `${sheetTitle}!A1:${lastColumn}1`);
  const firstRow = current[0] || [];

  if (!firstRow.length) {
    await sheetsUpdateValues(env, `${sheetTitle}!A1:${lastColumn}1`, [requiredHeaders]);
    return { headers: requiredHeaders, sheetId: found?.properties?.sheetId };
  }

  const merged = firstRow.slice();
  for (const header of requiredHeaders) {
    if (!merged.includes(header)) merged.push(header);
  }

  if (merged.length !== firstRow.length) {
    await sheetsUpdateValues(env, `${sheetTitle}!A1:${indexToColumn(merged.length - 1)}1`, [merged]);
  }

  return { headers: merged, sheetId: found?.properties?.sheetId };
}

async function appendBooking(env, booking) {
  const ctx = await ensureSheetHeaders(env, getBookingsSheetName(env), BOOKING_HEADERS);
  const row = ctx.headers.map((header) => booking[header] ?? "");
  await withGoogleRetry(() => sheetsAppend(env, `${getBookingsSheetName(env)}!A:${indexToColumn(ctx.headers.length - 1)}`, [row]));
}

async function getBookings(env) {
  const ctx = await ensureSheetHeaders(env, getBookingsSheetName(env), BOOKING_HEADERS);
  const values = await sheetsGetValues(env, `${getBookingsSheetName(env)}!A:${indexToColumn(ctx.headers.length - 1)}`);
  if (!values.length) return [];
  const headers = values[0];
  return values.slice(1).map((row) => rowToObject(headers, row));
}

async function findBookingById(env, bookingId) {
  const rows = await getBookings(env);
  return rows.find((row) => row.booking_id === bookingId) || null;
}

async function updateBookingStatus(env, bookingId, status, extra = {}) {
  const ctx = await ensureSheetHeaders(env, getBookingsSheetName(env), BOOKING_HEADERS);
  const headers = ctx.headers;
  const values = await sheetsGetValues(env, `${getBookingsSheetName(env)}!A:${indexToColumn(headers.length - 1)}`);
  const idIndex = headers.indexOf("booking_id");
  if (idIndex === -1) throw new Error("booking_id column not found");

  for (let i = 1; i < values.length; i++) {
    if ((values[i][idIndex] || "") !== bookingId) continue;
    const rowNumber = i + 1;
    const updates = { ...rowToObject(headers, values[i]), ...extra, status, updated_at: formatDateTime(env, new Date()) };

    for (const [key, value] of Object.entries(updates)) {
      const idx = headers.indexOf(key);
      if (idx === -1) continue;
      await withGoogleRetry(() => sheetsUpdateValues(env, `${getBookingsSheetName(env)}!${indexToColumn(idx)}${rowNumber}`, [[value]]));
    }
    return true;
  }
  return false;
}

async function getUsers(env) {
  const ctx = await ensureSheetHeaders(env, getUsersSheetName(env), USER_HEADERS);
  const values = await sheetsGetValues(env, `${getUsersSheetName(env)}!A:${indexToColumn(ctx.headers.length - 1)}`);
  if (!values.length) return [];
  const headers = values[0];
  return values.slice(1).map((row) => rowToObject(headers, row));
}

async function getUserRegistration(env, telegramId) {
  const rows = await getUsers(env);
  const matches = rows.filter((row) => String(row.telegram_id || "") === String(telegramId || ""));
  if (!matches.length) return null;

  const registered = matches.filter((row) => String(row.is_registered || "").toLowerCase() === "true");
  const pool = registered.length ? registered : matches;
  return pool.sort((a, b) => parseDateTimeValue(b.last_login_at || b.registered_at) - parseDateTimeValue(a.last_login_at || a.registered_at))[0] || null;
}

async function upsertUser(env, user) {
  const ctx = await ensureSheetHeaders(env, getUsersSheetName(env), USER_HEADERS);
  const headers = ctx.headers;
  const values = await sheetsGetValues(env, `${getUsersSheetName(env)}!A:${indexToColumn(headers.length - 1)}`);
  const idIndex = headers.indexOf("telegram_id");

  for (let i = 1; i < values.length; i++) {
    if ((values[i][idIndex] || "") !== String(user.telegram_id || "")) continue;
    const rowNumber = i + 1;
    const merged = { ...rowToObject(headers, values[i]), ...user };
    for (const [key, value] of Object.entries(merged)) {
      const idx = headers.indexOf(key);
      if (idx === -1) continue;
      await withGoogleRetry(() => sheetsUpdateValues(env, `${getUsersSheetName(env)}!${indexToColumn(idx)}${rowNumber}`, [[value]]));
    }
    return;
  }

  const row = headers.map((header) => user[header] ?? "");
  await withGoogleRetry(() => sheetsAppend(env, `${getUsersSheetName(env)}!A:${indexToColumn(headers.length - 1)}`, [row]));
}

async function touchUserLastLogin(env, telegramId) {
  const registration = await getUserRegistration(env, telegramId);
  if (!registration) return;
  await upsertUser(env, { ...registration, last_login_at: formatDateTime(env, new Date()) });
}

function rowToObject(headers, row) {
  const item = {};
  headers.forEach((header, index) => {
    item[header] = row[index] ?? "";
  });
  return item;
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
  if (cachedGoogleToken && cachedGoogleToken.expiresAt > Date.now() + 60000) {
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
  const base64 = String(pem || "")
    .replace("-----BEGIN PRIVATE KEY-----", "")
    .replace("-----END PRIVATE KEY-----", "")
    .replace(/\s+/g, "");
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

async function sendMessage(env, chatId, text, extra = {}) {
  return await telegramApi(env, "sendMessage", {
    chat_id: chatId,
    text,
    parse_mode: "HTML",
    disable_web_page_preview: true,
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

async function sendChunkedText(env, chatId, text, replyMarkup) {
  const chunks = splitLongText(text, 3800);
  for (let i = 0; i < chunks.length; i++) {
    await sendMessage(env, chatId, chunks[i], i === chunks.length - 1 ? { reply_markup: replyMarkup } : {});
  }
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

function formatDateTime(env, date) {
  return new Intl.DateTimeFormat("uk-UA", {
    timeZone: env.TIMEZONE || DEFAULT_TIMEZONE,
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false
  }).format(date);
}

function formatAdminLabel(adminUser) {
  if (!adminUser) return "Адміністратор";
  if (adminUser.username) return `@${adminUser.username}`;
  const fullName = [adminUser.first_name, adminUser.last_name].filter(Boolean).join(" ").trim();
  return fullName || String(adminUser.id || "Адміністратор");
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

async function logError(env, scope, error, details = null) {
  try {
    console.error(scope, error instanceof Error ? error.message : String(error), details || "");
    if (!env.GOOGLE_SHEET_ID || !env.GOOGLE_SERVICE_ACCOUNT_JSON) return;
    await ensureSheetHeaders(env, getLogsSheetName(env), LOG_HEADERS);
    const row = [
      formatDateTime(env, new Date()),
      "error",
      scope,
      error instanceof Error ? error.message : String(error),
      details ? JSON.stringify(details).slice(0, 3000) : ""
    ];
    await sheetsAppend(env, `${getLogsSheetName(env)}!A:E`, [row]);
  } catch (nestedError) {
    console.error("logError failed", nestedError instanceof Error ? nestedError.message : String(nestedError));
  }
}
