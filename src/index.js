const DEFAULT_TIMEZONE = "Europe/Kyiv";
const DEFAULT_CALENDAR_ID = "techno.perspektiva@gmail.com";
const DEFAULT_MODEL = "gemini-2.5-flash";
const MAX_HISTORY_MESSAGES = 14;
const MAX_TELEGRAM_MESSAGE = 3800;

const SHEETS = {
  TASKS: "Задачі",
  BUDGET: "Бюджет",
  SHOPPING: "Покупки",
  LOG: "Журнал",
  NOTES: "Нотатки"
};

const TASK_HEADERS = [
  "task_id",
  "created_at",
  "updated_at",
  "chat_id",
  "user_id",
  "username",
  "title",
  "details",
  "due_at",
  "due_date",
  "status",
  "calendar_event_id",
  "source_text"
];

const BUDGET_HEADERS = [
  "entry_id",
  "created_at",
  "updated_at",
  "chat_id",
  "user_id",
  "username",
  "type",
  "amount",
  "currency",
  "category",
  "note",
  "source_text"
];

const SHOPPING_HEADERS = [
  "item_id",
  "created_at",
  "updated_at",
  "chat_id",
  "user_id",
  "username",
  "list_type",
  "item_name",
  "details",
  "quantity",
  "status",
  "source_text"
];

const NOTE_HEADERS = [
  "note_id",
  "created_at",
  "chat_id",
  "user_id",
  "username",
  "note_text",
  "source"
];

const LOG_HEADERS = [
  "created_at",
  "chat_id",
  "user_id",
  "username",
  "input_type",
  "original_text",
  "normalized_text",
  "action",
  "result_summary",
  "calendar_event_id",
  "status"
];

const TEXT = {
  START:
    "<b>Привіт. Я твій особистий Telegram-асистент.</b>\n\n" +
    "Я можу:\n" +
    "• створювати задачі і нагадування\n" +
    "• показувати список задач і план на сьогодні\n" +
    "• вести бюджет і покупки\n" +
    "• зберігати нотатки\n" +
    "• шукати інформацію і місця поруч\n" +
    "• розуміти голосові\n\n" +
    "Натисни кнопку нижче або просто напиши, що потрібно.",
  HELP:
    "<b>Приклади:</b>\n\n" +
    "• Створи задачу: тренування сьогодні о 19:00\n" +
    "• Додай витрату 320 грн на таксі\n" +
    "• Додай у продукти молоко 2 шт\n" +
    "• Занотуй: купити зарядку в машину\n" +
    "• Знайди кав'ярню біля Позняків\n" +
    "• Покажи мій план на сьогодні\n\n" +
    "<b>Команди:</b>\n/start\n/help\n/today\n/week\n/reset",
  RESET: "Готово, контекст очищено.",
  VOICE_BUSY: "Отримав голосове. Розшифровую…",
  PHOTO_BUSY: "Фото отримав. Аналізую…",
  SEARCH_BUSY: "Шукаю і збираю коротку відповідь…",
  TASK_PROMPT:
    "Напиши задачу одним повідомленням.\n\n" +
    "Наприклад:\n<i>Подзвонити клієнту завтра о 15:00</i>",
  NOTE_PROMPT: "Напиши текст нотатки одним повідомленням.",
  BUDGET_PROMPT:
    "Напиши витрату або дохід одним повідомленням.\n\n" +
    "Наприклад:\n<i>Витрата 320 грн таксі</i>\n<i>Дохід 5000 грн аванс</i>",
  PRODUCTS_PROMPT:
    "Напиши, що додати в продукти.\n\nНаприклад:\n<i>молоко 2 шт</i>",
  BUY_PROMPT:
    "Напиши, що потрібно купити.\n\nНаприклад:\n<i>кабель USB-C 2м</i>",
  NOTE_SAVED: "Готово, занотував.",
  TASK_SAVED: "Готово, задачу збережено.",
  BUDGET_SAVED: "Готово, записав у бюджет.",
  SHOPPING_SAVED: "Готово, додав у список.",
  NOTHING_FOUND: "Поки нічого не знайшов.",
  FALLBACK: "Не до кінця зрозумів запит. Спробуй коротше або натисни потрібну кнопку.",
  ERROR_CALENDAR: "Не зміг записати подію в Google Calendar. Перевір доступ календаря для service account.",
  ERROR_SHEETS: "Не зміг записати в Google Sheets. Перевір доступ таблиці для service account.",
  ERROR_GENERIC: "Сталася помилка. Спробуй ще раз трохи пізніше.",
  NEARBY_LOCATION_PROMPT:
    "Надішли геолокацію або напиши район / місто. Потім я попрошу, що саме шукати.",
  NEARBY_QUERY_PROMPT: "Напиши, що саме знайти поруч. Наприклад: <i>кав'ярня</i> або <i>спортзал</i>.",
  COMPLETE_PICK_PROMPT: "Відправ номер задачі зі списку, яку треба позначити як виконану.",
  COMPLETE_DONE: "Позначив задачу як виконану.",
  COMPLETE_NOT_FOUND: "Не знайшов такої задачі в сьогоднішньому списку.",
  TODAY_EMPTY: "На сьогодні активних задач і подій не знайшов.",
};

const MENU = {
  CREATE_TASK: "➕ Створити задачу",
  LIST_TASKS: "📋 Список задач",
  COMPLETE_TODAY: "✅ Виконати сьогодні",
  TODAY: "📅 Сьогодні",
  BUDGET: "💰 Бюджет",
  PRODUCTS: "🛒 Купити продукти",
  NOTE: "📝 Нотатка",
  BUY: "📦 Що купити",
  SEARCH: "🔎 Знайти",
  NEARBY: "📍 Місця поруч",
  WEEK_REPORT: "📊 Звіт за тиждень"
};

const STATES = {
  WAIT_TASK: "wait_task",
  WAIT_NOTE: "wait_note",
  WAIT_BUDGET: "wait_budget",
  WAIT_PRODUCTS: "wait_products",
  WAIT_BUY: "wait_buy",
  WAIT_SEARCH: "wait_search",
  WAIT_NEARBY_LOCATION: "wait_nearby_location",
  WAIT_NEARBY_QUERY: "wait_nearby_query",
  WAIT_COMPLETE_TODAY: "wait_complete_today"
};

const TASK_STATUS = {
  NEW: "Нова",
  IN_PROGRESS: "В роботі",
  DONE: "Виконано",
  MOVED: "Перенесено",
  CANCELED: "Скасовано"
};

const SHOPPING_STATUS = {
  OPEN: "Купити",
  DONE: "Куплено"
};

export default {
  async fetch(request, env, ctx) {
    try {
      const url = new URL(request.url);

      if (request.method === "GET" && url.pathname === "/") {
        return jsonResponse({ ok: true, service: "personal-assistant-v3", status: "running" });
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
  validateEnv(env);

  const url = new URL(request.url);
  const key = url.searchParams.get("key");
  if (!env.SETUP_KEY || key !== env.SETUP_KEY) {
    return jsonResponse({ ok: false, error: "Unauthorized" }, 401);
  }

  const origin = `${url.protocol}//${url.host}`;
  const webhookUrl = `${origin}/webhook`;

  const webhookResult = await telegramApi(env, "setWebhook", {
    url: webhookUrl,
    secret_token: env.TELEGRAM_WEBHOOK_SECRET,
    allowed_updates: ["message"]
  });

  const commandsResult = await telegramApi(env, "setMyCommands", {
    commands: [
      { command: "start", description: "Запустити асистента" },
      { command: "today", description: "План на сьогодні" },
      { command: "week", description: "Звіт за тиждень" },
      { command: "help", description: "Підказка" },
      { command: "reset", description: "Очистити контекст" }
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
    if (duplicate) return jsonResponse({ ok: true, duplicate: true });
    await markUpdateProcessed(env, updateId);
  }

  ctx.waitUntil(processUpdate(update, env));
  return jsonResponse({ ok: true });
}

async function processUpdate(update, env) {
  try {
    if (update.message) {
      await handleMessage(update.message, env);
    }
  } catch (error) {
    console.error("processUpdate error:", error instanceof Error ? error.stack || error.message : String(error));
  }
}

async function handleMessage(message, env) {
  const chatId = message.chat?.id;
  const from = message.from || {};
  if (!chatId) return;

  const rawText = (message.text || message.caption || "").trim();
  const normalizedButton = normalizeButtonText(rawText);

  if (rawText === "/start") {
    await clearChatState(env, chatId);
    await clearHistory(env, chatId);
    await sendMessage(env, chatId, TEXT.START, { reply_markup: mainMenuKeyboard() });
    return;
  }

  if (rawText === "/help") {
    await sendMessage(env, chatId, TEXT.HELP, { reply_markup: mainMenuKeyboard() });
    return;
  }

  if (rawText === "/reset") {
    await clearChatState(env, chatId);
    await clearHistory(env, chatId);
    await sendMessage(env, chatId, TEXT.RESET, { reply_markup: mainMenuKeyboard() });
    return;
  }

  if (rawText === "/today" || normalizedButton === normalizeButtonText(MENU.TODAY)) {
    await sendTodayOverview(env, chatId);
    return;
  }

  if (rawText === "/week" || normalizedButton === normalizeButtonText(MENU.WEEK_REPORT)) {
    await sendWeeklyReport(env, chatId);
    return;
  }

  if (normalizedButton === normalizeButtonText(MENU.CREATE_TASK)) {
    await saveChatState(env, chatId, { step: STATES.WAIT_TASK });
    await sendMessage(env, chatId, TEXT.TASK_PROMPT, { reply_markup: mainMenuKeyboard() });
    return;
  }

  if (normalizedButton === normalizeButtonText(MENU.NOTE)) {
    await saveChatState(env, chatId, { step: STATES.WAIT_NOTE });
    await sendMessage(env, chatId, TEXT.NOTE_PROMPT, { reply_markup: mainMenuKeyboard() });
    return;
  }

  if (normalizedButton === normalizeButtonText(MENU.BUDGET)) {
    await saveChatState(env, chatId, { step: STATES.WAIT_BUDGET });
    await sendBudgetSummaryAndPrompt(env, chatId);
    return;
  }

  if (normalizedButton === normalizeButtonText(MENU.PRODUCTS)) {
    await saveChatState(env, chatId, { step: STATES.WAIT_PRODUCTS });
    await sendShoppingSummaryAndPrompt(env, chatId, "products");
    return;
  }

  if (normalizedButton === normalizeButtonText(MENU.BUY)) {
    await saveChatState(env, chatId, { step: STATES.WAIT_BUY });
    await sendShoppingSummaryAndPrompt(env, chatId, "buy");
    return;
  }

  if (normalizedButton === normalizeButtonText(MENU.SEARCH)) {
    await saveChatState(env, chatId, { step: STATES.WAIT_SEARCH });
    await sendMessage(env, chatId, "Напиши, що саме знайти.", { reply_markup: mainMenuKeyboard() });
    return;
  }

  if (normalizedButton === normalizeButtonText(MENU.NEARBY)) {
    await saveChatState(env, chatId, { step: STATES.WAIT_NEARBY_LOCATION });
    await sendMessage(env, chatId, TEXT.NEARBY_LOCATION_PROMPT, { reply_markup: nearbyLocationKeyboard() });
    return;
  }

  if (normalizedButton === normalizeButtonText(MENU.LIST_TASKS)) {
    await sendTaskList(env, chatId);
    return;
  }

  if (normalizedButton === normalizeButtonText(MENU.COMPLETE_TODAY)) {
    await prepareCompleteToday(env, chatId);
    return;
  }

  const state = await getChatState(env, chatId);

  if (message.location && state?.step === STATES.WAIT_NEARBY_LOCATION) {
    await saveChatState(env, chatId, {
      step: STATES.WAIT_NEARBY_QUERY,
      location: {
        latitude: message.location.latitude,
        longitude: message.location.longitude
      }
    });
    await sendMessage(env, chatId, TEXT.NEARBY_QUERY_PROMPT, { reply_markup: mainMenuKeyboard() });
    return;
  }

  if (message.voice || message.audio) {
    await sendChatAction(env, chatId, "typing");
    await sendMessage(env, chatId, TEXT.VOICE_BUSY, { reply_markup: mainMenuKeyboard() });
    const transcript = await transcribeTelegramAudio(env, message.voice || message.audio);
    await processAssistantInput(env, {
      chatId,
      from,
      inputType: "voice",
      originalText: rawText || "[voice]",
      normalizedText: transcript,
      state,
      message
    });
    return;
  }

  if (message.photo?.length) {
    await sendChatAction(env, chatId, "typing");
    await sendMessage(env, chatId, TEXT.PHOTO_BUSY, { reply_markup: mainMenuKeyboard() });
    const analysis = await analyzeTelegramPhoto(
      env,
      message.photo,
      rawText || "Опиши коротко, що на фото, і дай практичну пораду українською."
    );
    await appendHistory(env, chatId, { role: "user", text: rawText || "[photo]" });
    await appendHistory(env, chatId, { role: "assistant", text: analysis });
    await logInteraction(env, {
      user: from,
      chatId,
      inputType: "photo",
      originalText: rawText || "[photo]",
      normalizedText: rawText || "[photo]",
      action: "analyze_photo",
      resultSummary: truncate(analysis, 250),
      status: "ok"
    });
    await sendMessage(env, chatId, analysis, { reply_markup: mainMenuKeyboard() });
    return;
  }

  if (!rawText) {
    await sendMessage(env, chatId, TEXT.FALLBACK, { reply_markup: mainMenuKeyboard() });
    return;
  }

  if (state?.step === STATES.WAIT_NOTE) {
    await saveNote(env, chatId, from, rawText, "manual");
    await clearChatState(env, chatId);
    await sendMessage(env, chatId, TEXT.NOTE_SAVED, { reply_markup: mainMenuKeyboard() });
    return;
  }

  if (state?.step === STATES.WAIT_TASK) {
    await createTaskFromText(env, chatId, from, rawText, "manual");
    await clearChatState(env, chatId);
    return;
  }

  if (state?.step === STATES.WAIT_BUDGET) {
    await addBudgetFromText(env, chatId, from, rawText, "manual");
    await clearChatState(env, chatId);
    return;
  }

  if (state?.step === STATES.WAIT_PRODUCTS) {
    await addShoppingItem(env, chatId, from, rawText, "products", "manual");
    await clearChatState(env, chatId);
    return;
  }

  if (state?.step === STATES.WAIT_BUY) {
    await addShoppingItem(env, chatId, from, rawText, "buy", "manual");
    await clearChatState(env, chatId);
    return;
  }

  if (state?.step === STATES.WAIT_SEARCH) {
    await clearChatState(env, chatId);
    await runSearch(env, chatId, from, rawText, "text");
    return;
  }

  if (state?.step === STATES.WAIT_NEARBY_LOCATION) {
    await saveChatState(env, chatId, { step: STATES.WAIT_NEARBY_QUERY, areaText: rawText });
    await sendMessage(env, chatId, TEXT.NEARBY_QUERY_PROMPT, { reply_markup: mainMenuKeyboard() });
    return;
  }

  if (state?.step === STATES.WAIT_NEARBY_QUERY) {
    await clearChatState(env, chatId);
    await runNearbyPlaces(env, chatId, from, rawText, state);
    return;
  }

  if (state?.step === STATES.WAIT_COMPLETE_TODAY) {
    await completeTaskFromSelection(env, chatId, rawText);
    await clearChatState(env, chatId);
    return;
  }

  await processAssistantInput(env, {
    chatId,
    from,
    inputType: "text",
    originalText: rawText,
    normalizedText: rawText,
    state,
    message
  });
}

async function processAssistantInput(env, payload) {
  const { chatId, from, inputType, originalText, normalizedText } = payload;
  await sendChatAction(env, chatId, "typing");

  const history = await getHistory(env, chatId);
  const intent = await understandIntent(env, normalizedText, history);

  try {
    if (intent.action === "create_task") {
      await createTaskFromText(env, chatId, from, normalizedText, inputType, intent);
      return;
    }

    if (intent.action === "add_note") {
      await saveNote(env, chatId, from, intent.note_text || normalizedText, inputType);
      await logAndReply(env, { chatId, from, inputType, originalText, normalizedText, action: "add_note", resultSummary: truncate(intent.note_text || normalizedText, 200), status: "ok" }, `${TEXT.NOTE_SAVED}\n\n<b>Нотатка:</b> ${escapeHtml(intent.note_text || normalizedText)}`);
      return;
    }

    if (intent.action === "add_budget") {
      await addBudgetFromText(env, chatId, from, normalizedText, inputType, intent);
      return;
    }

    if (intent.action === "add_shopping") {
      await addShoppingItem(env, chatId, from, normalizedText, intent.list_type || "buy", inputType, intent);
      return;
    }

    if (intent.action === "list_tasks") {
      await sendTaskList(env, chatId);
      return;
    }

    if (intent.action === "today_overview") {
      await sendTodayOverview(env, chatId);
      return;
    }

    if (intent.action === "weekly_report") {
      await sendWeeklyReport(env, chatId);
      return;
    }

    if (intent.action === "nearby_places") {
      await runNearbyPlaces(env, chatId, from, intent.query || normalizedText, { areaText: intent.area || "" });
      return;
    }

    if (intent.action === "search") {
      await runSearch(env, chatId, from, intent.query || normalizedText, inputType);
      return;
    }

    const answer = await answerChat(env, normalizedText, history);
    await logAndReply(env, {
      chatId,
      from,
      inputType,
      originalText,
      normalizedText,
      action: "answer",
      resultSummary: truncate(answer, 250),
      status: "ok"
    }, answer);
  } catch (error) {
    console.error("processAssistantInput error:", error instanceof Error ? error.stack || error.message : String(error));
    const message = String(error?.message || error);
    const userText = message.includes("CALENDAR_WRITE_FAILED")
      ? TEXT.ERROR_CALENDAR
      : message.includes("SHEETS_WRITE_FAILED")
        ? TEXT.ERROR_SHEETS
        : TEXT.ERROR_GENERIC;

    await logInteraction(env, {
      user: from,
      chatId,
      inputType,
      originalText,
      normalizedText,
      action: intentSafeAction(intent),
      resultSummary: truncate(message, 250),
      status: "error"
    });
    await sendMessage(env, chatId, userText, { reply_markup: mainMenuKeyboard() });
  }
}

function intentSafeAction(intent) {
  return intent?.action || "unknown";
}

async function understandIntent(env, text, history = []) {
  const timezone = env.TIMEZONE || DEFAULT_TIMEZONE;
  const historyText = history.slice(-8).map((item) => `${item.role}: ${item.text}`).join("\n");
  const prompt = [
    "Ти маршрутизатор Telegram-асистента.",
    `Часовий пояс: ${timezone}.`,
    `Поточний час ISO: ${new Date().toISOString()}.`,
    `Контекст:\n${historyText || "(порожньо)"}`,
    `Користувач сказав: ${text}`,
    "Поверни тільки валідний JSON без markdown.",
    "action може бути тільки: create_task, list_tasks, today_overview, weekly_report, add_note, add_budget, add_shopping, nearby_places, search, answer.",
    "Для create_task поверни: title, details, due_at, due_date. Якщо є конкретний час - due_at у ISO 8601 з часовим зміщенням. Якщо є тільки дата без часу - due_date у YYYY-MM-DD.",
    "Для add_budget поверни: amount, currency, type(expense|income), category, note.",
    "Для add_shopping поверни: list_type(products|buy), item_name, details, quantity.",
    "Для nearby_places поверни: query, area.",
    "Для search поверни: query.",
    "Для add_note поверни note_text.",
    "Для answer поверни answer_text.",
    "Завжди додавай confidence від 0 до 1."
  ].join("\n\n");

  const raw = await geminiText(env, {
    contents: [{ role: "user", parts: [{ text: prompt }] }],
    responseMimeType: "application/json",
    useSearch: false
  });

  try {
    const parsed = JSON.parse(raw);
    if (!parsed.action) parsed.action = "answer";
    return parsed;
  } catch {
    return { action: "answer", answer_text: text, confidence: 0.2 };
  }
}

async function createTaskFromText(env, chatId, from, rawText, inputType = "text", parsedIntent = null) {
  const intent = parsedIntent || await understandIntent(env, rawText, await getHistory(env, chatId));
  const title = (intent.title || rawText).trim();
  const taskId = generateId("TASK");
  const now = nowLocalString(env);
  const dueAt = normalizeIsoOrEmpty(intent.due_at);
  const dueDate = normalizeDateOrEmpty(intent.due_date, dueAt);
  let calendarEventId = "";

  if (dueAt || dueDate) {
    try {
      const event = await calendarInsertEvent(env, buildCalendarEvent(title, intent.details || "", dueAt, dueDate, from, env));
      calendarEventId = event.id || "";
    } catch (error) {
      throw new Error(`CALENDAR_WRITE_FAILED: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  const row = [
    taskId,
    now,
    now,
    String(chatId || ""),
    String(from?.id || ""),
    usernameLabel(from),
    title,
    intent.details || "",
    dueAt,
    dueDate,
    TASK_STATUS.NEW,
    calendarEventId,
    rawText
  ];

  try {
    await ensureWorksheet(env, SHEETS.TASKS, TASK_HEADERS);
    await sheetsAppend(env, `${SHEETS.TASKS}!A:M`, [row]);
  } catch (error) {
    throw new Error(`SHEETS_WRITE_FAILED: ${error instanceof Error ? error.message : String(error)}`);
  }

  const reply = [
    TEXT.TASK_SAVED,
    "",
    `<b>${escapeHtml(title)}</b>`,
    dueAt ? `🕒 ${escapeHtml(formatDateTime(dueAt, env.TIMEZONE || DEFAULT_TIMEZONE))}` : "",
    !dueAt && dueDate ? `📅 ${escapeHtml(dueDate)}` : "",
    calendarEventId ? "📌 Подію також додано в Google Calendar." : ""
  ].filter(Boolean).join("\n");

  await logAndReply(env, {
    chatId,
    from,
    inputType,
    originalText: rawText,
    normalizedText: rawText,
    action: "create_task",
    resultSummary: truncate(title, 200),
    calendarEventId,
    status: "ok"
  }, reply);
}

function buildCalendarEvent(title, details, dueAt, dueDate, user, env) {
  const timezone = env.TIMEZONE || DEFAULT_TIMEZONE;
  const description = [details, user?.username ? `Telegram: @${user.username}` : "", user?.id ? `Telegram ID: ${user.id}` : ""]
    .filter(Boolean)
    .join("\n");

  const event = {
    summary: title,
    description,
    reminders: {
      useDefault: false,
      overrides: [{ method: "popup", minutes: 30 }]
    }
  };

  if (dueAt) {
    event.start = { dateTime: dueAt, timeZone: timezone };
    event.end = { dateTime: addHourKeepingOffset(dueAt, 1), timeZone: timezone };
  } else {
    const date = dueDate || new Date().toISOString().slice(0, 10);
    event.start = { date, timeZone: timezone };
    event.end = { date: addDaysToIsoDate(date, 1), timeZone: timezone };
  }
  return event;
}

async function addBudgetFromText(env, chatId, from, rawText, inputType = "text", parsedIntent = null) {
  const intent = parsedIntent || await understandIntent(env, rawText, await getHistory(env, chatId));
  const entryId = generateId("BUD");
  const now = nowLocalString(env);
  const amount = normalizeAmount(intent.amount);
  const type = intent.type === "income" ? "Дохід" : "Витрата";
  const currency = intent.currency || "грн";
  const category = intent.category || "Інше";
  const note = intent.note || rawText;

  const row = [
    entryId,
    now,
    now,
    String(chatId || ""),
    String(from?.id || ""),
    usernameLabel(from),
    type,
    amount,
    currency,
    category,
    note,
    rawText
  ];

  try {
    await ensureWorksheet(env, SHEETS.BUDGET, BUDGET_HEADERS);
    await sheetsAppend(env, `${SHEETS.BUDGET}!A:L`, [row]);
  } catch (error) {
    throw new Error(`SHEETS_WRITE_FAILED: ${error instanceof Error ? error.message : String(error)}`);
  }

  const reply = `${TEXT.BUDGET_SAVED}\n\n<b>${escapeHtml(type)}</b>: ${escapeHtml(String(amount))} ${escapeHtml(currency)}\n${escapeHtml(category)}${note ? `\n${escapeHtml(note)}` : ""}`;
  await logAndReply(env, {
    chatId,
    from,
    inputType,
    originalText: rawText,
    normalizedText: rawText,
    action: "add_budget",
    resultSummary: `${type} ${amount} ${currency}`,
    status: "ok"
  }, reply);
}

async function addShoppingItem(env, chatId, from, rawText, listType = "buy", inputType = "text", parsedIntent = null) {
  const intent = parsedIntent || await understandIntent(env, rawText, await getHistory(env, chatId));
  const itemId = generateId("BUY");
  const now = nowLocalString(env);
  const row = [
    itemId,
    now,
    now,
    String(chatId || ""),
    String(from?.id || ""),
    usernameLabel(from),
    listType === "products" ? "Продукти" : "Покупки",
    intent.item_name || rawText,
    intent.details || "",
    intent.quantity || "",
    SHOPPING_STATUS.OPEN,
    rawText
  ];

  try {
    await ensureWorksheet(env, SHEETS.SHOPPING, SHOPPING_HEADERS);
    await sheetsAppend(env, `${SHEETS.SHOPPING}!A:L`, [row]);
  } catch (error) {
    throw new Error(`SHEETS_WRITE_FAILED: ${error instanceof Error ? error.message : String(error)}`);
  }

  const reply = `${TEXT.SHOPPING_SAVED}\n\n<b>${escapeHtml(intent.item_name || rawText)}</b>${intent.quantity ? `\nКількість: ${escapeHtml(String(intent.quantity))}` : ""}${intent.details ? `\n${escapeHtml(intent.details)}` : ""}`;
  await logAndReply(env, {
    chatId,
    from,
    inputType,
    originalText: rawText,
    normalizedText: rawText,
    action: "add_shopping",
    resultSummary: truncate(intent.item_name || rawText, 200),
    status: "ok"
  }, reply);
}

async function saveNote(env, chatId, user, text, source) {
  if (!text?.trim()) return;
  const row = [
    generateId("NOTE"),
    nowLocalString(env),
    String(chatId || ""),
    String(user?.id || ""),
    usernameLabel(user),
    text.trim(),
    source || "text"
  ];
  try {
    await ensureWorksheet(env, SHEETS.NOTES, NOTE_HEADERS);
    await sheetsAppend(env, `${SHEETS.NOTES}!A:G`, [row]);
  } catch (error) {
    throw new Error(`SHEETS_WRITE_FAILED: ${error instanceof Error ? error.message : String(error)}`);
  }
}

async function sendTaskList(env, chatId) {
  const tasks = await getTasks(env, chatId);
  const active = tasks.filter((t) => t.status !== TASK_STATUS.DONE && t.status !== TASK_STATUS.CANCELED);
  const overdue = active.filter((t) => isTaskOverdue(t));
  const today = active.filter((t) => isTaskToday(t, env));
  const done = tasks.filter((t) => t.status === TASK_STATUS.DONE).slice(-5).reverse();

  const lines = ["<b>Задачі</b>"];
  lines.push(`\n• Активні: <b>${active.length}</b>`);
  lines.push(`• На сьогодні: <b>${today.length}</b>`);
  lines.push(`• Прострочені: <b>${overdue.length}</b>`);

  if (today.length) {
    lines.push("\n<b>Сьогодні:</b>");
    for (const task of today.slice(0, 8)) lines.push(formatTaskLine(task, env));
  }

  if (overdue.length) {
    lines.push("\n<b>Прострочені:</b>");
    for (const task of overdue.slice(0, 6)) lines.push(formatTaskLine(task, env));
  }

  if (done.length) {
    lines.push("\n<b>Останні виконані:</b>");
    for (const task of done.slice(0, 5)) lines.push(formatTaskLine(task, env));
  }

  await sendLongMessage(env, chatId, lines.join("\n"), { reply_markup: mainMenuKeyboard() });
}

async function prepareCompleteToday(env, chatId) {
  const tasks = (await getTasks(env, chatId)).filter((t) => t.status !== TASK_STATUS.DONE && t.status !== TASK_STATUS.CANCELED && isTaskToday(t, env));
  if (!tasks.length) {
    await sendMessage(env, chatId, "На сьогодні активних задач немає.", { reply_markup: mainMenuKeyboard() });
    return;
  }

  const map = {};
  const lines = ["<b>Що виконати сьогодні:</b>"];
  tasks.slice(0, 15).forEach((task, index) => {
    const n = String(index + 1);
    map[n] = task.task_id;
    lines.push(`\n${n}. ${formatTaskLine(task, env)}`);
  });
  lines.push(`\n${TEXT.COMPLETE_PICK_PROMPT}`);
  await saveChatState(env, chatId, { step: STATES.WAIT_COMPLETE_TODAY, selectionMap: map });
  await sendLongMessage(env, chatId, lines.join("\n"), { reply_markup: mainMenuKeyboard() });
}

async function completeTaskFromSelection(env, chatId, rawText) {
  const state = await getChatState(env, chatId);
  const taskId = state?.selectionMap?.[rawText.trim()];
  if (!taskId) {
    await sendMessage(env, chatId, TEXT.COMPLETE_NOT_FOUND, { reply_markup: mainMenuKeyboard() });
    return;
  }
  const updated = await updateTaskStatus(env, taskId, TASK_STATUS.DONE);
  if (!updated) {
    await sendMessage(env, chatId, TEXT.COMPLETE_NOT_FOUND, { reply_markup: mainMenuKeyboard() });
    return;
  }
  await sendMessage(env, chatId, TEXT.COMPLETE_DONE, { reply_markup: mainMenuKeyboard() });
}

async function sendTodayOverview(env, chatId) {
  const tasks = (await getTasks(env, chatId)).filter((t) => t.status !== TASK_STATUS.DONE && t.status !== TASK_STATUS.CANCELED && isTaskToday(t, env));
  const range = resolveListRange("today", env);
  let events = [];
  try {
    events = await listCalendarEvents(env, range.timeMin, range.timeMax, 10);
  } catch (error) {
    console.error("sendTodayOverview calendar error:", error instanceof Error ? error.message : String(error));
  }

  const shopping = (await getShoppingItems(env, chatId)).filter((i) => i.status !== SHOPPING_STATUS.DONE && i.list_type === "Продукти").slice(0, 8);

  if (!tasks.length && !events.length && !shopping.length) {
    await sendMessage(env, chatId, TEXT.TODAY_EMPTY, { reply_markup: mainMenuKeyboard() });
    return;
  }

  const lines = ["<b>План на сьогодні</b>"];
  if (tasks.length) {
    lines.push("\n<b>Задачі:</b>");
    for (const task of tasks.slice(0, 10)) lines.push(formatTaskLine(task, env));
  }
  if (events.length) {
    lines.push("\n<b>Календар:</b>");
    for (const item of events.slice(0, 10)) {
      const when = item.start?.dateTime ? formatDateTime(item.start.dateTime, env.TIMEZONE || DEFAULT_TIMEZONE) : `${item.start?.date || "-"} (весь день)`;
      lines.push(`• <b>${escapeHtml(item.summary || "Без назви")}</b> — ${escapeHtml(when)}`);
    }
  }
  if (shopping.length) {
    lines.push("\n<b>Продукти:</b>");
    for (const item of shopping) lines.push(`• ${escapeHtml(item.item_name || "-")}${item.quantity ? ` — ${escapeHtml(item.quantity)}` : ""}`);
  }

  await sendLongMessage(env, chatId, lines.join("\n"), { reply_markup: mainMenuKeyboard() });
}

async function sendWeeklyReport(env, chatId) {
  const tasks = await getTasks(env, chatId);
  const budget = await getBudgetEntries(env, chatId);
  const shopping = await getShoppingItems(env, chatId);
  const weekAgo = Date.now() - 7 * 24 * 60 * 60 * 1000;

  const recentTasks = tasks.filter((t) => parseLocalDateGuess(t.created_at) >= weekAgo);
  const completedTasks = tasks.filter((t) => t.status === TASK_STATUS.DONE && parseLocalDateGuess(t.updated_at || t.created_at) >= weekAgo);
  const movedTasks = tasks.filter((t) => t.status === TASK_STATUS.MOVED && parseLocalDateGuess(t.updated_at || t.created_at) >= weekAgo);
  const canceledTasks = tasks.filter((t) => t.status === TASK_STATUS.CANCELED && parseLocalDateGuess(t.updated_at || t.created_at) >= weekAgo);

  const recentBudget = budget.filter((b) => parseLocalDateGuess(b.created_at) >= weekAgo);
  const expenses = recentBudget.filter((b) => b.type === "Витрата").reduce((sum, item) => sum + Number(item.amount || 0), 0);
  const income = recentBudget.filter((b) => b.type === "Дохід").reduce((sum, item) => sum + Number(item.amount || 0), 0);

  const openProducts = shopping.filter((i) => i.status !== SHOPPING_STATUS.DONE && i.list_type === "Продукти").length;
  const openBuy = shopping.filter((i) => i.status !== SHOPPING_STATUS.DONE && i.list_type === "Покупки").length;

  const lines = [
    "<b>Звіт за тиждень</b>",
    `\n• Створено задач: <b>${recentTasks.length}</b>`,
    `• Виконано: <b>${completedTasks.length}</b>`,
    `• Перенесено: <b>${movedTasks.length}</b>`,
    `• Скасовано: <b>${canceledTasks.length}</b>`,
    `\n• Витрати: <b>${formatMoney(expenses)}</b>`,
    `• Доходи: <b>${formatMoney(income)}</b>`,
    `\n• У продуктах лишилось: <b>${openProducts}</b>`,
    `• У списку покупок лишилось: <b>${openBuy}</b>`
  ];

  await sendMessage(env, chatId, lines.join("\n"), { reply_markup: mainMenuKeyboard() });
}

async function runSearch(env, chatId, from, query, inputType) {
  await sendMessage(env, chatId, TEXT.SEARCH_BUSY, { reply_markup: mainMenuKeyboard() });
  const history = await getHistory(env, chatId);
  const answer = await answerWithGoogleSearch(env, query, history);
  await logAndReply(env, {
    chatId,
    from,
    inputType,
    originalText: query,
    normalizedText: query,
    action: "search",
    resultSummary: truncate(answer, 250),
    status: "ok"
  }, answer);
}

async function runNearbyPlaces(env, chatId, from, query, state = {}) {
  const area = buildNearbyAreaLabel(state);
  const prompt = [
    `Знайди до 5 реальних місць для запиту: ${query}.`,
    area ? `Локація або район: ${area}.` : "",
    "Поверни тільки валідний JSON масив без markdown.",
    "Для кожного елемента поля: name, address, why."
  ].filter(Boolean).join("\n");

  const raw = await geminiText(env, {
    contents: [{ role: "user", parts: [{ text: prompt }] }],
    responseMimeType: "application/json",
    useSearch: true
  });

  let items = [];
  try {
    items = JSON.parse(raw);
  } catch {
    items = [];
  }

  if (!Array.isArray(items) || !items.length) {
    await sendMessage(env, chatId, TEXT.NOTHING_FOUND, { reply_markup: mainMenuKeyboard() });
    return;
  }

  const lines = [`<b>Що знайшов${area ? ` для ${escapeHtml(area)}` : ""}:</b>`];
  for (const item of items.slice(0, 5)) {
    const mapQuery = [item.name, item.address, area].filter(Boolean).join(" ");
    const mapsLink = `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(mapQuery)}`;
    lines.push(`\n• <b>${escapeHtml(item.name || "Місце")}</b>`);
    if (item.address) lines.push(`${escapeHtml(item.address)}`);
    if (item.why) lines.push(`${escapeHtml(item.why)}`);
    lines.push(`${escapeHtml(mapsLink)}`);
  }

  await logAndReply(env, {
    chatId,
    from,
    inputType: "text",
    originalText: query,
    normalizedText: query,
    action: "nearby_places",
    resultSummary: truncate(`${query} ${area}`.trim(), 200),
    status: "ok"
  }, lines.join("\n"));
}

function buildNearbyAreaLabel(state) {
  if (!state) return "";
  if (state.areaText) return state.areaText;
  if (state.location?.latitude && state.location?.longitude) {
    return `${state.location.latitude}, ${state.location.longitude}`;
  }
  return "";
}

async function answerChat(env, text, history = []) {
  const contents = [];
  const systemPrompt = [
    "Ти особистий Telegram-асистент.",
    "Відповідай українською, коротко, дружньо, по суті.",
    "Не вигадуй факти. Якщо немає впевненості — прямо скажи про це.",
    "Не використовуй markdown-таблиці."
  ].join(" ");
  contents.push({ role: "user", parts: [{ text: systemPrompt }] });
  for (const item of history.slice(-8)) {
    contents.push({ role: item.role === "assistant" ? "model" : "user", parts: [{ text: item.text }] });
  }
  contents.push({ role: "user", parts: [{ text }] });
  return await geminiText(env, { contents, useSearch: false });
}

async function answerWithGoogleSearch(env, text, history = []) {
  const contents = [];
  const systemPrompt = [
    "Ти особистий Telegram-асистент.",
    "Через Google Search grounding знайди актуальну інформацію і дай коротку практичну відповідь українською.",
    "Якщо доречно — дай 3 варіанти.",
    "У кінці додай блок 'Джерела:' зі списком посилань, якщо вони є."
  ].join(" ");
  contents.push({ role: "user", parts: [{ text: systemPrompt }] });
  for (const item of history.slice(-6)) {
    contents.push({ role: item.role === "assistant" ? "model" : "user", parts: [{ text: item.text }] });
  }
  contents.push({ role: "user", parts: [{ text }] });
  return await geminiText(env, { contents, useSearch: true });
}

async function transcribeTelegramAudio(env, audio) {
  const fileInfo = await telegramApi(env, "getFile", { file_id: audio.file_id });
  const path = fileInfo.result?.file_path;
  if (!path) throw new Error("Telegram file path not found");

  const tgUrl = `https://api.telegram.org/file/bot${env.BOT_TOKEN}/${path}`;
  const response = await fetch(tgUrl);
  if (!response.ok) throw new Error(`Failed to download Telegram audio: ${response.status}`);
  const bytes = new Uint8Array(await response.arrayBuffer());
  const base64 = arrayBufferToBase64(bytes);

  const prompt = "Розшифруй це голосове українською. Поверни тільки чистий текст без пояснень.";
  return await geminiText(env, {
    contents: [{ role: "user", parts: [
      { text: prompt },
      { inline_data: { mime_type: audio.mime_type || "audio/ogg", data: base64 } }
    ] }],
    useSearch: false
  });
}

async function analyzeTelegramPhoto(env, photos, promptText) {
  const best = photos[photos.length - 1];
  const fileInfo = await telegramApi(env, "getFile", { file_id: best.file_id });
  const path = fileInfo.result?.file_path;
  if (!path) throw new Error("Telegram photo path not found");

  const tgUrl = `https://api.telegram.org/file/bot${env.BOT_TOKEN}/${path}`;
  const response = await fetch(tgUrl);
  if (!response.ok) throw new Error(`Failed to download Telegram photo: ${response.status}`);
  const bytes = new Uint8Array(await response.arrayBuffer());
  const base64 = arrayBufferToBase64(bytes);

  return await geminiText(env, {
    contents: [{ role: "user", parts: [
      { text: promptText },
      { inline_data: { mime_type: "image/jpeg", data: base64 } }
    ] }],
    useSearch: false
  });
}

function mainMenuKeyboard() {
  return {
    keyboard: [
      [{ text: MENU.CREATE_TASK }, { text: MENU.LIST_TASKS }],
      [{ text: MENU.COMPLETE_TODAY }, { text: MENU.TODAY }],
      [{ text: MENU.BUDGET }, { text: MENU.PRODUCTS }],
      [{ text: MENU.NOTE }, { text: MENU.BUY }],
      [{ text: MENU.SEARCH }, { text: MENU.NEARBY }],
      [{ text: MENU.WEEK_REPORT }]
    ],
    resize_keyboard: true
  };
}

function nearbyLocationKeyboard() {
  return {
    keyboard: [
      [{ text: "Надіслати геолокацію", request_location: true }],
      [{ text: MENU.NEARBY }, { text: MENU.TODAY }]
    ],
    resize_keyboard: true,
    one_time_keyboard: true
  };
}

async function sendLongMessage(env, chatId, text, extra = {}) {
  const chunks = splitMessage(text, MAX_TELEGRAM_MESSAGE);
  for (let i = 0; i < chunks.length; i++) {
    await sendMessage(env, chatId, chunks[i], i === chunks.length - 1 ? extra : {});
  }
}

function splitMessage(text, maxLen) {
  if (text.length <= maxLen) return [text];
  const chunks = [];
  let current = "";
  for (const line of text.split("\n")) {
    if ((current + "\n" + line).length > maxLen && current) {
      chunks.push(current);
      current = line;
    } else {
      current = current ? `${current}\n${line}` : line;
    }
  }
  if (current) chunks.push(current);
  return chunks;
}

async function sendMessage(env, chatId, text, extra = {}) {
  return await telegramApi(env, "sendMessage", {
    chat_id: chatId,
    text,
    parse_mode: "HTML",
    disable_web_page_preview: false,
    ...extra
  });
}

async function sendChatAction(env, chatId, action) {
  return await telegramApi(env, "sendChatAction", { chat_id: chatId, action });
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

async function getTasks(env, chatId) {
  await ensureWorksheet(env, SHEETS.TASKS, TASK_HEADERS);
  const rows = await sheetsGetValues(env, `${SHEETS.TASKS}!A:M`);
  return mapRows(rows).filter((row) => String(row.chat_id || "") === String(chatId || ""));
}

async function getBudgetEntries(env, chatId) {
  await ensureWorksheet(env, SHEETS.BUDGET, BUDGET_HEADERS);
  const rows = await sheetsGetValues(env, `${SHEETS.BUDGET}!A:L`);
  return mapRows(rows).filter((row) => String(row.chat_id || "") === String(chatId || ""));
}

async function getShoppingItems(env, chatId) {
  await ensureWorksheet(env, SHEETS.SHOPPING, SHOPPING_HEADERS);
  const rows = await sheetsGetValues(env, `${SHEETS.SHOPPING}!A:L`);
  return mapRows(rows).filter((row) => String(row.chat_id || "") === String(chatId || ""));
}

async function updateTaskStatus(env, taskId, newStatus) {
  await ensureWorksheet(env, SHEETS.TASKS, TASK_HEADERS);
  const values = await sheetsGetValues(env, `${SHEETS.TASKS}!A:M`);
  if (values.length <= 1) return false;
  const headers = values[0];
  const idIndex = headers.indexOf("task_id");
  const statusIndex = headers.indexOf("status");
  const updatedIndex = headers.indexOf("updated_at");
  if (idIndex === -1 || statusIndex === -1 || updatedIndex === -1) return false;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if ((row[idIndex] || "") === taskId) {
      const rowNumber = i + 1;
      await sheetsUpdateValues(env, `${SHEETS.TASKS}!${indexToColumn(statusIndex)}${rowNumber}`, [[newStatus]]);
      await sheetsUpdateValues(env, `${SHEETS.TASKS}!${indexToColumn(updatedIndex)}${rowNumber}`, [[nowLocalString(env)]]);
      return true;
    }
  }
  return false;
}

function mapRows(values) {
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

function formatTaskLine(task, env) {
  const when = task.due_at
    ? formatDateTime(task.due_at, env.TIMEZONE || DEFAULT_TIMEZONE)
    : task.due_date || "без дати";
  return `• <b>${escapeHtml(task.title || "Без назви")}</b> — ${escapeHtml(when)} <i>(${escapeHtml(task.status || TASK_STATUS.NEW)})</i>`;
}

function isTaskToday(task, env) {
  const today = localDateKey(new Date(), env.TIMEZONE || DEFAULT_TIMEZONE);
  if (task.due_date) return task.due_date === today;
  if (task.due_at) return localDateKey(new Date(task.due_at), env.TIMEZONE || DEFAULT_TIMEZONE) === today;
  return false;
}

function isTaskOverdue(task) {
  if (!task.due_at && !task.due_date) return false;
  const now = Date.now();
  if (task.due_at) return new Date(task.due_at).getTime() < now;
  return parseDateOnly(task.due_date).getTime() < parseDateOnly(new Date().toISOString().slice(0, 10)).getTime();
}

function resolveListRange(range, env) {
  const timezone = env.TIMEZONE || DEFAULT_TIMEZONE;
  const now = new Date();
  const todayStart = startOfDayInTimezone(now, timezone, 0);
  const tomorrowStart = startOfDayInTimezone(now, timezone, 1);
  const weekEnd = startOfDayInTimezone(now, timezone, 7);
  if (range === "week") return { timeMin: todayStart, timeMax: weekEnd };
  return { timeMin: todayStart, timeMax: tomorrowStart };
}

async function logAndReply(env, logEntry, replyText) {
  const { chatId } = logEntry;
  await appendHistory(env, chatId, { role: "user", text: logEntry.normalizedText || logEntry.originalText || "" });
  await appendHistory(env, chatId, { role: "assistant", text: stripHtml(replyText) });
  await logInteraction(env, logEntry);
  await sendLongMessage(env, chatId, replyText, { reply_markup: mainMenuKeyboard() });
}

async function logInteraction(env, entry) {
  try {
    await ensureWorksheet(env, SHEETS.LOG, LOG_HEADERS);
    const row = [
      nowLocalString(env),
      String(entry.chatId || ""),
      String(entry.user?.id || ""),
      usernameLabel(entry.user),
      entry.inputType || "",
      entry.originalText || "",
      entry.normalizedText || "",
      entry.action || "",
      entry.resultSummary || "",
      entry.calendarEventId || "",
      entry.status || ""
    ];
    await sheetsAppend(env, `${SHEETS.LOG}!A:K`, [row]);
  } catch (error) {
    console.error("logInteraction failed:", error instanceof Error ? error.message : String(error));
  }
}

async function isDuplicateUpdate(env, updateId) {
  return (await env.STATE_KV.get(`update:${updateId}`)) === "1";
}

async function markUpdateProcessed(env, updateId) {
  await env.STATE_KV.put(`update:${updateId}`, "1", { expirationTtl: 60 * 30 });
}

async function getChatState(env, chatId) {
  const raw = await env.STATE_KV.get(`assistant:state:${chatId}`);
  return raw ? JSON.parse(raw) : null;
}

async function saveChatState(env, chatId, state) {
  await env.STATE_KV.put(`assistant:state:${chatId}`, JSON.stringify(state), { expirationTtl: 60 * 60 * 24 * 14 });
}

async function clearChatState(env, chatId) {
  await env.STATE_KV.delete(`assistant:state:${chatId}`);
}

async function getHistory(env, chatId) {
  const raw = await env.STATE_KV.get(`assistant:history:${chatId}`);
  return raw ? JSON.parse(raw) : [];
}

async function appendHistory(env, chatId, item) {
  const history = await getHistory(env, chatId);
  history.push({ role: item.role, text: String(item.text || "") });
  while (history.length > MAX_HISTORY_MESSAGES) history.shift();
  await env.STATE_KV.put(`assistant:history:${chatId}`, JSON.stringify(history), { expirationTtl: 60 * 60 * 24 * 30 });
}

async function clearHistory(env, chatId) {
  await env.STATE_KV.delete(`assistant:history:${chatId}`);
}

async function ensureWorksheet(env, title, headers) {
  const metadata = await sheetsGetMetadata(env);
  const found = metadata.sheets?.find((sheet) => sheet.properties?.title === title);
  if (!found) {
    await sheetsBatchUpdate(env, { requests: [{ addSheet: { properties: { title } } }] });
  }

  const range = `${title}!A1:${indexToColumn(headers.length - 1)}1`;
  const values = await sheetsGetValues(env, range);
  const current = values[0] || [];
  if (!current.length) {
    await sheetsUpdateValues(env, range, [headers]);
    return;
  }

  const merged = [...current];
  let changed = false;
  for (let i = 0; i < headers.length; i++) {
    if (merged[i] !== headers[i]) {
      merged[i] = headers[i];
      changed = true;
    }
  }
  if (changed) {
    await sheetsUpdateValues(env, `${title}!A1:${indexToColumn(merged.length - 1)}1`, [merged]);
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

async function calendarInsertEvent(env, event) {
  return await googleApi(
    env,
    `https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(env.CALENDAR_ID || DEFAULT_CALENDAR_ID)}/events`,
    { method: "POST", body: event }
  );
}

async function listCalendarEvents(env, timeMin, timeMax, maxResults = 10) {
  const params = new URLSearchParams({
    timeMin,
    timeMax,
    singleEvents: "true",
    orderBy: "startTime",
    maxResults: String(maxResults)
  });
  const data = await googleApi(
    env,
    `https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(env.CALENDAR_ID || DEFAULT_CALENDAR_ID)}/events?${params.toString()}`
  );
  return data.items || [];
}

let cachedGoogleToken = null;

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

async function getGoogleAccessToken(env) {
  if (cachedGoogleToken && cachedGoogleToken.expiresAt > Date.now() + 60_000) {
    return cachedGoogleToken.token;
  }

  const serviceAccount = JSON.parse(env.GOOGLE_SERVICE_ACCOUNT_JSON);
  const now = Math.floor(Date.now() / 1000);
  const jwtHeader = { alg: "RS256", typ: "JWT" };
  const jwtClaimSet = {
    iss: serviceAccount.client_email,
    scope: [
      "https://www.googleapis.com/auth/spreadsheets",
      "https://www.googleapis.com/auth/calendar"
    ].join(" "),
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

async function geminiText(env, { contents, responseMimeType = "text/plain", useSearch = false }) {
  const model = env.GEMINI_MODEL || DEFAULT_MODEL;
  const body = {
    contents,
    generationConfig: {
      temperature: 0.2,
      responseMimeType
    }
  };
  if (useSearch) {
    body.tools = [{ google_search: {} }];
  }

  const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent`, {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-goog-api-key": env.GEMINI_API_KEY
    },
    body: JSON.stringify(body)
  });

  const data = await response.json();
  if (!response.ok) {
    throw new Error(`Gemini API error: ${JSON.stringify(data)}`);
  }

  const candidate = data.candidates?.[0];
  const parts = candidate?.content?.parts || [];
  const text = parts.map((part) => part.text || "").join("").trim();
  if (!text) {
    throw new Error(`Gemini empty response: ${JSON.stringify(data)}`);
  }

  if (useSearch && responseMimeType !== "application/json") {
    const sources = extractGroundingUris(data);
    if (sources.length) {
      return `${text}\n\n<b>Джерела:</b>\n${sources.slice(0, 5).map((url) => `• ${escapeHtml(url)}`).join("\n")}`;
    }
  }
  return text;
}

function extractGroundingUris(data) {
  const candidates = data.candidates || [];
  const uris = [];
  for (const candidate of candidates) {
    const chunks = candidate.groundingMetadata?.groundingChunks || [];
    for (const chunk of chunks) {
      const uri = chunk?.web?.uri;
      if (uri && !uris.includes(uri)) uris.push(uri);
    }
  }
  return uris;
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

function arrayBufferToBase64(buffer) {
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
}

function formatDateTime(iso, timeZone = DEFAULT_TIMEZONE) {
  try {
    return new Intl.DateTimeFormat("uk-UA", {
      timeZone,
      dateStyle: "short",
      timeStyle: "short"
    }).format(new Date(iso));
  } catch {
    return iso;
  }
}

function localDateKey(date, timeZone) {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  }).formatToParts(date);
  const year = parts.find((p) => p.type === "year")?.value;
  const month = parts.find((p) => p.type === "month")?.value;
  const day = parts.find((p) => p.type === "day")?.value;
  return `${year}-${month}-${day}`;
}

function nowLocalString(env) {
  const timezone = env.TIMEZONE || DEFAULT_TIMEZONE;
  return new Intl.DateTimeFormat("uk-UA", {
    timeZone: timezone,
    dateStyle: "short",
    timeStyle: "medium"
  }).format(new Date());
}

function startOfDayInTimezone(baseDate, timeZone, plusDays = 0) {
  const target = new Date(baseDate.getTime() + plusDays * 24 * 60 * 60 * 1000);
  const day = localDateKey(target, timeZone);
  const offset = getTimeZoneOffsetString(timeZone, target);
  return `${day}T00:00:00${offset}`;
}

function getTimeZoneOffsetString(timeZone, date = new Date()) {
  try {
    const parts = new Intl.DateTimeFormat("en-US", {
      timeZone,
      timeZoneName: "longOffset",
      hour: "2-digit"
    }).formatToParts(date);
    const tz = parts.find((p) => p.type === "timeZoneName")?.value || "GMT+00:00";
    return tz.replace("GMT", "") || "+00:00";
  } catch {
    return "+00:00";
  }
}

function addHourKeepingOffset(iso, hours) {
  const date = new Date(iso);
  date.setHours(date.getHours() + hours);
  return date.toISOString();
}

function addDaysToIsoDate(dateStr, days) {
  const date = new Date(`${dateStr}T00:00:00Z`);
  date.setUTCDate(date.getUTCDate() + days);
  return date.toISOString().slice(0, 10);
}

function parseDateOnly(dateStr) {
  return new Date(`${dateStr}T00:00:00`);
}

function parseLocalDateGuess(value) {
  if (!value) return 0;
  const direct = Date.parse(value);
  if (!Number.isNaN(direct)) return direct;
  const match = String(value).match(/(\d{2})\.(\d{2})\.(\d{2,4}),?\s*(\d{2}:\d{2}:\d{2})?/);
  if (match) {
    const year = match[3].length === 2 ? `20${match[3]}` : match[3];
    return Date.parse(`${year}-${match[2]}-${match[1]}T${match[4] || "00:00:00"}`);
  }
  return 0;
}

function normalizeDateFromLocalString(value) {
  const ts = parseLocalDateGuess(value);
  if (!ts) return "";
  return new Date(ts).toISOString().slice(0, 10);
}

function normalizeAmount(value) {
  if (value === undefined || value === null || value === "") return "";
  const clean = String(value).replace(/,/g, ".").replace(/[^\d.\-]/g, "");
  const num = Number(clean);
  return Number.isFinite(num) ? num : "";
}

function normalizeIsoOrEmpty(value) {
  if (!value) return "";
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? "" : value;
}

function normalizeDateOrEmpty(dueDate, fallbackIso) {
  if (dueDate && /^\d{4}-\d{2}-\d{2}$/.test(dueDate)) return dueDate;
  if (fallbackIso) return fallbackIso.slice(0, 10);
  return "";
}

function generateId(prefix) {
  const rand = Math.random().toString(36).slice(2, 7).toUpperCase();
  return `${prefix}-${Date.now()}-${rand}`;
}

function usernameLabel(user) {
  if (user?.username) return `@${user.username}`;
  return user?.first_name || "";
}

function normalizeButtonText(value) {
  return String(value || "").replace(/\s+/g, " ").trim().toLowerCase();
}

function formatMoney(value) {
  return `${Number(value || 0).toFixed(2)} грн`;
}

function truncate(value, max) {
  const str = String(value || "");
  return str.length > max ? `${str.slice(0, max - 1)}…` : str;
}

function stripHtml(value) {
  return String(value || "").replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
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

function validateEnv(env) {
  const required = ["BOT_TOKEN", "GEMINI_API_KEY", "GOOGLE_SHEET_ID", "GOOGLE_SERVICE_ACCOUNT_JSON", "STATE_KV"];
  for (const key of required) {
    if (!env[key]) throw new Error(`Missing environment variable: ${key}`);
  }
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
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}
