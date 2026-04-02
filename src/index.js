const DEFAULT_TIMEZONE = "Europe/Kyiv";
const DEFAULT_CALENDAR_ID = "techno.perspektiva@gmail.com";
const DEFAULT_WORKSHEET_NAME = "AssistantLog";
const DEFAULT_NOTES_WORKSHEET_NAME = "AssistantNotes";
const DEFAULT_MODEL = "gemini-2.5-flash";
const MAX_HISTORY_MESSAGES = 12;

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

const NOTE_HEADERS = [
  "created_at",
  "chat_id",
  "user_id",
  "username",
  "note_text",
  "source"
];

const TEXT = {
  START:
    "<b>Привіт. Я твій Telegram-асистент.</b>\n\nЩо я вмію:\n• створювати події в Google Calendar\n• зберігати нотатки в Google Sheets\n• розуміти голосові\n• шукати і рекомендувати потрібні речі\n• аналізувати фото\n\nПросто напиши або надішли голосове: \n<i>Нагадай сьогодні о 19:00 тренування</i>",
  HELP:
    "<b>Приклади:</b>\n\n• Нагадай сьогодні о 19:00 тренування\n• Запиши зустріч завтра о 14:30 з Ігорем\n• Занотуй: купити кабель і гофру\n• Що у мене сьогодні по календарю\n• Знайди мені хороший спортзал біля Позняків\n• Проаналізуй це фото\n\n<b>Команди:</b>\n/start\n/help\n/today\n/tomorrow\n/week\n/events\n/note текст\n/reset",
  RESET: "Контекст чату очищено.",
  NOTE_PROMPT: "Надішли наступним повідомленням текст нотатки.",
  NOTE_SAVED: "Готово, занотував у Google Sheets.",
  CAL_CREATED: "Подію створено в Google Calendar.",
  CAL_EMPTY: "На цей період подій не знайшов.",
  FALLBACK: "Не вдалося обробити запит. Спробуй перефразувати коротше.",
  VOICE_BUSY: "Отримав голосове. Розшифровую і обробляю...",
  PHOTO_BUSY: "Фото отримав. Дивлюся, що на ньому.",
  SEARCH_BUSY: "Шукаю і збираю коротку відповідь...",
  ERROR_CALENDAR:
    "Не зміг записати подію в календар. Перевір доступ календаря для service account або спробуй ще раз.",
  ERROR_SHEETS:
    "Не зміг записати в Google Sheets. Перевір доступ таблиці для service account.",
};

export default {
  async fetch(request, env, ctx) {
    try {
      const url = new URL(request.url);

      if (request.method === "GET" && url.pathname === "/") {
        return jsonResponse({ ok: true, service: "personal-assistant", status: "running" });
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
      { command: "help", description: "Що я вмію" },
      { command: "today", description: "Події на сьогодні" },
      { command: "tomorrow", description: "Події на завтра" },
      { command: "week", description: "Події на 7 днів" },
      { command: "events", description: "Найближчі події" },
      { command: "note", description: "Зберегти нотатку" },
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
  const lowerText = rawText.toLowerCase();

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

  if (rawText === "/today" || lowerText === "сьогодні") {
    await sendUpcomingEvents(env, chatId, 0, 1, "сьогодні");
    return;
  }

  if (rawText === "/tomorrow" || lowerText === "завтра") {
    await sendUpcomingEvents(env, chatId, 1, 1, "завтра");
    return;
  }

  if (rawText === "/week" || lowerText === "тиждень") {
    await sendUpcomingEvents(env, chatId, 0, 7, "на 7 днів");
    return;
  }

  if (rawText === "/events" || lowerText === "події") {
    await sendUpcomingEvents(env, chatId, 0, 14, "найближчі події", 10);
    return;
  }

  if (rawText === "/note") {
    await saveChatState(env, chatId, { step: "waiting_note" });
    await sendMessage(env, chatId, TEXT.NOTE_PROMPT, { reply_markup: mainMenuKeyboard() });
    return;
  }

  if (rawText.startsWith("/note ")) {
    const noteText = rawText.slice(6).trim();
    await saveNote(env, chatId, from, noteText, "command");
    await sendMessage(env, chatId, TEXT.NOTE_SAVED, { reply_markup: mainMenuKeyboard() });
    return;
  }

  const state = await getChatState(env, chatId);
  if (state?.step === "waiting_note" && rawText) {
    await saveNote(env, chatId, from, rawText, "manual");
    await clearChatState(env, chatId);
    await sendMessage(env, chatId, TEXT.NOTE_SAVED, { reply_markup: mainMenuKeyboard() });
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
      originalText: rawText,
      normalizedText: transcript,
      message,
    });
    return;
  }

  if (message.photo?.length) {
    await sendChatAction(env, chatId, "typing");
    await sendMessage(env, chatId, TEXT.PHOTO_BUSY, { reply_markup: mainMenuKeyboard() });
    const analysis = await analyzeTelegramPhoto(env, message.photo, rawText || "Опиши, що на фото, і якщо доречно дай коротку практичну пораду.");
    await appendHistory(env, chatId, { role: "user", text: rawText || "[photo]" });
    await appendHistory(env, chatId, { role: "assistant", text: analysis });
    await logInteraction(env, {
      user: from,
      chatId,
      inputType: "photo",
      originalText: rawText || "[photo]",
      normalizedText: rawText || "[photo]",
      action: "analyze_photo",
      resultSummary: truncate(analysis, 400),
      status: "ok"
    });
    await sendMessage(env, chatId, analysis, { reply_markup: mainMenuKeyboard() });
    return;
  }

  if (rawText) {
    await processAssistantInput(env, {
      chatId,
      from,
      inputType: "text",
      originalText: rawText,
      normalizedText: rawText,
      message,
    });
    return;
  }

  await sendMessage(env, chatId, TEXT.FALLBACK, { reply_markup: mainMenuKeyboard() });
}

async function processAssistantInput(env, payload) {
  const { chatId, from, inputType, originalText, normalizedText } = payload;
  await sendChatAction(env, chatId, "typing");

  const history = await getHistory(env, chatId);
  const intent = await understandIntent(env, normalizedText, history);

  let finalText = TEXT.FALLBACK;
  let action = intent.action || "answer";
  let resultSummary = "";
  let calendarEventId = "";

  try {
    if (action === "add_note") {
      const noteText = intent.note_text || normalizedText;
      await saveNote(env, chatId, from, noteText, inputType);
      finalText = `${TEXT.NOTE_SAVED}\n\n<b>Нотатка:</b> ${escapeHtml(noteText)}`;
      resultSummary = noteText;
    } else if (action === "create_event") {
      const created = await createCalendarEventFromIntent(env, intent, normalizedText, from);
      calendarEventId = created.id || "";
      finalText = `${TEXT.CAL_CREATED}\n\n${formatCreatedEvent(created, env)}`;
      resultSummary = `${created.summary || "подія"} ${created.start?.dateTime || created.start?.date || ""}`.trim();
    } else if (action === "list_events") {
      const range = resolveListRange(intent.range || "today", env);
      const events = await listCalendarEvents(env, range.timeMin, range.timeMax, 10);
      finalText = formatEventsText(events, range.label);
      resultSummary = `events:${events.length}`;
    } else if (action === "search") {
      await sendMessage(env, chatId, TEXT.SEARCH_BUSY, { reply_markup: mainMenuKeyboard() });
      finalText = await answerWithGoogleSearch(env, normalizedText, history);
      resultSummary = truncate(finalText, 300);
    } else {
      finalText = await answerChat(env, normalizedText, history);
      resultSummary = truncate(finalText, 300);
    }

    await appendHistory(env, chatId, { role: "user", text: normalizedText });
    await appendHistory(env, chatId, { role: "assistant", text: stripHtml(finalText) });
    await logInteraction(env, {
      user: from,
      chatId,
      inputType,
      originalText,
      normalizedText,
      action,
      resultSummary,
      calendarEventId,
      status: "ok"
    });
    await sendMessage(env, chatId, finalText, { reply_markup: mainMenuKeyboard() });
  } catch (error) {
    console.error("processAssistantInput error:", error instanceof Error ? error.stack || error.message : String(error));
    const message = String(error?.message || error);
    let userText = TEXT.FALLBACK;
    if (message.includes("CALENDAR_WRITE_FAILED")) userText = TEXT.ERROR_CALENDAR;
    if (message.includes("SHEETS_WRITE_FAILED")) userText = TEXT.ERROR_SHEETS;

    await logInteraction(env, {
      user: from,
      chatId,
      inputType,
      originalText,
      normalizedText,
      action,
      resultSummary: truncate(message, 400),
      calendarEventId,
      status: "error"
    });
    await sendMessage(env, chatId, userText, { reply_markup: mainMenuKeyboard() });
  }
}

async function understandIntent(env, text, history = []) {
  const now = new Date();
  const timezone = env.TIMEZONE || DEFAULT_TIMEZONE;
  const historyText = history.map((item) => `${item.role}: ${item.text}`).join("\n");
  const prompt = [
    `Ти маршрутизатор Telegram-асистента.`,
    `Часовий пояс користувача: ${timezone}.`,
    `Поточний ISO час: ${now.toISOString()}.`,
    `Останній контекст чату:\n${historyText || "(порожньо)"}`,
    `Користувач сказав: ${text}`,
    `Поверни лише валідний JSON без markdown.`,
    `action може бути тільки: create_event, list_events, add_note, search, answer.`,
    `Якщо користувач просить записати зустріч/нагадування/тренування/дзвінок/подію - create_event.`,
    `Для create_event поверни: title, start_iso, end_iso, all_day, note_text. start_iso/end_iso у форматі ISO 8601 з часовим зміщенням. Якщо кінець неочевидний - постав +1 година.`,
    `Для list_events поверни range: today | tomorrow | week | upcoming.`,
    `Для add_note поверни note_text.`,
    `Для search поверни query.`,
    `Для answer поверни answer_text коротко.`,
    `Завжди додавай поле confidence від 0 до 1.`
  ].join("\n\n");

  const raw = await geminiText(env, {
    contents: [{ role: "user", parts: [{ text: prompt }] }],
    responseMimeType: "application/json",
    useSearch: false,
  });

  try {
    const parsed = JSON.parse(raw);
    if (!parsed.action) parsed.action = "answer";
    return parsed;
  } catch {
    return { action: "answer", answer_text: text, confidence: 0.2 };
  }
}

async function answerChat(env, text, history = []) {
  const contents = [];
  const systemPrompt = [
    "Ти особистий Telegram-асистент користувача.",
    "Відповідай українською, коротко, по суті, дружньо.",
    "Не вигадуй факти. Якщо для відповіді потрібна актуальна інформація з інтернету, краще використай search інструмент, але зараз інструмент не викликай сам — це вже зроблено на етапі роутингу.",
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
    "Знайди актуальну інформацію через Google Search grounding і дай коротку практичну відповідь українською.",
    "Коли доречно, подай 3 варіанти або 3 кроки.",
    "Наприкінці додай блок 'Джерела:' і короткий список посилань, якщо вони є в grounding.",
  ].join(" ");

  contents.push({ role: "user", parts: [{ text: systemPrompt }] });
  for (const item of history.slice(-6)) {
    contents.push({ role: item.role === "assistant" ? "model" : "user", parts: [{ text: item.text }] });
  }
  contents.push({ role: "user", parts: [{ text }] });

  return await geminiText(env, { contents, useSearch: true });
}

async function transcribeTelegramAudio(env, audio) {
  const fileId = audio.file_id;
  const fileInfo = await telegramApi(env, "getFile", { file_id: fileId });
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
    useSearch: false,
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
    useSearch: false,
  });
}

async function createCalendarEventFromIntent(env, intent, fallbackText, user) {
  const timezone = env.TIMEZONE || DEFAULT_TIMEZONE;
  const title = intent.title || fallbackText;
  const descriptionParts = [];
  if (intent.note_text) descriptionParts.push(intent.note_text);
  if (user?.username) descriptionParts.push(`Telegram: @${user.username}`);
  if (user?.id) descriptionParts.push(`Telegram ID: ${user.id}`);

  const event = {
    summary: title,
    description: descriptionParts.join("\n"),
    reminders: {
      useDefault: false,
      overrides: [{ method: "popup", minutes: 30 }]
    }
  };

  if (intent.all_day) {
    const startDate = (intent.start_iso || new Date().toISOString()).slice(0, 10);
    const endDate = addDaysToIsoDate(startDate, 1);
    event.start = { date: startDate, timeZone: timezone };
    event.end = { date: endDate, timeZone: timezone };
  } else {
    if (!intent.start_iso) throw new Error("CALENDAR_WRITE_FAILED: start_iso missing");
    event.start = { dateTime: intent.start_iso, timeZone: timezone };
    event.end = { dateTime: intent.end_iso || addHour(intent.start_iso, 1), timeZone: timezone };
  }

  try {
    return await calendarInsertEvent(env, event);
  } catch (error) {
    throw new Error(`CALENDAR_WRITE_FAILED: ${error instanceof Error ? error.message : String(error)}`);
  }
}

function resolveListRange(range, env) {
  const timezone = env.TIMEZONE || DEFAULT_TIMEZONE;
  const now = new Date();
  const todayStart = startOfDayInTimezone(now, timezone, 0);
  const tomorrowStart = startOfDayInTimezone(now, timezone, 1);
  const weekEnd = startOfDayInTimezone(now, timezone, 7);
  const dayAfterTomorrow = startOfDayInTimezone(now, timezone, 2);

  if (range === "tomorrow") {
    return { timeMin: tomorrowStart, timeMax: dayAfterTomorrow, label: "завтра" };
  }
  if (range === "week") {
    return { timeMin: todayStart, timeMax: weekEnd, label: "на 7 днів" };
  }
  if (range === "upcoming") {
    return { timeMin: new Date().toISOString(), timeMax: weekEnd, label: "найближчі події" };
  }
  return { timeMin: todayStart, timeMax: tomorrowStart, label: "сьогодні" };
}

async function sendUpcomingEvents(env, chatId, startOffsetDays, spanDays, label, maxResults = 10) {
  const timezone = env.TIMEZONE || DEFAULT_TIMEZONE;
  const timeMin = startOfDayInTimezone(new Date(), timezone, startOffsetDays);
  const timeMax = startOfDayInTimezone(new Date(), timezone, startOffsetDays + spanDays);
  const events = await listCalendarEvents(env, timeMin, timeMax, maxResults);
  await sendMessage(env, chatId, formatEventsText(events, label), { reply_markup: mainMenuKeyboard() });
}

function formatEventsText(events, label) {
  if (!events.length) return `${TEXT.CAL_EMPTY}\n\nПеріод: <b>${escapeHtml(label)}</b>`;

  const lines = [`<b>Події ${escapeHtml(label)}:</b>`];
  for (const item of events) {
    const when = item.start?.dateTime
      ? formatDateTime(item.start.dateTime)
      : `${item.start?.date || "-"} (весь день)`;
    lines.push(`\n• <b>${escapeHtml(item.summary || "Без назви")}</b>\n${escapeHtml(when)}`);
  }
  return lines.join("\n");
}

function formatCreatedEvent(event) {
  const when = event.start?.dateTime
    ? formatDateTime(event.start.dateTime)
    : `${event.start?.date || "-"} (весь день)`;
  return [
    `<b>${escapeHtml(event.summary || "Подія")}</b>`,
    `🗓 ${escapeHtml(when)}`,
    event.htmlLink ? `🔗 ${escapeHtml(event.htmlLink)}` : ""
  ].filter(Boolean).join("\n");
}

function mainMenuKeyboard() {
  return {
    keyboard: [
      [{ text: "Сьогодні" }, { text: "Завтра" }],
      [{ text: "Тиждень" }, { text: "Події" }]
    ],
    resize_keyboard: true
  };
}

async function sendMessage(env, chatId, text, extra = {}) {
  return await telegramApi(env, "sendMessage", {
    chat_id: chatId,
    text,
    parse_mode: "HTML",
    disable_web_page_preview: false,
    ...extra,
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

async function saveNote(env, chatId, user, text, source) {
  if (!text?.trim()) return;
  const row = [
    nowLocalString(env),
    String(chatId || ""),
    String(user?.id || ""),
    user?.username ? `@${user.username}` : (user?.first_name || ""),
    text.trim(),
    source || "text",
  ];
  try {
    await ensureWorksheet(env, getNotesWorksheetName(env), NOTE_HEADERS);
    await sheetsAppend(env, `${getNotesWorksheetName(env)}!A:F`, [row]);
  } catch (error) {
    throw new Error(`SHEETS_WRITE_FAILED: ${error instanceof Error ? error.message : String(error)}`);
  }
}

async function logInteraction(env, entry) {
  try {
    await ensureWorksheet(env, getWorksheetName(env), LOG_HEADERS);
    const row = [
      nowLocalString(env),
      String(entry.chatId || ""),
      String(entry.user?.id || ""),
      entry.user?.username ? `@${entry.user.username}` : (entry.user?.first_name || ""),
      entry.inputType || "",
      entry.originalText || "",
      entry.normalizedText || "",
      entry.action || "",
      entry.resultSummary || "",
      entry.calendarEventId || "",
      entry.status || "",
    ];
    await sheetsAppend(env, `${getWorksheetName(env)}!A:K`, [row]);
  } catch (error) {
    console.error("logInteraction failed:", error instanceof Error ? error.message : String(error));
  }
}

function getWorksheetName(env) {
  return env.WORKSHEET_NAME || DEFAULT_WORKSHEET_NAME;
}

function getNotesWorksheetName(env) {
  return env.NOTES_WORKSHEET_NAME || DEFAULT_NOTES_WORKSHEET_NAME;
}

async function ensureWorksheet(env, title, headers) {
  const metadata = await sheetsGetMetadata(env);
  const found = metadata.sheets?.find((sheet) => sheet.properties?.title === title);
  if (!found) {
    await sheetsBatchUpdate(env, {
      requests: [{ addSheet: { properties: { title } } }]
    });
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
    body,
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
    {
      method: "PUT",
      body: { range, majorDimension: "ROWS", values }
    }
  );
}

async function sheetsAppend(env, range, values) {
  return await googleApi(
    env,
    `https://sheets.googleapis.com/v4/spreadsheets/${env.GOOGLE_SHEET_ID}/values/${encodeURIComponent(range)}:append?valueInputOption=USER_ENTERED`,
    {
      method: "POST",
      body: { range, majorDimension: "ROWS", values }
    }
  );
}

async function calendarInsertEvent(env, event) {
  return await googleApi(
    env,
    `https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(env.CALENDAR_ID || DEFAULT_CALENDAR_ID)}/events`,
    {
      method: "POST",
      body: event,
    }
  );
}

async function listCalendarEvents(env, timeMin, timeMax, maxResults = 10) {
  const params = new URLSearchParams({
    timeMin,
    timeMax,
    singleEvents: "true",
    orderBy: "startTime",
    maxResults: String(maxResults),
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
      "content-type": "application/json",
    },
    body: options.body ? JSON.stringify(options.body) : undefined,
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
    iat: now,
  };

  const unsignedToken = `${base64UrlEncode(JSON.stringify(jwtHeader))}.${base64UrlEncode(JSON.stringify(jwtClaimSet))}`;
  const signature = await signJwt(unsignedToken, serviceAccount.private_key);
  const assertion = `${unsignedToken}.${signature}`;

  const response = await fetch(serviceAccount.token_uri, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion,
    }),
  });

  const tokenData = await response.json();
  if (!response.ok) {
    throw new Error(`Google OAuth error: ${JSON.stringify(tokenData)}`);
  }

  cachedGoogleToken = {
    token: tokenData.access_token,
    expiresAt: Date.now() + (tokenData.expires_in || 3600) * 1000,
  };
  return cachedGoogleToken.token;
}

async function geminiText(env, { contents, responseMimeType = "text/plain", useSearch = false }) {
  const model = env.GEMINI_MODEL || DEFAULT_MODEL;
  const body = {
    contents,
    generationConfig: {
      temperature: 0.2,
      responseMimeType,
    },
  };
  if (useSearch) {
    body.tools = [{ google_search: {} }];
  }

  const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent`, {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-goog-api-key": env.GEMINI_API_KEY,
    },
    body: JSON.stringify(body),
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

  if (useSearch) {
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
      timeStyle: "short",
    }).format(new Date(iso));
  } catch {
    return iso;
  }
}

function nowLocalString(env) {
  const timezone = env.TIMEZONE || DEFAULT_TIMEZONE;
  return new Intl.DateTimeFormat("uk-UA", {
    timeZone: timezone,
    dateStyle: "short",
    timeStyle: "medium",
  }).format(new Date());
}

function startOfDayInTimezone(baseDate, timeZone, plusDays = 0) {
  const target = new Date(baseDate.getTime() + plusDays * 24 * 60 * 60 * 1000);
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(target);
  const year = parts.find((p) => p.type === "year")?.value;
  const month = parts.find((p) => p.type === "month")?.value;
  const day = parts.find((p) => p.type === "day")?.value;
  const offset = getTimeZoneOffsetString(timeZone, target);
  return `${year}-${month}-${day}T00:00:00${offset}`;
}

function getTimeZoneOffsetString(timeZone, date = new Date()) {
  try {
    const parts = new Intl.DateTimeFormat("en-US", {
      timeZone,
      timeZoneName: "shortOffset",
      hour: "2-digit",
    }).formatToParts(date);
    const tz = parts.find((p) => p.type === "timeZoneName")?.value || "GMT+00:00";
    return tz.replace("GMT", "") || "+00:00";
  } catch {
    return "+00:00";
  }
}

function addHour(iso, hours) {
  const date = new Date(iso);
  date.setHours(date.getHours() + hours);
  return date.toISOString();
}

function addDaysToIsoDate(dateStr, days) {
  const date = new Date(`${dateStr}T00:00:00Z`);
  date.setUTCDate(date.getUTCDate() + days);
  return date.toISOString().slice(0, 10);
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
