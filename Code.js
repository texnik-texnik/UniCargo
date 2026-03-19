const PROPS = PropertiesService.getScriptProperties();
const TOKEN = PROPS.getProperty("TELEGRAM_TOKEN");
const ADMIN_LINK = PROPS.getProperty("ADMIN_LINK") || "";
const DEFAULT_TRACKS_SPREADSHEET_ID = "";
const DEFAULT_USERS_SPREADSHEET_ID = "";

const TRACKS_SPREADSHEET_ID = PROPS.getProperty("TRACKS_SPREADSHEET_ID") || DEFAULT_TRACKS_SPREADSHEET_ID;
const USERS_SPREADSHEET_ID = PROPS.getProperty("USERS_SPREADSHEET_ID") || DEFAULT_USERS_SPREADSHEET_ID;
const TRACKS_FOLDER_ID = PROPS.getProperty("TRACKS_FOLDER_ID") || "";
const TRACKS_FILE_PATTERN = PROPS.getProperty("TRACKS_FILE_PATTERN") || "";
const SHEET_NAME = PROPS.getProperty("TRACK_SHEET_NAME") || "";

const FILE_ID_TRACK_SEARCH = PROPS.getProperty("FILE_ID_TRACK_SEARCH") || "";
const FILE_ID_ADDRESS_CHINA = PROPS.getProperty("FILE_ID_ADDRESS_CHINA") || "";
const FILE_ID_LOC_KHUROSON = PROPS.getProperty("FILE_ID_LOC_KHUROSON") || "";
const FILE_ID_PRICE = PROPS.getProperty("FILE_ID_PRICE") || "";
const FILE_ID_TIME_DELIVERY = PROPS.getProperty("FILE_ID_TIME_DELIVERY") || "";
const FILE_ID_BANNED = PROPS.getProperty("FILE_ID_BANNED") || "";

const PRICE_FLAT_UPTO_0_5_STR = PROPS.getProperty("PRICE_FLAT_UPTO_0_5") || "26 сомони";
const PRICE_PER_KG_ABOVE_0_5_STR = PROPS.getProperty("PRICE_PER_KG_ABOVE_0_5") || "26 сомони/кг";
const PRICE_PER_M3_STR = PROPS.getProperty("PRICE_PER_M3") || "270 $";
/**********************
 *  СОСТОЯНИЯ
 *********************/
const STATE_CHOOSE_WAREHOUSE = "choose_warehouse";
const STATE_CHOOSE_LANG = "choose_lang";
const STATE_REGISTER_NAME = "register_name";
const STATE_REGISTER_PHONE = "register_phone";
const STATE_WAIT_TRACK = "wait_track";
const STATE_CHANGE_LANG = "change_lang";
const STATE_CHANGE_NAME = "change_name";
const STATE_CHANGE_PHONE = "change_phone";
const STATE_CHECK_ADDR = "check_addr_cn";



/**********************
 *  ЛОКАЛИЗАЦИЯ
 *********************/
const TEXT = {
  ru: {
    start_welcome: (name) => `Ассалому алейкум, ${name}! С возвращением!`,
    choose_lang: "Выберите язык / Забонро интихоб кунед:",
    choose_ru: "Русский 🇷🇺",
    choose_tj: "Тоҷикӣ 🇹🇯",

    ask_name: "Пожалуйста, введите ваше имя:",
    ask_phone: "Введите номер телефона, который установлен в вашем смартфоне, если номер на котором зарегистрирован ваш телеграм работает тогда нажмите кнопку ниже:👇🏻",
    btn_share: "📱 Поделиться контактом",
    saved_ok: "✅ Вы успешно зарегистрированы!",
    choose_warehouse: "Выберите склад:",

    btn_zafarobod: "🚚 Зафаробод",
    btn_rasulov: "🚚 Дж. Расулов",
    btn_istaravshan: "🚚 Истаравшан",
    menu_title: "Ассалому Алайкум! Выберите действие:",
    profile_change_wh: "Пункт выдачи товара",
    btn_change_wh: "🏢 Пункт выдачи товара",

    btn_check_track: "🔍 Проверить трек-код",
    btn_register: "📝 Регистрация",
    btn_addr_cn: "📦 Адрес в Китае",
    btn_addr_khu: "🏢 Адрес в Таджикистане",
    btn_price: "💰 Тарифы",
    btn_time: "⏰ Срок доставки",
    btn_banned: "🚫 Запрещённые грузы",
    btn_admin: "☎️ Связь с админом",
    btn_change_lang: "🌐 Сменить язык",
    btn_profile: "👤 Мой профиль",
    
    btn_check_addr_cn: "🧪 Проверить адрес Китая",
    ask_addr_cn_input: "Вставьте ПОЛНЫЙ адрес в Китае для проверки (как в заказе):",
    addr_ok: "✅ Адрес выглядит корректно.",
    addr_fail_header: "❌ Найдены проблемы:",
    addr_correct: "✅ Правильный адрес для копирования:",

    profile_title: "👤 Ваш профиль",
    profile_name: "Имя",
    profile_phone: "Телефон",
    profile_lang: "Язык",
    profile_edit_name: "✏️ Изменить имя",
    profile_edit_phone: "📱 Изменить телефон",
    profile_change_lang: "🌐 Сменить язык",
    ask_new_name: "Введите новое имя:",
    ask_new_phone: "Отправьте новый номер (или нажмите кнопку ниже):",
    saved_name_ok: "✅ Имя обновлено.",
    saved_phone_ok: "✅ Телефон обновлён.",
    canceled: "❌ Отменено.",
    btn_cancel: "❌ Отмена",
    addr_fields_intro: "Скопируйте и вставьте поля в Pinduoduo (каждое сообщение — отдельное поле):",
    
    price_caption:
  `⚖️ До 0.5 кг — ${PRICE_FLAT_UPTO_0_5_STR}\n` +
  `⚖️ Свыше 0.5 кг — ${PRICE_PER_KG_ABOVE_0_5_STR}\n` +
  `📦 1 куб — ${PRICE_PER_M3_STR}`,
    time_caption:
      "Срок доставки из Китая до вашего района или города: 20–25 дней.",
    banned_caption:
      "НАШЕ КАРГО НЕ ПРИНИМАЕТ СЛЕДУЮЩИЕ ГРУЗЫ:\n\n• Стеклянные изделия\n• Оружие и военные предметы\n• Опасные материалы\n• Наркотические и психотропные вещества\n• Животные и растения\n• Порнографические материалы\n• Пищевые и медицинские продукты\n• Фальшивые деньги и документы\n‼️❌ Если вы заказали эти товары то вы подтверждаете что мы не несём ответственность за доставку и принятия вашего товара!",

    track_found: (code, status, weight) =>
      `✅ Трек-код ${code} найден! Ваш заказ принят на складе в Китае.\n⌛ Доставим в Таджикистан в кратчайшие сроки!`,
    track_not_found: "❌ Трек-код не найден! Ваш закас ещё не получен на складе в Китае.",
    admin_caption: "📞 Связь с админом:",
    ask_track_code: "Введите трек-код вашего заказа:",
    lang_switched: "✅ Язык переключён.",
    tracks_cache_flushed: "🔄 Источник треков обновлён.",

    sub_offer: "Хотите подписаться на обновления по этому треку?",
    btn_subscribe: "🔔 Подписаться",
    btn_unsubscribe: "🔕 Отписаться",
    subscribed_ok: "✅ Подписка оформлена.",
    already_subscribed: "ℹ️ Вы уже подписаны.",
    unsubscribed_ok: "✅ Подписка отменена.",
    not_subscribed: "ℹ️ Вы не были подписаны.",
    // Вставьте этот блок внутрь объекта TEXT.ru, например, после banned_caption
    addr_validation: {
  please_register: "Пожалуйста, зарегистрируйтесь, чтобы использовать эту функцию.",
  generic_error: "Произошла техническая ошибка при проверке адреса. Попробуйте позже.",
  empty_address: "Адрес не может быть пустым.",
  no_chinese_chars: "В адресе отсутствуют китайские иероглифы.",
  client_code_not_found: (code) => `Не найден ваш клиентский код ${code}.`,
  client_code_twice: (code) => `Ваш код ${code} должен быть указан дважды.`,
  other_code_found: (codes) => `Обнаружен посторонний код клиента: ${codes}.`,
  cn_phone_not_found: "Не найден китайский номер телефона (11 цифр, начинается на 1).",
  cn_phone_mismatch: (expected, found) => `Китайский номер не совпадает. Ожидался: ${expected}, найден: ${found}.`,
  region_not_found: "Не найдены обязательные части адреса (провинция 省 / город 市).",
  not_yiwu: "Адрес должен относиться к городу Иу (义乌市).",
  details_not_found: "Отсутствуют детали адреса (например: здание 栋, корпус 号, склад 仓).",
  warehouse_char_missing: "В адресе отсутствует обязательный иероглиф «склад» (仓).",
  chinese_part_mismatch: "Основная китайская часть адреса не совпадает с эталоном.",
  name_not_found: (name) => `В адресе не найдено ваше имя: "${name}".`,
  phone_not_found: (phone) => `В адресе не найден ваш номер телефона: ${phone}.`,
  keyword_missing: (keyword) => `В адресе отсутствует обязательное слово/код: "${keyword}".`
    },
    status_update_initial: (code, date) => `🔔 Ваш трек-код ${code} прибыл на склад в Китае!\nДата приёмки: ${formatDate(date)}`,
    admin_only: "Команда доступна только администратору.",
    admin_panel_title: "🛠 Панель админа",
    admin_stats: (u, s) => `Пользователей: ${u}\nАктивных подписок: ${s}`
  },
  tj: {
    start_welcome: (name) => `Ассалому алейкум, ${name}! Хуш омадед!`,
    choose_lang: "Забонро интихоб кунед / Выберите язык:",
    choose_ru: "Русский 🇷🇺",
    choose_tj: "Тоҷикӣ 🇹🇯",

    ask_name: "Лутфан номи худро ворид кунед:",
    ask_phone: "Рақами телефони дастиатонро дохил кунед, агар рақаме ки бо он телеграмро регистратсия када бошед, фаъол бошад тугмаро зер кунед👇🏻:",
    btn_share: "📱 Ирсоли рақами телефон",
    saved_ok: "✅ Шумо бомуваффақият сабт шудед!",
    choose_warehouse: "Складро интихоб кунед:",
    btn_zafarobod: "🚚 Зафаробод",
    btn_rasulov: "🚚 Ҷ. Расулов",
    btn_istaravshan: "🚚 Истаравшан",
    menu_title: "Ассалому Алайкум! Амалиётро интихоб кунед:",
    profile_change_wh: "Нуқтаи қабули бор",
    btn_change_wh: "🏢 Нуқтаи қабули бор",

    btn_check_track: "🔍 Санҷишӣ трек-код",
    btn_register: "📝 Сабти ном",
    btn_addr_cn: "📦 Адреси Хитой",
    btn_addr_khu: "🏢 Адреси Тоҷикистон",
    btn_price: "💰 Нархи хизматрасонӣ",
    btn_time: "⏰ Вақти расондан",
    btn_banned: "🚫 Борҳои манъшуда",
    btn_admin: "☎️ Алоқа бо админ",
    btn_change_lang: "🌐 Иваз кардани забон",
    btn_profile: "👤 Профили ман",
    btn_check_addr_cn: "🧪 Санҷиши адреси Хитой",
    ask_addr_cn_input: "Адреси пурраи Хитойро барои санҷиш фиристед:",
    addr_ok: "✅ Суроға дуруст аст.",
    addr_fail_header: "❌ Мушкилот ёфт шуд:",
    addr_correct: "✅ Суроғаи дуруст барои нусхабардорӣ:",

    profile_title: "👤 Профили шумо",
    profile_name: "Ном",
    profile_phone: "Телефон",
    profile_lang: "Забон",
    profile_edit_name: "✏️ Тағйири ном",
    profile_edit_phone: "📱 Тағйири телефон",
    profile_change_lang: "🌐 Иваз кардани забон",
    ask_new_name: "Номи навро ворид кунед:",
    ask_new_phone: "Рақами навро фиристед (ё тугмаро пахш кунед):",
    saved_name_ok: "✅ Ном нав шуд.",
    saved_phone_ok: "✅ Телефон нав шуд.",
    canceled: "❌ Бекор шуд.",
    btn_cancel: "❌ Бекор",
    price_caption:
  `⚖️ То 0.5 кг — ${PRICE_FLAT_UPTO_0_5_STR}\n` +
  `⚖️ Аз 0.5 кг боло — ${PRICE_PER_KG_ABOVE_0_5_STR}\n` +
  `📦 Нархи як куб: ${PRICE_PER_M3_STR}`,
    time_caption:
      "Вақти расондан аз Хитой то ноҳия ё шаҳри шумо: 18–25 рӯз!",
    banned_caption:
      "КАРГОИ МО ИН НАМУДИ БОРҲОРО ҚАБУЛ НАМЕКУНАД:\n\n• Ашёҳои шишагӣ\n• Яроқ ва ашёҳои ҳарбӣ\n• Маводҳои хатарнок\n• Маводҳои наркотикӣ ва психотропӣ\n• Ҳайвонот ва растаниҳо\n• Маводҳои порнографӣ\n• Маҳсулотҳои хӯрока и тиббӣ\n• Пул ва ҳуҷҷатҳои қалбакӣ\n‼️❌ Агар шумо ин маҳсулотро фармоиш дода бошед, шумо тасдиқ мекунед, ки мо барои интиқол ва қабули молҳои шумо масъул нестем!",
      

    track_found: (code, status, weight) =>
      `✅ Трек-коди ${code} ёфт шуд! Закази шумо дар склади Хитой қабул шудааст.\n⌛ Дар муддати кӯтоҳтарин ба Тоҷикистон мерасонем!`,
    track_not_found: "❌ Трек-код ёфт нашуд! Бори шумо дар склади Хитой холо кабул нашудааст.",
    admin_caption: "📞 Алоқа бо админ:",
    ask_track_code: "Трек-коди заказро ворид кунед:",
    lang_switched: "✅ Забон иваз шуд.",
    tracks_cache_flushed: "🔄 Манбаи трекҳо нав шуд.",

    sub_offer: "Мехоҳед ба навсозиҳои ин трек обуна шавед?",
    btn_subscribe: "🔔 Обуна шудан",
    btn_unsubscribe: "🔕 Лағв кардани обуна",
    subscribed_ok: "✅ Обуна фаъол шуд.",
    already_subscribed: "ℹ️ Шумо аллакай обунаед.",
    unsubscribed_ok: "✅ Обуна бекор шуд.",
    not_subscribed: "ℹ️ Шумо обуна набудед.",
    status_update_initial: (code, date) => `🔔 Трек-коди шумо ${code} ба анбори Чин расид!\nСанаи қабул: ${formatDate(date)}`,
    admin_only: "Фармон танҳо барои админ дастрас аст.",
// Вставьте этот блок внутрь объекта TEXT.tj
    addr_validation: {
      please_register: "Лутфан, барои истифодаи ин функсия аввал аз қайд гузаред.",
      generic_error: "Ҳангоми санҷиши суроға хатогии техникӣ рух дод. Лутфан баъдтар кӯшиш кунед.",
      empty_address: "Суроға наметавонад холӣ бошад.",
      no_chinese_chars: "Дар суроға иероглифҳои хитоӣ вуҷуд надоранд.",
      client_code_not_found: (code) => `Коди муштарии шумо (${code}) ёфт нашуд.`,
      client_code_twice: (code) => `Коди шумо (${code}) бояд ду маротиба нишон дода шавад.`,
      other_code_found: (codes) => `Коди дигари муштарӣ ёфт шуд: ${codes}.`,
      cn_phone_not_found: "Рақами телефони хитоӣ ёфт нашуд (11 рақам, бо 1 сар мешавад).",
      cn_phone_mismatch: (expected, found) => `Рақами телефони хитоӣ мувофиқ нест. Интизорӣ: ${expected}, ёфт шуд: ${found}.`,
      region_not_found: "Қисмҳои ҳатмии суроға ёфт нашуданд (вилоят 省 / шаҳр 市).",
      not_yiwu: "Суроға бояд ба шаҳри Иу (义乌市) тааллуқ дошта бошад.",
      details_not_found: "Тафсилоти суроға мавҷуд нест (масалан: бино 栋, корпус 号, анбор 仓).",
      warehouse_char_missing: "Дар суроға иероглифи ҳатмии «анбор» (仓) вуҷуд надорад.",
      chinese_part_mismatch: "Қисми асосии хитоии суроға ба намуна мувофиқат намекунад.",
      name_not_found: (name) => `Дар суроға номи шумо ёфт нашуд: "${name}".`,
      phone_not_found: (phone) => `Дар суроға рақами телефони шумо ёфт нашуд: ${phone}.`,
      keyword_missing: (keyword) => `Дар суроға калима/коди ҳатмӣ вуҷуд надорад: "${keyword}".`
    },    
    admin_panel_title: "🛠 Панели админ",
    admin_stats: (u, s) => `Корбарон: ${u}\nОбунаҳои фаъол: ${s}`
  }
};
// Google Dvanced Drive API
function getOrCreateSubfolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

// Требуется Advanced Drive Service (Drive) включённый в проекте
function convertFileToGoogleSheet_(file) {
  const resource = {
    title: file.getName().replace(/\.(xlsx|xls|csv)$/i, ""),
    mimeType: 'application/vnd.google-apps.spreadsheet',
    parents: [{ id: TRACKS_FOLDER_ID }]
  };
  const blob = file.getBlob();
  const newFile = Drive.Files.insert(resource, blob); // создаст Google Sheet
  return newFile.id;
}

function convertNewUploads_() {
  if (!TRACKS_FOLDER_ID) { Logger.log("TRACKS_FOLDER_ID is empty"); return; }
  const folder = DriveApp.getFolderById(TRACKS_FOLDER_ID);
  const archive = getOrCreateSubfolder_(folder, "_archive");

  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    const mime = f.getMimeType();

    // Пропускаем уже Google Таблицы
    if (mime === MimeType.GOOGLE_SHEETS) continue;

    // Обрабатываем Excel/CSV
    const isExcel = mime === MimeType.MICROSOFT_EXCEL ||
      mime === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
      mime === "application/vnd.ms-excel";
    const isCsv = mime === MimeType.CSV || mime === "text/csv";

    if (isExcel || isCsv) {
      try {
        const newId = convertFileToGoogleSheet_(f);
        // Перемещаем оригинал в архив, чтобы не конвертировать повторно
        archive.addFile(f);
        folder.removeFile(f);
        Logger.log("Converted: " + f.getName() + " -> " + newId);
      } catch (e) {
        Logger.log("Convert error for " + f.getName() + ": " + e);
      }
    }
  }
}

/*******************************************************
 *  ОБРАБОТЧИКИ СОСТОЯНИЙ И КОМАНД (HANDLERS)
 *******************************************************/

// --- Обработчики состояний (когда бот чего-то ждет от пользователя) ---

function handleStateWaitingForTrack(chatId, userId, text, lang) {
  const L = TEXT[lang];
  try {
    const codes = text.split(/[\s,;\n]+/).map(s => s.trim()).filter(Boolean).slice(0, 50);
    if (codes.length === 0) {
      sendMessage(chatId, L.ask_track_code);
      return;
    }

    const res = findTracksInfoBulk_(codes);
    const lines = [];

    for (const rawCode of codes) {
      const displayCode = String(rawCode).trim().toUpperCase();
      const normalizedCode = normalizeCode_(rawCode);
      const info = res[normalizedCode] || { found: false, dateText: null };

      if (info.found) {
        const dateString = info.dateText ? String(info.dateText).trim() : "-";
        const line = lang === "ru" 
          ? `✅ ${displayCode} найден! Получен: ${dateString}.` 
          : `✅ ${displayCode} ёфт шуд! Қабул шуд: ${dateString}.`;
        
        // --- ОТЛАДОЧНЫЙ ЛОГ ---
        debugLog("handleStateWaitingForTrack", `[${displayCode}] Готовая строка для юзера`, line);
        
        lines.push(line);
      } else {
        const line = lang === "ru" 
          ? `❌ ${displayCode} не найден.` 
          : `❌ ${displayCode} ёфт нашуд.`;
        lines.push(line);
      }
    }
    sendMessage(chatId, lines.join("\n"));
  } catch (e) {
    Logger.log("track search error: " + e);
    sendMessage(chatId, lang === "ru" ? "Техническая ошибка при поиске." : "Хатои техникӣ.");
  } finally {
    clearUserState(userId);
  }
}

function handleStateChangingLang(chatId, userId, text, lang) {
  if (text === TEXT.ru.choose_ru || text === TEXT.tj.choose_ru || text === TEXT.ru.choose_tj || text === TEXT.tj.choose_tj) {
    const newLang = (text === TEXT.ru.choose_ru || text === TEXT.tj.choose_ru) ? "ru" : "tj";
    updateUserLang(userId, newLang);
    setLangCache(userId, newLang);
    clearUserState(userId);
    sendMessage(chatId, TEXT[newLang].lang_switched);
    sendMenu(chatId, newLang);
  } else {
    sendLanguageChoice(chatId, lang);
  }
}

function handleStateChangingName(chatId, userId, text, lang) {
  const L = TEXT[lang];
  updateUserName_(userId, text);
  clearUserState(userId);
  sendMessageWithKeyboardRemove(chatId, L.saved_name_ok);
  sendMenu(chatId, lang);
  logEvent_("change_name", userId, `name=${text}`);
}

function handleStateChangingPhone(chatId, userId, text, contact, lang) {
  const L = TEXT[lang];
  let phone = contact?.phone_number ? normalizeTjPhone(contact.phone_number) : (text ? normalizeTjPhone(text) : null);
  if (phone && isValidTjPhone(phone)) {
    updateUserPhone_(userId, phone);
    clearUserState(userId);
    sendMessageWithKeyboardRemove(chatId, L.saved_phone_ok);
    sendMenu(chatId, lang);
    logEvent_("change_phone", userId, `phone=${phone}`);
  } else {
    sendContactButton(chatId, lang);
  }
}

function handleStateCheckAddress(chatId, userId, text, lang) {
  checkChinaAddressAndReply_(chatId, userId, text, lang);
  clearUserState(userId);
  sendMenu(chatId, lang);
}


// --- Обработчики команд (когда пользователь нажимает кнопку в меню) ---

function handleCommandCheckTrack(chatId, userId, lang) {
  const L = TEXT[lang];
  setUserState(userId, STATE_WAIT_TRACK);
  sendPhotoFromDrive(chatId, FILE_ID_TRACK_SEARCH, L.ask_track_code);
}

function handleCommandChangeWarehouse(chatId, userId, lang) {
  setUserState(userId, STATE_CHOOSE_WAREHOUSE);
  sendWarehouseChoice(chatId, lang);
}

function handleCommandAdminContact(chatId, user, lang) {
  const { addr, admin } = getTjData(user.warehouse);
  sendCopyableText(chatId, addr + "\n📞 Админ: " + admin);
}

// ЗАМЕНИТЕ ВАШУ ФУНКЦИЮ НА ЭТУ
function handleCommandChinaAddress(chatId, user, lang) {
  const warehouse = user.warehouse; // Получаем склад пользователя
  const { addr, fileId } = getChinaData(warehouse, user.name, user.phone);
  
  // --- НОВЫЙ КОД: Получаем ID видео для конкретного склада ---
  const videoId = getWhProp("PINDUODUO_VIDEO", warehouse);
  if (videoId) {
    sendVideoFromDrive(chatId, videoId);
  }
  
  // --- Старый код остается без изменений ---
  sendPhotoFromDrive(chatId, fileId, addr);
  sendCopyableText(chatId, addr);
}

function handleCommandTjAddress(chatId, user, lang) {
  const { addr, admin } = getTjData(user.warehouse);
  sendCopyableText(chatId, addr + "\n📞 Админ: " + admin);
}

function handleCommandPrice(chatId, user, lang) {
  const L = TEXT[lang];
  sendPhotoFromDrive(chatId, FILE_ID_PRICE, L.price_caption);
}

function handleCommandTime(chatId, user, lang) {
  const L = TEXT[lang];
  sendPhotoFromDrive(chatId, FILE_ID_TIME_DELIVERY, L.time_caption);
}

function handleCommandBanned(chatId, user, lang) {
  const L = TEXT[lang];
  sendPhotoFromDrive(chatId, FILE_ID_BANNED, L.banned_caption);
}

function handleCommandProfile(chatId, user, lang) {
  showProfile_(chatId, user, lang);
}

function handleCommandCheckAddressStart(chatId, userId, lang) {
  const L = TEXT[lang];
  setUserState(userId, STATE_CHECK_ADDR);
  sendCancelKeyboard_(chatId, L.ask_addr_cn_input, lang);
}


// --- Обработчики процесса регистрации ---
function handleRegistration(chatId, userId, text, contact, lang, state, user) {
  const L = TEXT[lang];
  // 1. Выбор языка
  if (text === TEXT.ru.choose_ru || text === TEXT.tj.choose_ru || text === TEXT.ru.choose_tj || text === TEXT.tj.choose_tj) {
    const newLang = (text.includes(TEXT.ru.choose_ru)) ? "ru" : "tj";
    setLangCache(userId, newLang);
    setUserState(userId, STATE_REGISTER_NAME);
    sendMessage(chatId, TEXT[newLang].ask_name);
    return true; // Запрос обработан
  }
  // 2. Ввод имени
  if (state === STATE_REGISTER_NAME && text) {
    setUserState(userId, STATE_REGISTER_PHONE, { name: text });
    sendContactButton(chatId, lang);
    return true;
  }
  // 3. Ввод телефона
  if (state === STATE_REGISTER_PHONE) {
    const temp = getUserTemp(userId) || {};
    const name = temp.name || "ном";
    let phone = contact?.phone_number ? normalizeTjPhone(contact.phone_number) : (text ? normalizeTjPhone(text) : null);
    if (phone && isValidTjPhone(phone)) {
      saveOrUpdateUser_({ userId, name, phone, lang });
      clearUserState(userId);
      sendMessageWithKeyboardRemove(chatId, L.saved_ok);
      setUserState(userId, STATE_CHOOSE_WAREHOUSE);
      sendWarehouseChoice(chatId, lang);
      logEvent_("register", userId, `name=${name}, phone=${phone}, lang=${lang}`);
    } else {
      sendContactButton(chatId, lang); // Попросить еще раз
    }
    return true;
  }
  
  // 4. Если ничего не подошло, начинаем регистрацию с выбора языка
  setUserState(userId, STATE_CHOOSE_LANG);
  sendLanguageChoice(chatId, lang);
  return true;
}
/**********************
 *  ВХОДНАЯ ТОЧКА (РЕФАКТОРИНГ)
 *********************/
function doPost(e) {
  try {
    if (!TOKEN) throw new Error("TELEGRAM_TOKEN не задан в Properties.");

    const upd = JSON.parse(e.postData.contents);

    // 1. Обработка callback-кнопок (остается без изменений)
    if (upd.callback_query) {
      handleCallbackQuery(upd.callback_query);
      return;
    }

    const msg = upd.message;
    if (!msg) return;

    // 2. Инициализация переменных
    const chatId = msg.chat.id;
    const userId = String(msg.from.id);
    const text = (msg.text || "").trim();
    const contact = msg.contact || null;
    let user = getUserFromCsv(userId);
    let lang = user?.lang || getLangCache(userId) || detectLang_(msg.from?.language_code) || "ru";
    const L = TEXT[lang];

    // 3. Обработка команд администратора (текстовые команды с /)
    const adminCommands = {
      "/reloadtracks": () => { flushTracksCache_(); sendMessage(chatId, L.tracks_cache_flushed); },
      "/admin": () => {
        const ucount = countUsers_();
        const scount = countActiveSubs_();
        sendMessage(chatId, `${L.admin_panel_title}\n${L.admin_stats(ucount, scount)}`);
      },
      "/setup_checker": () => { setupSubscriptionChecker_(); sendMessage(chatId, "✅ Триггер подписок создан."); },
      "/broadcast": () => {
        const args = text.split(/\s+/).slice(1); // убираем команду, оставляем аргументы
        handleBroadcastCommand(chatId, userId, args);
      },
      "/stopbroadcast": () => {
        if (!isAdmin_(userId)) { sendMessage(chatId, L.admin_only); return; }
        setBroadcastFlag_(false);
        sendMessage(chatId, "🛑 Рассылка остановлена.");
        Logger.log("Рассылка остановлена администратором");
      }
    };

    if (adminCommands[text] || text.startsWith("/broadcast ")) {
      if (!isAdmin_(userId)) { sendMessage(chatId, L.admin_only); return; }
      if (adminCommands[text]) {
        adminCommands[text]();
      } else if (text.startsWith("/broadcast ")) {
        const args = text.split(/\s+/).slice(1);
        handleBroadcastCommand(chatId, userId, args);
      }
      return;
    }

    // 4. Обработка команды /start
    if (text === "/start") {
      if (user) {
        sendMessage(chatId, L.start_welcome(user.name));
        sendMenu(chatId, lang);
      } else {
        setUserState(userId, STATE_CHOOSE_LANG);
        sendLanguageChoice(chatId, lang);
      }
      return;
    }

    // 5. Маршрутизатор для НЕЗАРЕГИСТРИРОВАННЫХ пользователей
    if (!user) {
      const state = getUserState(userId);
      // Сначала проверяем выбор склада, т.к. он может быть последним шагом регистрации
      if (handleChooseWarehouse(state, text, userId, chatId, lang)) return;
      // Затем передаем управление общему обработчику регистрации
      handleRegistration(chatId, userId, text, contact, lang, state, user);
      return;
    }
    
    // ==========================================================
    // ===== Логика для ЗАРЕГИСТРИРОВАННЫХ пользователей =======
    // ==========================================================
    
    const state = getUserState(userId);
    
    // 6. Обработка универсальной отмены
    if (text === L.btn_cancel) {
      clearUserState(userId);
      sendMessageWithKeyboardRemove(chatId, L.canceled);
      sendMenu(chatId, lang);
      return;
    }

    // 7. Приоритетный маршрутизатор по СОСТОЯНИЯМ
    const stateHandlers = {
      [STATE_WAIT_TRACK]: (text) => handleStateWaitingForTrack(chatId, userId, text, lang),
      [STATE_CHANGE_LANG]: (text) => handleStateChangingLang(chatId, userId, text, lang),
      [STATE_CHANGE_NAME]: (text) => handleStateChangingName(chatId, userId, text, lang),
      [STATE_CHANGE_PHONE]: (text, contact) => handleStateChangingPhone(chatId, userId, text, contact, lang),
      [STATE_CHECK_ADDR]: (text) => handleStateCheckAddress(chatId, userId, text, lang),
      [STATE_CHOOSE_WAREHOUSE]: (text) => {
        if (handleChooseWarehouse(state, text, userId, chatId, lang)) { /* already handled */ }
      }
      // TODO: Добавить обработчики состояний для заказа (STATE_ORDER_LINK и т.д.)
    };

    if (state && stateHandlers[state]) {
      stateHandlers[state](text, contact);
      return;
    }

    // 8. Маршрутизатор по ТЕКСТУ КОМАНД (кнопкам меню)
    const commandHandlers = {
      [L.btn_check_track]: () => handleCommandCheckTrack(chatId, userId, lang),
      [L.btn_change_wh]: () => handleCommandChangeWarehouse(chatId, userId, lang),
      [L.btn_admin]: () => handleCommandAdminContact(chatId, user, lang),
      [L.btn_addr_cn]: () => handleCommandChinaAddress(chatId, user, lang),
      [L.btn_addr_khu]: () => handleCommandTjAddress(chatId, user, lang),
      [L.btn_price]: () => handleCommandPrice(chatId, user, lang),
      [L.btn_time]: () => handleCommandTime(chatId, user, lang),
      [L.btn_banned]: () => handleCommandBanned(chatId, user, lang),
      [L.btn_profile]: () => handleCommandProfile(chatId, user, lang),
      [L.btn_check_addr_cn]: () => handleCommandCheckAddressStart(chatId, userId, lang),
      // "/checkaddr" - альтернативная текстовая команда
      ["/checkaddr"]: () => handleCommandCheckAddressStart(chatId, userId, lang),
      [L.btn_register]: () => sendMenu(chatId, lang) // Кнопка "Регистрация" для зарег. просто показывает меню
    };

    if (commandHandlers[text]) {
      commandHandlers[text]();
      return;
    }
    
    // 9. Если ничего не подошло - показываем меню
    sendMenu(chatId, lang);

  } catch (err) {
    Logger.log(err);
    logEvent_("error", "system", String(err));
  }
}

/*******************************************************
 *  ОБРАБОТЧИК НАЖАТИЙ НА ИНЛАЙН-КНОПКИ
 *******************************************************/
function handleCallbackQuery(cq) {
  const userId = String(cq.from.id);
  const chatId = cq.message?.chat?.id;
  const data = cq.data || "";
  const user = getUserFromCsv(userId);
  const lang = user?.lang || getLangCache(userId) || "ru";
  const L = TEXT[lang] || TEXT.ru;

  try {
    if (data === "edit_name") {
      setUserState(userId, STATE_CHANGE_NAME);
      sendCancelKeyboard_(chatId, L.ask_new_name, lang);
      answerCallbackQuery(cq.id, " ");
      return;
    }
    if (data === "change_wh") {
      setUserState(userId, STATE_CHOOSE_WAREHOUSE);
      sendWarehouseChoice(chatId, lang);
      answerCallbackQuery(cq.id, " ");
      return;
    }

    if (data === "edit_phone") {
      setUserState(userId, STATE_CHANGE_PHONE);
      sendContactButton(chatId, lang);
      answerCallbackQuery(cq.id, " ");
      return;
    }
    if (data === "change_lang") {
      setUserState(userId, STATE_CHANGE_LANG);
      sendLanguageChoice(chatId, lang);
      answerCallbackQuery(cq.id, " ");
      return;
    }

    // Подписки
    if (data.startsWith("sub:")) {
      const code = data.split(":")[1];
      const exists = isSubscribed_(userId, code);
      if (exists) {
        sendMessage(chatId, L.already_subscribed);
      } else {
        subscribeToTrack_(userId, code);
        sendMessage(chatId, L.subscribed_ok);
        logEvent_("subscribe", userId, `code=${code}`);
      }
      answerCallbackQuery(cq.id, " ");
      return;
    }
    if (data.startsWith("unsub:")) {
      const code = data.split(":")[1];
      const exists = isSubscribed_(userId, code);
      if (!exists) {
        sendMessage(chatId, L.not_subscribed);
      } else {
        unsubscribeFromTrack_(userId, code);
        sendMessage(chatId, L.unsubscribed_ok);
        logEvent_("unsubscribe", userId, `code=${code}`);
      }
      answerCallbackQuery(cq.id, " ");
      return;
    }

    answerCallbackQuery(cq.id, " ");
  } catch (e) {
    Logger.log("handleCallbackQuery error: " + e);
    answerCallbackQuery(cq.id, "Error");
  }
}

/**********************
 *  ВСПОМОГАТЕЛЬНЫЕ
 *********************/

// --- ДОБАВИТЬ функцию handleChooseWarehouse ---
function handleChooseWarehouse(state, text, userId, chatId, lang) {
  if (state !== STATE_CHOOSE_WAREHOUSE) return false;
  let warehouse = null;
  if (text === TEXT.ru.btn_zafarobod || text === TEXT.tj.btn_zafarobod) warehouse = "ZAFAROBOD";
  if (text === TEXT.ru.btn_rasulov || text === TEXT.tj.btn_rasulov) warehouse = "RASULOV";
  if (text === TEXT.ru.btn_istaravshan || text === TEXT.tj.btn_istaravshan) warehouse = "ISTARAVSHAN";

  if (warehouse) {
    updateUserWarehouse(userId, warehouse);
    clearUserState(userId);
    sendMessageWithKeyboardRemove(chatId, TEXT[lang].saved_ok);
    sendMenu(chatId, lang);
    return true;
  }
  return false;
}

// Заменить существующую getWhProp на эту версию
function getWhProp(keyBase, warehouse) {
  if (!warehouse) return "";
  // защита: явно приводим warehouse к строке
  return PROPS.getProperty(`${keyBase}_${String(warehouse).toUpperCase()}`) || "";
}

// Заменить существующую getChinaData на этот код
function getChinaData(warehouse, name, phone) {
  // общий шаблон адреса (в Properties: CN_ADDR_TEMPLATE)
  const template = PROPS.getProperty("CN_ADDR_TEMPLATE") || "";

  // эти параметры ЗАВИСЯТ от склада
  const code = getWhProp("CN_CODE", warehouse) || "";
  const phoneCn = PROPS.getProperty("CN_PHONE_COMMON") || "";
  const fileId = getWhProp("CN_PHOTO", warehouse) || "";
  const phoneDigits = onlyDigits(phone || "");
  // подставляем код и персональные поля
  const addr = template
    .replace(/{code}/g, code)
    .replace(/{name}/g, name || "")
    .replace(/{phone}/g, phoneDigits || "")
    .replace(/{cn_phone}/g, phoneCn || "");

  return { addr, code, phoneCn, fileId };
}

function getTjData(warehouse) {
  return {
    addr: getWhProp("TJ_ADDR", warehouse),
    admin: getWhProp("TJ_ADMIN", warehouse)
  };
}

// --- Новая функция для показа кнопок склада (ДОБАВИТЬ) ---
function sendWarehouseChoice(chatId, lang) {
  const L = TEXT[lang] || TEXT.ru;
  const kb = [
    [{ text: L.btn_zafarobod }],
    [{ text: L.btn_istaravshan }]
  ];
  const url = `https://api.telegram.org/bot${TOKEN}/sendMessage`;
  const payload = {
    chat_id: chatId,
    text: L.choose_warehouse,
    reply_markup: JSON.stringify({
      keyboard: kb,
      resize_keyboard: true,
      one_time_keyboard: true
    })
  };
  UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  });
}

function formatDate(d) {
  if (!d) return "-";
  try {
    const dt = (d instanceof Date) ? d : new Date(d);
    return Utilities.formatDate(dt, Session.getScriptTimeZone(), "dd.MM.yyyy");
  } catch (_) { return String(d); }
}
function detectLang_(code) {
  if (!code) return null;
  if (String(code).toLowerCase().startsWith("ru")) return "ru";
  return "tj";
}

// Телефон RT (+992…)
function normalizeTjPhone(raw) {
  let phone = String(raw).replace(/[^\d+]/g, "");
  if (phone.startsWith("+")) return phone;
  if (phone.startsWith("992")) return "+" + phone;
  if (phone.startsWith("8")) return "+992" + phone.substring(1);
  if (phone.startsWith("0")) return "+992" + phone.substring(1);
  return "+992" + phone;
}
function isValidTjPhone(phone) {
  return /^\+992\d{9}$/.test(phone);
}

/* --------- Cache для состояний и языка --------- */
function setUserState(userId, state, tempData) {
  const c = CacheService.getUserCache();
  c.put(userId + ":state", state, 1800);
  if (tempData) c.put(userId + ":temp", JSON.stringify(tempData), 1800);
}
function getUserState(userId) {
  return CacheService.getUserCache().get(userId + ":state");
}
function getUserTemp(userId) {
  const t = CacheService.getUserCache().get(userId + ":temp");
  return t ? JSON.parse(t) : {};
}
function clearUserState(userId) {
  const c = CacheService.getUserCache();
  c.remove(userId + ":state");
  c.remove(userId + ":temp");
}
function setLangCache(userId, lang) {
  CacheService.getUserCache().put(userId + ":lang", lang, 86400 * 30);
}
function getLangCache(userId) {
  return CacheService.getUserCache().get(userId + ":lang");
}

/* --------- Users --------- */
function getUsersSheet_() {
  const ss = SpreadsheetApp.openById(USERS_SPREADSHEET_ID);
  let sheet = ss.getSheetByName("Users");
  if (!sheet) sheet = ss.insertSheet("Users");
  const header = ["userId", "name", "phone", "date", "lang", "warehouse"];
  const firstRow = sheet.getRange(1, 1, 1, header.length).getValues()[0];
  const isHeaderEmpty = firstRow.every(v => v === "" || v === null || v === undefined);
  if (isHeaderEmpty) sheet.getRange(1, 1, 1, header.length).setValues([header]);
  return sheet;
}
function getUserFromCsv(userId) {
  const sheet = getUsersSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const rng = sheet.getRange(2, 1, last - 1, 1);
  const f = rng.createTextFinder(String(userId)).matchEntireCell(true).findNext();
  if (!f) return null;
  const r = f.getRow();
  const row = sheet.getRange(r, 1, 1, 6).getValues()[0];
  const [id, name, phone, date, lang, warehouse] = row;
  return { userId: String(id), name: name || "", phone: phone || "", date: date || "", lang: lang || "tj", warehouse: warehouse || "" };
}
// --- ДОБАВИТЬ функцию updateUserWarehouse ---
function updateUserWarehouse(userId, warehouse) {
  const sheet = getUsersSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return;
  const rng = sheet.getRange(2, 1, last - 1, 1);
  const f = rng.createTextFinder(String(userId)).matchEntireCell(true).findNext();
  if (!f) return;
  const r = f.getRow();
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try { sheet.getRange(r, 6).setValue(String(warehouse)); }
  finally { lock.releaseLock(); }
}
function saveOrUpdateUser_({ userId, name, phone, lang }) {
  const sheet = getUsersSheet_();
  const nowIso = new Date().toISOString();
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const last = sheet.getLastRow();
    if (last >= 2) {
      const rng = sheet.getRange(2, 1, last - 1, 1);
      const f = rng.createTextFinder(String(userId)).matchEntireCell(true).findNext();
      if (f) {
        const r = f.getRow();
        sheet.getRange(r, 2, 1, 4).setValues([[String(name), String(phone), nowIso, String(lang || "tj")]]);
        return;
      }
    }
    sheet.appendRow([String(userId), String(name), String(phone), nowIso, String(lang || "tj"), ""]);
  } finally {
    lock.releaseLock();
  }
}
function updateUserLang(userId, newLang) {
  const sheet = getUsersSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return;
  const rng = sheet.getRange(2, 1, last - 1, 1);
  const f = rng.createTextFinder(String(userId)).matchEntireCell(true).findNext();
  if (!f) return;
  const r = f.getRow();
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try { sheet.getRange(r, 5).setValue(String(newLang)); } finally { lock.releaseLock(); }
}
function updateUserName_(userId, newName) {
  const sheet = getUsersSheet_();
  const last = sheet.getLastRow(); if (last < 2) return;
  const rng = sheet.getRange(2, 1, last - 1, 1);
  const f = rng.createTextFinder(String(userId)).matchEntireCell(true).findNext();
  if (!f) return;
  const r = f.getRow();
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try { sheet.getRange(r, 2).setValue(String(newName)); } finally { lock.releaseLock(); }
}
function updateUserPhone_(userId, newPhone) {
  const sheet = getUsersSheet_();
  const last = sheet.getLastRow(); if (last < 2) return;
  const rng = sheet.getRange(2, 1, last - 1, 1);
  const f = rng.createTextFinder(String(userId)).matchEntireCell(true).findNext();
  if (!f) return;
  const r = f.getRow();
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try { sheet.getRange(r, 3).setValue(String(newPhone)); } finally { lock.releaseLock(); }
}

/* --------- Таблица треков --------- */
function getTracksSpreadsheetId_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("tracks_ss_id");
  if (cached) return cached;
  let ssId = null;
  if (TRACKS_FOLDER_ID) ssId = findLatestSpreadsheetIdWithSheetInFolder_(TRACKS_FOLDER_ID, TRACKS_FILE_PATTERN, SHEET_NAME);
  if (!ssId) ssId = TRACKS_SPREADSHEET_ID;
  cache.put("tracks_ss_id", ssId, 300);
  return ssId;
}
function flushTracksCache_() {
  CacheService.getScriptCache().remove("tracks_ss_id");
}
function normalizeCode_(s) {
  return String(s || "")
    .replace(/[\u200B-\u200D\uFEFF]/g, "") // zero‑width
    .replace(/\u00A0/g, " ")               // NBSP -> space
    .trim()
    .replace(/\s+/g, "")                   // убрать любые пробелы
    .toUpperCase();
}

function getB1Date_(sh) {
  try {
    const v = sh.getRange(1, 2).getValue(); // B1
    return parseDateFromText_(v);           // Date или null
  } catch (e) {
    Logger.log("getB1Date_ error: " + e);
    return null;
  }
}

function getFileLevelDate_(sh) {
  try {
    // 1) пробуем B1
    const b1 = sh.getRange(1, 2).getValue();
    const dtB1 = parseDateFromText_(b1);
    if (dtB1) return dtB1;

    // 2) пробуем имя файла
    const fileName = sh.getParent().getName();
    const dtName = parseDateFromText_(fileName);
    if (dtName) return dtName;
  } catch (e) {
    Logger.log("getFileLevelDate_ error: " + e);
  }
  return null;
}
function findLatestSpreadsheetIdWithSheetInFolder_(folderId, pattern, sheetName) {
  let query = `mimeType = 'application/vnd.google-apps.spreadsheet' and '${folderId}' in parents`;
  if (pattern) {
    const safe = pattern.replace(/'/g, "\\'");
    query += ` and title contains '${safe}'`;
  }
  const it = DriveApp.searchFiles(query);
  const arr = [];
  while (it.hasNext()) {
    const f = it.next();
    arr.push({ id: f.getId(), updated: f.getLastUpdated(), name: f.getName() });
  }
  if (arr.length === 0) return null;
  arr.sort((a, b) => b.updated - a.updated);
  for (let i = 0; i < arr.length; i++) {
    try {
      const ss = SpreadsheetApp.openById(arr[i].id);
      if (ss.getSheetByName(sheetName)) return arr[i].id;
    } catch (_) { }
  }
  return arr[0].id;
}
function getTracksSheet_() {
  const ssId = getTracksSpreadsheetId_();
  if (!ssId) throw new Error("Не задан источник треков: TRACKS_SPREADSHEET_ID или TRACKS_FOLDER_ID");
  const ss = SpreadsheetApp.openById(ssId);
  // если SHEET_NAME пуст/не найден — берём первый лист
  let sh = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : null;
  if (!sh) sh = ss.getSheets()[0];
  if (!sh) throw new Error("В файле треков нет листов.");
  return sh;
}
function findTrackRow_(codeRaw) {
  const code = String(codeRaw || "").trim().toUpperCase();
  if (!code) return null;
  const sh = getTracksSheet_();
  const last = sh.getLastRow();
  if (last < 1) return null;
  const rng = sh.getRange(1, 1, last, 1);
  const found = rng.createTextFinder(code).matchEntireCell(true).findNext();
  return found ? found.getRow() : null;
}

function getSheetsInFolder_(folderId, pattern) {
  let query = `mimeType = 'application/vnd.google-apps.spreadsheet' and '${folderId}' in parents`;
  if (pattern) {
    const safe = String(pattern).replace(/'/g, "\\'");
    // В DriveApp.searchFiles используется поле title, не name
    query += ` and title contains '${safe}'`;
  }
  try {
    const it = DriveApp.searchFiles(query);
    const files = [];
    while (it.hasNext()) {
      const f = it.next();
      files.push({ id: f.getId(), updated: f.getLastUpdated(), name: f.getName() });
    }
    files.sort((a, b) => b.updated - a.updated); // свежие — первыми
    return files;
  } catch (e) {
    Logger.log("Drive search error: " + e);
    return [];
  }
}

function getFirstSheetFromFile_(fileId) {
  const ss = SpreadsheetApp.openById(fileId);
  let sh = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : null;
  if (!sh) sh = ss.getSheets()[0];
  if (!sh) throw new Error("В файле нет листов: " + fileId);
  return sh;
}

/**
/**
 * Выполняет массовый поиск трек-кодов и возвращает информацию о них.
 * ВОЗВРАЩАЕТ ИСХОДНЫЙ ТЕКСТ из ячейки с датой (колонка B).
 */
function findTracksInfoBulk_(codesRaw) {
  const originalList = (codesRaw || []).map(s => String(s || "").trim()).filter(Boolean);
  const normCodes = originalList.map(normalizeCode_);
  const want = new Set(normCodes);
  const results = {};

  normCodes.forEach(nc => {
    if (!results[nc]) results[nc] = { found: false, dateText: null };
  });

  const searchInData = (data) => {
    for (let r = 0; r < data.length; r++) {
      const code = data[r][0];
      const dateValue = data[r][1]; // Значение из ячейки B
      const nc = normalizeCode_(code);
      
      if (nc && want.has(nc)) {
        // --- ОТЛАДОЧНЫЙ ЛОГ ---
        debugLog("findTracksInfoBulk", `[${nc}] Прочитано из ячейки`, dateValue);
        
        results[nc] = { found: true, dateText: dateValue || null }; 
        want.delete(nc);
      }
    }
  };

  if (TRACKS_FOLDER_ID) {
    const files = getSheetsInFolder_(TRACKS_FOLDER_ID, TRACKS_FILE_PATTERN);
    for (let i = 0; i < files.length; i++) {
      if (want.size === 0) break;
      try {
        const sh = getFirstSheetFromFile_(files[i].id);
        const last = sh.getLastRow();
        if (last >= 1) {
          const data = sh.getRange(1, 1, last, 2).getValues();
          searchInData(data);
        }
      } catch (e) { Logger.log(`findTracksInfoBulk_ skip file ${files[i].id} error: ${e}`); }
    }
    return results;
  }

  try {
    const sh = getTracksSheet_();
    const last = sh.getLastRow();
    if (last >= 1) {
      const data = sh.getRange(1, 1, last, 2).getValues();
      searchInData(data);
    }
  } catch (e) { Logger.log(`findTracksInfoBulk_ single-file error: ${e}`); }
  
  return results;
}

function findTrackExists_(codeRaw) {
  const code = normalizeCode_(codeRaw);
  if (!code) return false;

  // По папке: проверяем все файлы
  if (TRACKS_FOLDER_ID) {
    const files = getSheetsInFolder_(TRACKS_FOLDER_ID, TRACKS_FILE_PATTERN);
    for (let i = 0; i < files.length; i++) {
      try {
        const sh = getFirstSheetFromFile_(files[i].id);
        const last = sh.getLastRow();
        if (last < 1) continue;
        const col = sh.getRange(1, 1, last, 1).getValues().flat();
        for (let r = 0; r < col.length; r++) {
          if (normalizeCode_(col[r]) === code) return true;
        }
      } catch (e) {
        Logger.log("findTrackExists_ skip file " + files[i].id + " error: " + e);
      }
    }
    return false;
  }

  // Один файл
  try {
    const sh = getTracksSheet_();
    const last = sh.getLastRow();
    if (last < 1) return false;
    const col = sh.getRange(1, 1, last, 1).getValues().flat();
    for (let r = 0; r < col.length; r++) {
      if (normalizeCode_(col[r]) === code) return true;
    }
    return false;
  } catch (e) {
    Logger.log("findTrackExists_ single-file error: " + e);
    return false;
  }
}

/* --------- Подписки на треки --------- */
function getUserTracksSheet_() {
  const ss = SpreadsheetApp.openById(USERS_SPREADSHEET_ID);
  let sh = ss.getSheetByName("UserTracks");
  if (!sh) sh = ss.insertSheet("UserTracks");
  const header = ["userId", "code", "lastStatus", "lastNotified", "active", "lastReminder"];
  const first = sh.getRange(1, 1, 1, header.length).getValues()[0];
  const empty = first.every(v => v === "" || v === null || v === undefined);
  if (empty) sh.getRange(1, 1, 1, header.length).setValues([header]);
  return sh;
}
function isSubscribed_(userId, code) {
  const sh = getUserTracksSheet_();
  const last = sh.getLastRow(); if (last < 2) return false;
  const rng = sh.getRange(2, 1, last - 1, 2); // userId, code
  const tf = rng.createTextFinder(String(code)).matchEntireCell(true).findAll();
  if (!tf || tf.length === 0) return false;
  for (let i = 0; i < tf.length; i++) {
    const r = tf[i].getRow();
    const uid = String(sh.getRange(r, 1).getValue());
    const active = sh.getRange(r, 5).getValue();
    if (uid === String(userId) && String(active) === "1") return true;
  }
  return false;
}
/*******************************************************
 *  ПОДПИСКИ НА ТРЕКИ (ПЕРЕПИСАННАЯ ЛОГИКА)
 *******************************************************/

/**
 * Создает или повторно активирует подписку на трек-код.
 * Устанавливает начальный статус 'PENDING'.
 */
function subscribeToTrack_(userId, code) {
  const sh = getUserTracksSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const last = sh.getLastRow();
    // Ищем, есть ли уже запись для этой пары userId + code
    if (last >= 2) {
      const allData = sh.getRange(2, 1, last - 1, 2).getValues(); // userId, code
      for (let i = 0; i < allData.length; i++) {
        if (String(allData[i][0]) === String(userId) && normalizeCode_(allData[i][1]) === normalizeCode_(code)) {
          const r = i + 2; // +2 потому что getRange с 2-й строки, а массив с 0
          // Если нашли - просто реактивируем и сбрасываем статус
          sh.getRange(r, 3, 1, 3).setValues([
            ['PENDING', new Date().toISOString(), '1'] // lastStatus, lastNotified, active
          ]);
          return;
        }
      }
    }
    // Если не нашли - добавляем новую строку
    sh.appendRow([
      String(userId),
      String(code).trim().toUpperCase(),
      'PENDING',                   // lastStatus
      new Date().toISOString(),    // lastNotified (дата подписки)
      '1',                         // active
      ''                           // lastReminder (не используется, но оставляем для структуры)
    ]);
  } finally {
    lock.releaseLock();
  }
}


// ЗАМЕНИТЕ ВАШУ СТАРУЮ parseDateFromText_ НА ЭТУ
function parseDateFromText_(txt) {
  if (txt instanceof Date && !isNaN(txt)) return txt;
  if (!txt) return null;
  
  // Берем только первую часть до пробела или тире
  let dateString = String(txt).trim().split(/[\s\-–—]/)[0];
  let parts;

  // Формат: ДД.ММ.ГГГГ или ДД/ММ/ГГГГ
  parts = dateString.match(/^(\d{1,2})[.\/](\d{1,2})[.\/](\d{4})$/);
  if (parts) {
    const day = parseInt(parts[1], 10), month = parseInt(parts[2], 10), year = parseInt(parts[3], 10);
    if (year > 1970 && month >= 1 && month <= 12) {
      const dt = new Date(Date.UTC(year, month - 1, day));
      if (dt.getUTCMonth() === month - 1) return dt;
    }
  }

  // Формат: ГГГГ-ММ-ДД или ГГГГ.ММ.ДД
  parts = dateString.match(/^(\d{4})[-.\/](\d{1,2})[-.\/](\d{1,2})$/);
  if (parts) {
    const year = parseInt(parts[1], 10), month = parseInt(parts[2], 10), day = parseInt(parts[3], 10);
    if (year > 1970 && month >= 1 && month <= 12) {
      const dt = new Date(Date.UTC(year, month - 1, day));
      if (dt.getUTCMonth() === month - 1) return dt;
    }
  }

  return null; // Если ни один из надежных форматов не подошел
}


// ЗАМЕНИТЕ ВАШУ СТАРУЮ checkSubscriptions НА ЭТУ
function checkSubscriptions() {
  const sh = getUserTracksSheet_();
  const last = sh.getLastRow();
  if (last < 2) return;

  const data = sh.getRange(2, 1, last - 1, 5).getValues(); 

  const pendingSubs = [];
  for (let i = 0; i < data.length; i++) {
    const [userId, code, lastStatus, _, active] = data[i];
    if (String(active) === '1' && String(lastStatus).toUpperCase() === 'PENDING') {
      pendingSubs.push({ rowIndex: i + 2, userId: String(userId), code: String(code) });
    }
  }

  if (pendingSubs.length === 0) return;

  const codesToCheck = pendingSubs.map(sub => sub.code);
  const searchResults = findTracksInfoBulk_(codesToCheck);

  const updatesToSheet = []; 

  for (const sub of pendingSubs) {
    const normCode = normalizeCode_(sub.code);
    const result = searchResults[normCode];

    if (result && result.found) {
      const user = getUserFromCsv(sub.userId);
      if (!user) continue; 

      const L = TEXT[user.lang || "ru"];
      
      // ИСПРАВЛЕНИЕ: Используем новую надежную функцию для парсинга
      const arrivalDate = parseDateFromText_(result.dateText);

      try {
        sendMessage(sub.userId, L.status_update_initial(sub.code, arrivalDate));
        
        updatesToSheet.push({
          rowIndex: sub.rowIndex,
          newStatus: arrivalDate ? arrivalDate.toISOString() : new Date().toISOString(),
          newNotifyDate: new Date().toISOString()
        });

        logEvent_("sub_found_notified", sub.userId, `code=${sub.code}`);
      } catch (e) {
        Logger.log(`Не удалось отправить уведомление для ${sub.userId} по треку ${sub.code}: ${e}`);
      }
    }
  }

  if (updatesToSheet.length > 0) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      for (const update of updatesToSheet) {
        sh.getRange(update.rowIndex, 3, 1, 2).setValues([[update.newStatus, update.newNotifyDate]]);
      }
    } finally {
      lock.releaseLock();
    }
    Logger.log(`Обработано и обновлено ${updatesToSheet.length} подписок.`);
  }
}
function unsubscribeFromTrack_(userId, code) {
  const sh = getUserTracksSheet_();
  const last = sh.getLastRow(); if (last < 2) return;
  const rng = sh.getRange(2, 1, last - 1, 2);
  const tf = rng.createTextFinder(String(code)).matchEntireCell(true).findAll();
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    for (let i = 0; i < (tf ? tf.length : 0); i++) {
      const r = tf[i].getRow();
      const uid = String(sh.getRange(r, 1).getValue());
      if (uid === String(userId)) {
        sh.getRange(r, 5).setValue("0");
      }
    }
  } finally { lock.releaseLock(); }
}

function setupSubscriptionChecker_() {
  // удалить старые
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "checkSubscriptions") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("checkSubscriptions").timeBased().everyMinutes(15).create();
}

/* --------- Админ/Логи --------- */
function getAdminIds_() {
  const raw = PROPS.getProperty("ADMIN_IDS") || PROPS.getProperty("ADMIN_ID") || "";
  return raw.split(/[,\s]+/).map(s => s.trim()).filter(Boolean).map(String);
}
function isAdmin_(userId) {
  return getAdminIds_().includes(String(userId));
}
function countUsers_() {
  const sh = getUsersSheet_(); const last = sh.getLastRow(); return Math.max(0, last - 1);
}
function countActiveSubs_() {
  const sh = getUserTracksSheet_(); const last = sh.getLastRow(); if (last < 2) return 0;
  const vals = sh.getRange(2, 5, last - 1, 1).getValues().flat();
  return vals.filter(v => String(v) === "1").length;
}
function getLogsSheet_() {
  const ss = SpreadsheetApp.openById(USERS_SPREADSHEET_ID);
  let sh = ss.getSheetByName("Logs");
  if (!sh) sh = ss.insertSheet("Logs");
  const header = ["timestamp", "type", "userId", "details"];
  const first = sh.getRange(1, 1, 1, header.length).getValues()[0];
  const empty = first.every(v => v === "" || v === null || v === undefined);
  if (empty) sh.getRange(1, 1, 1, header.length).setValues([header]);
  return sh;
}
function logEvent_(type, userId, details) {
  try {
    const sh = getLogsSheet_();
    sh.appendRow([new Date().toISOString(), String(type), String(userId), String(details || "")]);
  } catch (_) { }
}

/* --------- Проверка адреса--------- */

function getBoolProp_(key, def) {
  var v = PROPS.getProperty(key);
  if (v === null || v === undefined || v === "") return !!def;
  v = String(v).trim().toLowerCase();
  return v === "1" || v === "true" || v === "yes" || v === "on";
}

function getCnAddrValidationOpts_() {
  return {
    customerPrefix: (PROPS.getProperty("CN_ADDR_VALID_PREFIX") || "JR-"),
    codeMin: parseInt(PROPS.getProperty("CN_ADDR_VALID_CODE_MIN") || "3", 10),
    codeMax: parseInt(PROPS.getProperty("CN_ADDR_VALID_CODE_MAX") || "6", 10),
    requireWarehouseChar: getBoolProp_("CN_REQUIRE_WAREHOUSE_CHAR", false),
    requireYiwu: getBoolProp_("CN_REQUIRE_YIWU", false),
    expectCodeTwice: getBoolProp_("CN_EXPECT_CODE_TWICE", false)
  };
}

function norm(s) { return String(s || "").replace(/[\u200B-\u200D\uFEFF]/g, "").replace(/[，,]/g, " ").replace(/\s+/g, " ").trim(); }
function onlyDigits(s) { return String(s || "").replace(/\D+/g, ""); }
function countOccur(text, sub) { if (!text || !sub) return 0; var c = 0, p = 0, i; while ((i = text.indexOf(sub, p)) !== -1) { c++; p = i + sub.length; } return c; }
function extractChineseCore(expected) { var m = String(expected || "").match(/[\u4e00-\u9fff0-9号栋幢单元楼室仓—-\s]+/g) || []; var best = "", bestLen = 0; for (var i = 0; i < m.length; i++) { if (/[\u4e00-\u9fff]/.test(m[i]) && m[i].length > bestLen) { best = m[i]; bestLen = m[i].length; } } return norm(best); }

// Базовая "общая" проверка (структура адреса)
function validateChinaAddress(raw, opts, L) {
  opts = opts || {};
  var res = { ok: true, issues: [], data: {} };
  if (!raw || typeof raw !== 'string') {
    res.ok = false; res.issues.push(L.addr_validation.empty_address); return res;
  }
  var s = norm(raw);
  if (!/[\u4e00-\u9fff]/.test(s)) { res.ok = false; res.issues.push(L.addr_validation.no_chinese_chars); }
  
  // --- НАЧАЛО ИЗМЕНЕНИЙ ---
  // Получаем шаблон из настроек. Если его нет, используем старый подход как fallback.
  const regexPattern = PROPS.getProperty("CN_ADDR_VALID_REGEX");
  const reCode = regexPattern 
    ? new RegExp(regexPattern, 'ig') 
    : new RegExp((opts.customerPrefix || 'JR-').replace(/[-/\^$*+?.()|[```{}]/g, '\$&') + '(\\d{' + (opts.codeMin || 3) + ',' + (opts.codeMax || 6) + '})', 'ig');
  // --- КОНЕЦ ИЗМЕНЕНИЙ ---

  var set = {}, arr = [], m; while ((m = reCode.exec(s)) !== null) { 
    // m[0] будет содержать полное совпадение (например, "union3")
    var c = m[0].toUpperCase(); 
    if (!set[c]) { set[c] = true; arr.push(c); } 
  }
  
  if (arr.length === 0) res.issues.push(L.addr_validation.client_code_not_found('')); // Убрали 'XXXX'
  else res.data.customerCodes = arr;

  // ... остальная часть функции без изменений ...
  var cnPhoneMatch = s.match(/\b1[3-9]\d{9}\b/);
  if (!cnPhoneMatch) res.issues.push(L.addr_validation.cn_phone_not_found); else res.data.cnPhone = cnPhoneMatch[0];
  if (!(/省/.test(s) && /市/.test(s) || /(北京市|上海市|天津市|重庆市)/.test(s) || /(香港|澳门)/.test(s))) res.issues.push(L.addr_validation.region_not_found);
  if (opts.requireYiwu && !/义乌市/.test(s)) res.issues.push(L.addr_validation.not_yiwu);
  if (!/(区|县|镇|乡|路|街|道|村|巷|号|栋|幢|单元|楼|室|仓)/.test(s)) res.issues.push(L.addr_validation.details_not_found);
  if (opts.requireWarehouseChar && !/仓/.test(s)) res.issues.push(L.addr_validation.warehouse_char_missing);
  var regionMatch = s.match(/[\u4e00-\u9fff]+省[\u4e00-\u9fff]+市[\u4e00-\u9fff]+(市|区|县)/) || s.match(/(北京市|上海市|天津市|重庆市)[\u4e00-\u9fff]+(区|县|市)/); if (regionMatch) res.data.region = regionMatch[0];
  res.ok = res.issues.length === 0;
  return res;
}

// Сравнение с эталоном из шаблона (properties)
// L добавлен в параметры
// Сравнение с эталоном из шаблона (properties) - ИСПРАВЛЕННАЯ ВЕРСИЯ
function validateAgainstTemplate_(raw, expectedAddr, code, cnPhone, userName, userPhone, opts, L) {
  opts = opts || {};
  var res = { ok: true, issues: [] };
  var s = norm(raw); var e = norm(expectedAddr); var sU = s.toUpperCase(); var eU = e.toUpperCase();
  var codeU = String(code || "").toUpperCase();
  
  // Проверка №1: Присутствует ли код самого пользователя
  var occUser = countOccur(sU, codeU);
  if (occUser === 0) res.issues.push(L.addr_validation.client_code_not_found(code));
  if (opts.expectCodeTwice && occUser < 2) res.issues.push(L.addr_validation.client_code_twice(code));

  // ====================== НАЧАЛО ИСПРАВЛЕНИЯ ======================
  // Проверка №2: Нет ли в адресе посторонних кодов (используем ту же гибкую логику)
  const regexPattern = PROPS.getProperty("CN_ADDR_VALID_REGEX");
  if (regexPattern) {
    const reCode = new RegExp(regexPattern, 'ig');
    let foundCodes = [];
    let match;
    while ((match = reCode.exec(sU)) !== null) {
      if (foundCodes.indexOf(match[0]) === -1) {
          foundCodes.push(match[0]);
      }
    }
    const otherCodes = foundCodes.filter(c => c !== codeU);
    if (otherCodes.length > 0) {
      res.issues.push(L.addr_validation.other_code_found(otherCodes.join(", ")));
    }
  }
  // ======================= КОНЕЦ ИСПРАВЛЕНИЯ =======================

  // Проверка №3: Китайский номер телефона
  var cnM = s.match(/\b1[3-9]\d{9}\b/g) || [];
  if (cnM.length === 0) { res.issues.push(L.addr_validation.cn_phone_not_found); } else if (cnPhone && cnM[0] !== cnPhone) { res.issues.push(L.addr_validation.cn_phone_mismatch(cnPhone, cnM[0])); }
  
  // Проверка №4: Основная китайская часть адреса
  var core = extractChineseCore(e); if (core) { if (s.replace(/\s+/g, "").indexOf(core.replace(/\s+/g, "")) === -1) res.issues.push(L.addr_validation.chinese_part_mismatch); }
  
  // Проверка №5: Требования типа "Иу" и иероглифа "склад"
  if (opts.requireYiwu && !/义乌市/.test(s)) res.issues.push(L.addr_validation.not_yiwu);
  if (opts.requireWarehouseChar && s.indexOf("仓") === -1) res.issues.push(L.addr_validation.warehouse_char_missing);
  
  // Проверка №6: Обязательные ключевые слова из настроек
  const requiredKeywordsRaw = PROPS.getProperty("CN_ADDR_REQUIRED_KEYWORDS") || "";
  const requiredKeywords = requiredKeywordsRaw.split(',').map(k => k.trim()).filter(Boolean);
  for (const keyword of requiredKeywords) {
    const isNumeric = /^\d+$/.test(keyword);
    const searchPattern = isNumeric ? new RegExp(`\\b${keyword}\\b`) : new RegExp(keyword, 'i');
    if (!searchPattern.test(s)) {
      res.issues.push(L.addr_validation.keyword_missing(keyword));
    }
  }

  // Проверка №7: Персональные данные (имя и телефон)
  var nm = String(userName || "").trim(); if (nm) { if (e.toLowerCase().indexOf(nm.toLowerCase()) !== -1 && s.toLowerCase().indexOf(nm.toLowerCase()) === -1) { res.issues.push(L.addr_validation.name_not_found(nm)); } }
  var pDigits = onlyDigits(userPhone); if (pDigits) { if (onlyDigits(e).indexOf(pDigits) !== -1 && onlyDigits(s).indexOf(pDigits) === -1) { res.issues.push(L.addr_validation.phone_not_found(pDigits)); } }
  
  res.ok = res.issues.length === 0;
  return res;
}

// Итоговый отчёт + показ правильного адреса при ошибках
function formatValidationReport_(res, lang) {
  var L = TEXT[lang] || TEXT.ru;
  if (res.ok) return L.addr_ok;
  return L.addr_fail_header + "\n- " + res.issues.join("\n- ");
}

/**
 * Проверяет адрес. Если находит ошибку, отправляет отчет, правильный адрес И видеоинструкцию.
 */
function checkChinaAddressAndReply_(chatId, userId, raw, lang) {
  const L = TEXT[lang] || TEXT.ru;
  try {
    var user = getUserFromCsv(String(userId));
    if (!user) {
      sendMessage(chatId, L.addr_validation.please_register);
      return;
    }

    // --- Шаг 1: Выполняем проверку адреса (как и раньше) ---
    var wh = user.warehouse;
    var gd = getChinaData(wh, user.name, user.phone);
    var expected = gd.addr; 
    var code = gd.code; 
    var cnPhone = PROPS.getProperty("CN_PHONE_COMMON") || gd.phoneCn || null;
    var opts = getCnAddrValidationOpts_();
    var codeDigits = onlyDigits(code).length; 
    var codePrefix = String(code).replace(/\d.*$/, '').toUpperCase();
    if (codeDigits) { opts.codeMin = codeDigits; opts.codeMax = codeDigits; }
    if (codePrefix) { opts.customerPrefix = codePrefix; }

    var base = validateChinaAddress(raw, opts, L);
    var tpl = validateAgainstTemplate_(raw, expected, code, cnPhone, user.name, user.phone, opts, L);
    
    var issues = []; 
    var seen = {};
    [].concat(base.issues || [], tpl.issues || []).forEach(function (x) { if (!x) return; if (!seen[x]) { seen[x] = true; issues.push(x); } });
    var res = { ok: issues.length === 0, issues: issues };

    // --- Шаг 2: Отправляем результат проверки ---
    var out = formatValidationReport_(res, lang);
    sendMessage(chatId, out);

    // ====================== НАЧАЛО ИЗМЕНЕНИЙ ======================
    // ШАГ 3: Если результат проверки НЕ "ok" (т.е. есть ошибки)
    if (!res.ok) {
      // 1. Отправляем сообщение с правильным адресом для копирования
      sendMessage(chatId, L.addr_correct);
      sendCopyableText(chatId, expected);
      
      // 2. Отправляем видеоинструкцию для склада этого пользователя
      const videoId = getWhProp("PINDUODUO_VIDEO", user.warehouse);
      if (videoId) {
        sendVideoFromDrive(chatId, videoId);
      }
    }
    // ======================= КОНЕЦ ИЗМЕНЕНИЙ =======================

  } catch (e) {
    Logger.log("addr check error: " + e);
    sendMessage(chatId, L.addr_validation.generic_error);
  }
}

/* --------- Сообщения, клавиатуры и медиа --------- */
function sendMessage(chatId, text) {
	try {
  const url = `https://api.telegram.org/bot${TOKEN}/sendMessage`;
  const payload = { chat_id: chatId, text };
  UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true // <-- ВАЖНО: не дает скрипту упасть при ошибках 4xx/5xx 
  });
  } catch (e) {
    Logger.log(`Не удалось отправить сообщение в чат ${chatId}: ${e.message}`);
  }
}

function sendMessageWithKeyboardRemove(chatId, text) {
	try {
  const url = `https://api.telegram.org/bot${TOKEN}/sendMessage`;
  const payload = { chat_id: chatId, text, reply_markup: JSON.stringify({ remove_keyboard: true }) };
  UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true // <-- ВАЖНО: не дает скрипту упасть при ошибках 4xx/5xx
    });
  } catch (e) {
    Logger.log(`Не удалось отправить сообщение в чат ${chatId}: ${e.message}`);
  }
}
function sendMessageWithInlineKb_(chatId, text, replyMarkupObj) {
	try {
  const url = `https://api.telegram.org/bot${TOKEN}/sendMessage`;
  const payload = { chat_id: chatId, text, reply_markup: JSON.stringify(replyMarkupObj) };
  UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", payload: JSON.stringify(payload),
      muteHttpExceptions: true // <-- ВАЖНО: не дает скрипту упасть при ошибках 4xx/5xx
    });
  } catch (e) {
    Logger.log(`Не удалось отправить сообщение в чат ${chatId}: ${e.message}`);
  }
}
function inlineKb_(rows) {
  return { inline_keyboard: rows.map(r => r.map(b => ({ text: b.text, callback_data: b.data }))) };
}
function sendLanguageChoice(chatId, lang) {
	try {
  const L = TEXT[lang] || TEXT.ru;
  const url = `https://api.telegram.org/bot${TOKEN}/sendMessage`;
  const payload = {
    chat_id: chatId,
    text: L.choose_lang,
    reply_markup: JSON.stringify({
      keyboard: [[{ text: TEXT.ru.choose_ru }, { text: TEXT.ru.choose_tj }]],
      one_time_keyboard: true,
      resize_keyboard: true
    })
  };
  UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", payload: JSON.stringify(payload),
      muteHttpExceptions: true // <-- ВАЖНО: не дает скрипту упасть при ошибках 4xx/5xx
    });
  } catch (e) {
    Logger.log(`Не удалось отправить сообщение в чат ${chatId}: ${e.message}`);
  }
}
function sendContactButton(chatId, lang) {
	try {
  const L = TEXT[lang] || TEXT.tj;
  const url = `https://api.telegram.org/bot${TOKEN}/sendMessage`;
  const payload = {
    chat_id: chatId,
    text: L.ask_phone,
    reply_markup: JSON.stringify({
      keyboard: [[{ text: L.btn_share, request_contact: true }], [{ text: L.btn_cancel }]],
      resize_keyboard: true,
      one_time_keyboard: true
    })
  };
  UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", payload: JSON.stringify(payload),
      muteHttpExceptions: true // <-- ВАЖНО: не дает скрипту упасть при ошибках 4xx/5xx
    });
  } catch (e) {
    Logger.log(`Не удалось отправить сообщение в чат ${chatId}: ${e.message}`);
  }
}
function sendCancelKeyboard_(chatId, text, lang) {
	try {
  const L = TEXT[lang] || TEXT.tj;
  const url = `https://api.telegram.org/bot${TOKEN}/sendMessage`;
  const payload = {
    chat_id: chatId,
    text,
    reply_markup: JSON.stringify({ keyboard: [[{ text: L.btn_cancel }]], resize_keyboard: true, one_time_keyboard: true })
  };
  UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", payload: JSON.stringify(payload),
      muteHttpExceptions: true // <-- ВАЖНО: не дает скрипту упасть при ошибках 4xx/5xx
    });
  } catch (e) {
    Logger.log(`Не удалось отправить сообщение в чат ${chatId}: ${e.message}`);
  }
}
function sendMenu(chatId, lang) {
	try {
  const L = TEXT[lang] || TEXT.tj;
  const url = `https://api.telegram.org/bot${TOKEN}/sendMessage`;
  const kb = [
    [{ text: L.btn_check_track }],
    [{ text: L.btn_addr_cn }, { text: L.btn_change_wh }],
    [{ text: L.btn_addr_khu }, { text: L.btn_price }],
    [{ text: L.btn_time }, { text: L.btn_banned }],
    [{ text: L.btn_admin }, { text: L.btn_profile }, { text: L.btn_check_addr_cn }],
  ];
  const payload = { chat_id: chatId, text: L.menu_title, reply_markup: JSON.stringify({ keyboard: kb, resize_keyboard: true }) };
  UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", payload: JSON.stringify(payload),
      muteHttpExceptions: true // <-- ВАЖНО: не дает скрипту упасть при ошибках 4xx/5xx
    });
  } catch (e) {
    Logger.log(`Не удалось отправить сообщение в чат ${chatId}: ${e.message}`);
  }
}
function sendAdminContact(chatId, lang) {
	try {
  const L = TEXT[lang] || TEXT.tj;
  const url = `https://api.telegram.org/bot${TOKEN}/sendMessage`;
  const payload = {
    chat_id: chatId,
    text: L.admin_caption,
    reply_markup: JSON.stringify({ inline_keyboard: [[{ text: lang === "ru" ? "Чат с админом" : "Чат бо админ", url: ADMIN_LINK }]] })
  };
  UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", payload: JSON.stringify(payload),
      muteHttpExceptions: true // <-- ВАЖНО: не дает скрипту упасть при ошибках 4xx/5xx
    });
  } catch (e) {
    Logger.log(`Не удалось отправить сообщение в чат ${chatId}: ${e.message}`);
  }
}
// ЗАМЕНИТЕ ВАШУ ФУНКЦИЮ НА ЭТУ
function sendPhotoFromDrive(chatId, fileId, caption) {
  if (!fileId) {
    sendMessage(chatId, caption || "");
    return;
  }
  try {
    const f = DriveApp.getFileById(fileId);
    try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { Logger.log("Не удалось изменить права доступа для файла: " + e); }

    const photoUrl = `https://drive.google.com/uc?export=view&id=${fileId}`;
    const url = `https://api.telegram.org/bot${TOKEN}/sendPhoto`;
    const payload = { 
      chat_id: String(chatId), 
      caption: (caption || ""), 
      photo: photoUrl 
    };

    UrlFetchApp.fetch(url, { 
      method: "post", 
      contentType: "application/json", 
      payload: JSON.stringify(payload),
      muteHttpExceptions: true // <-- Добавлено
    });
  } catch (e) {
    Logger.log(`Ошибка отправки фото с Drive (fileId: ${fileId}) в чат ${chatId}: ${e.message}`);
    // Fallback: если отправка фото не удалась, отправляем хотя бы текст
    sendMessage(chatId, caption || "");
  }
}

/**
 * Отправляет видео с Google Диска в чат.
 * @param {string|number} chatId - ID чата.
 * @param {string} fileId - ID видеофайла на Google Диске.
 * @param {string} [caption] - Необязательный текст под видео.
 * @returns {boolean} true если успешно, false если ошибка
 */
function sendVideoFromDrive(chatId, fileId, caption) {
  if (!fileId) {
    Logger.log("sendVideoFromDrive: fileId не предоставлен.");
    return false;
  }

  try {
    // Делаем файл доступным по ссылке
    const f = DriveApp.getFileById(fileId);
    try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { }

    const videoUrl = `https://drive.google.com/uc?export=view&id=${fileId}`;
    const url = `https://api.telegram.org/bot${TOKEN}/sendVideo`;

    const payload = {
      chat_id: String(chatId),
      video: videoUrl,
      caption: caption || ""
    };

    const response = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    if (responseCode === 200) {
      return true;
    } else {
      Logger.log(`Telegram API error (${responseCode}): ${response.getContentText()}`);
      return false;
    }

  } catch (e) {
    Logger.log(`Ошибка отправки видео (fileId: ${fileId}): ${e.message}`);
    return false;
  }
}

// ЗАМЕНИТЕ ВАШУ ФУНКЦИЮ НА ЭТУ
function answerCallbackQuery(id, text) { // <-- Убран '_'
  try {
    const url = `https://api.telegram.org/bot${TOKEN}/answerCallbackQuery`;
    const payload = { callback_query_id: id, text: text || "" };
    UrlFetchApp.fetch(url, { 
      method: "post", 
      contentType: "application/json", 
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log(`Не удалось ответить на callback_query (id: ${id}): ${e.message}`); // <-- Исправлен лог
  }
}

/* --------- Копируемые сообщения (адрес в Китае) --------- */
function sendCopyableText(chatId, text) {
	try {
  const url = `https://api.telegram.org/bot${TOKEN}/sendMessage`;
  const payload = { chat_id: chatId, text: `<pre>${escapeHtml(String(text))}</pre>`, parse_mode: "HTML" };
  UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", payload: JSON.stringify(payload),
      muteHttpExceptions: true // <-- ВАЖНО: не дает скрипту упасть при ошибках 4xx/5xx
    });
  } catch (e) {
    Logger.log(`Не удалось отправить копируемое сообщение в чат ${chatId}: ${e.message}`);
  }
}
function sendCopyableAddressCn(chatId, name, phone, warehouse) {
  const wh = warehouse || "ZAFAROBOD";
  const { addr, fileId } = getChinaData(wh, name, phone);
  sendPhotoFromDrive(chatId, fileId || FILE_ID_ADDRESS_CHINA, addr);
  sendCopyableText(chatId, addr);
}
function escapeHtml(s) { return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }

/* --------- Профиль --------- */
function showProfile_(chatId, user, lang) {
  const L = TEXT[lang] || TEXT.ru;
  const txt = `${L.profile_title}\n\n${L.profile_name}: ${user.name}\n${L.profile_phone}: ${user.phone}\n${L.profile_lang}: ${lang.toUpperCase()}`;
  const kb = {
    inline_keyboard: [
      [{ text: L.profile_edit_name, callback_data: "edit_name" }],
      [{ text: L.profile_edit_phone, callback_data: "edit_phone" }],
      [{ text: L.profile_change_lang, callback_data: "change_lang" }],
      [{ text: L.profile_change_wh, callback_data: "change_wh" }], // новая кнопка
    ]
  };
  sendMessageWithInlineKb_(chatId, txt, kb);
}

// Поместите эту функцию в конец вашего скрипта

/**
 * Архивирует файлы с треками старше 30 дней.
 * Запускается по триггеру раз в неделю.
 */
function archiveOldTrackFiles() {
  if (!TRACKS_FOLDER_ID) {
    Logger.log("TRACKS_FOLDER_ID не задан. Архивация пропущена.");
    return;
  }

  try {
    const folder = DriveApp.getFolderById(TRACKS_FOLDER_ID);
    const archive = getOrCreateSubfolder_(folder, "_archive"); // Ваша функция getOrCreateSubfolder_ уже есть
    
    const thirtyDaysAgo = new Date();
    thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);

    const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      
      // Пропускаем Google Таблицы, которые были созданы недавно
      // (на случай если конвертация занимает время)
      if (file.getLastUpdated() < thirtyDaysAgo) {
        archive.addFile(file); // Перемещаем в архив
        folder.removeFile(file); // Удаляем из основной папки
        Logger.log(`Архивирован старый файл: ${file.getName()}`);
      }
    }
  } catch (e) {
    Logger.log(`Ошибка при архивации файлов: ${e}`);
    logEvent_("error", "system", `File archiving failed: ${e.message}`);
  }
}

/**
 * Записывает отладочную информацию в лист 'DebugLogs'.
 * @param {string} context - Название функции или процесса (например, "findTracksInfoBulk").
 * @param {string} key - Название переменной или события (например, "dateValue from cell").
 * @param {*} value - Значение для логирования.
 */
function debugLog(context, key, value) {
  try {
    const ss = SpreadsheetApp.openById(USERS_SPREADSHEET_ID);
    const sheet = ss.getSheetByName("DebugLogs");
    if (!sheet) return; // Если листа нет, ничего не делаем

    const timestamp = new Date();
    let valueStr = "";
    
    // Преобразуем значение в строку для записи
    if (value === null) {
      valueStr = "null";
    } else if (value === undefined) {
      valueStr = "undefined";
    } else if (value instanceof Date) {
      valueStr = value.toISOString();
    } else if (typeof value === 'object') {
      valueStr = JSON.stringify(value);
    } else {
      valueStr = String(value);
    }

    // Ограничиваем длину значения, чтобы не сломать ячейку
    if (valueStr.length > 50000) {
      valueStr = valueStr.substring(0, 50000) + "... (truncated)";
    }
    
    const type = (value instanceof Date) ? "Date" : typeof value;

    sheet.appendRow([timestamp, context, key, valueStr, type]);

  } catch (e) {
    // Чтобы основная логика бота не прерывалась из-за ошибки логирования
    Logger.log(`Failed to write debug log: ${e.toString()}`);
  }
}

/**********************
 *  РАССЫЛКА (BROADCAST)
 *********************/

/**
 * Проверка, не запущена ли уже рассылка
 */
function isBroadcastRunning_() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty("BROADCAST_RUNNING") === "1";
}

/**
 * Установка флага рассылки
 */
function setBroadcastFlag_(running) {
  const props = PropertiesService.getScriptProperties();
  if (running) {
    props.setProperty("BROADCAST_RUNNING", "1");
  } else {
    props.deleteProperty("BROADCAST_RUNNING");
  }
}

/**
 * Массовая рассылка видео пользователям бота.
 * Отправляет видео из Google Drive с текстовой подписью.
 * Ограничение: макс. 20 видео в минуту (лимит Telegram)
 */
function broadcastVideo(videoFileId, caption, targetLang) {
  // Проверка: не запущена ли уже рассылка
  if (isBroadcastRunning_()) {
    Logger.log("broadcastVideo: рассылка уже запущена, пропускаем");
    notifyAdmins_("⚠️ Попытка запуска рассылки отклонена: другая рассылка уже выполняется");
    return { success: 0, failed: 0, error: "Рассылка уже запущена" };
  }
  
  // Устанавливаем флаг
  setBroadcastFlag_(true);
  
  try {
    if (!videoFileId) {
      Logger.log("broadcastVideo: не указан videoFileId");
      notifyAdmins_("❌ Ошибка рассылки: не указан ID видео");
      return { success: 0, failed: 0, error: "Не указан ID видео" };
    }

    const sheet = getUsersSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log("broadcastVideo: нет пользователей");
      return { success: 0, failed: 0, error: "Нет пользователей" };
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    
    let successCount = 0;
    let failedCount = 0;
    const results = [];

    Logger.log(`Начало рассылки видео. Всего пользователей: ${data.length}`);

    for (let i = 0; i < data.length; i++) {
      const [userId, name, phone, regDate, lang] = data[i];

      // Фильтр по языку если указан
      if (targetLang && String(lang) !== targetLang) {
        continue;
      }

      try {
        const sent = sendVideoFromDrive(String(userId), videoFileId, caption);
        if (sent !== false) {
          successCount++;
          results.push({ userId: String(userId), status: "sent" });
        } else {
          failedCount++;
          results.push({ userId: String(userId), status: "failed" });
        }

        // Пауза 3 секунды между видео (лимит Telegram ~20 видео/мин)
        Utilities.sleep(3000);

      } catch (e) {
        failedCount++;
        results.push({ userId: String(userId), status: "failed", error: String(e) });
        Logger.log(`Не удалось отправить пользователю ${userId}: ${e.message}`);

        // При ошибке тоже пауза
        Utilities.sleep(1000);
      }
    }

    Logger.log(`Рассылка видео завершена. Успешно: ${successCount}, Ошибок: ${failedCount}`);

    const report = `📊 Рассылка видео завершена!\n\n` +
      `✅ Успешно: ${successCount}\n` +
      `❌ Ошибок: ${failedCount}\n` +
      `📝 Всего отправлено: ${successCount + failedCount}`;

    notifyAdmins_(report);

    return { success: successCount, failed: failedCount, results: results };
    
  } finally {
    // Снимаем флаг после завершения (успешно или с ошибкой)
    setBroadcastFlag_(false);
  }
}

/**
 * Массовая текстовая рассылка пользователям бота.
 * Ограничение: макс. 30 сообщений в секунду
 */
function broadcastText(message, targetLang) {
  // Проверка: не запущена ли уже рассылка
  if (isBroadcastRunning_()) {
    Logger.log("broadcastText: рассылка уже запущена, пропускаем");
    notifyAdmins_("⚠️ Попытка запуска рассылки отклонена: другая рассылка уже выполняется");
    return { success: 0, failed: 0, error: "Рассылка уже запущена" };
  }
  
  // Устанавливаем флаг
  setBroadcastFlag_(true);
  
  try {
    const sheet = getUsersSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log("broadcastText: нет пользователей");
      return { success: 0, failed: 0, error: "Нет пользователей" };
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    
    let successCount = 0;
    let failedCount = 0;

    Logger.log(`Начало текстовой рассылки. Всего пользователей: ${data.length}`);

    for (let i = 0; i < data.length; i++) {
      const [userId, , , , lang] = data[i];
      
      if (targetLang && String(lang) !== targetLang) {
        continue;
      }

      try {
        sendMessage(String(userId), message);
        successCount++;
        
        // Пауза каждые 30 сообщений (0.5 сек)
        if ((i + 1) % 30 === 0) {
          Utilities.sleep(500);
        }
      } catch (e) {
        failedCount++;
        Logger.log(`Не удалось отправить пользователю ${userId}: ${e.message}`);
      }
    }

    Logger.log(`Текстовая рассылка завершена. Успешно: ${successCount}, Ошибок: ${failedCount}`);
    
    const report = `📊 Текстовая рассылка завершена!\n\n` +
      `✅ Успешно: ${successCount}\n` +
      `❌ Ошибок: ${failedCount}\n` +
      `📝 Всего: ${successCount + failedCount}`;
    
    notifyAdmins_(report);
    
    return { success: successCount, failed: failedCount };
    
  } finally {
    // Снимаем флаг после завершения
    setBroadcastFlag_(false);
  }
}

/**
 * Команда для запуска рассылки через чат (только для админа).
 * Использование: /broadcast video <fileId> <текст>
 * Или: /broadcast text <текст>
 * Или: /broadcast ru <текст> (только русскоязычным)
 * Или: /broadcast tj <текст> (только таджикскоязычным)
 */
function handleBroadcastCommand(chatId, userId, args) {
  if (!isAdmin_(userId)) {
    sendMessage(chatId, "❌ Эта команда доступна только администратору.");
    return;
  }

  if (args.length < 2) {
    sendMessage(chatId, 
      "📋 Использование:\n\n" +
      "/broadcast video <fileId> <текст> - рассылка видео\n" +
      "/broadcast text <текст> - текстовая рассылка\n" +
      "/broadcast ru <текст> - только русскоязычным\n" +
      "/broadcast tj <текст> - только таджикскоязычным\n\n" +
      "Пример:\n" +
      "/broadcast video 1ABC...xyz Всем новое видео!"
    );
    return;
  }

  const type = args[0].toLowerCase();
  const text = args.slice(1).join(" ");

  if (type === "video") {
    const parts = text.split(/\s+/);
    if (parts.length < 2) {
      sendMessage(chatId, "❌ Укажите fileId и текст. Пример: /broadcast video 1ABC...xyz Текст сообщения");
      return;
    }
    const fileId = parts[0];
    const caption = parts.slice(1).join(" ");
    
    sendMessage(chatId, "🔄 Запускаю рассылку видео... Это может занять несколько минут. Вы получите отчёт.");
    // Запускаем рассылку и сразу возвращаем управление
    try {
      broadcastVideo(fileId, caption, null);
    } catch (e) {
      sendMessage(chatId, "❌ Ошибка рассылки: " + e.message);
      Logger.log("broadcastVideo error: " + e);
    }
    
  } else if (type === "text") {
    sendMessage(chatId, "🔄 Запускаю текстовую рассылку... Это может занять несколько минут. Вы получите отчёт.");
    try {
      broadcastText(text, null);
    } catch (e) {
      sendMessage(chatId, "❌ Ошибка рассылки: " + e.message);
      Logger.log("broadcastText error: " + e);
    }
    
  } else if (type === "ru" || type === "tj") {
    sendMessage(chatId, `🔄 Запускаю рассылку ${type === "ru" ? "русскоязычным" : "таджикскоязычным"}...`);
    try {
      broadcastText(text, type);
    } catch (e) {
      sendMessage(chatId, "❌ Ошибка рассылки: " + e.message);
      Logger.log("broadcastText error: " + e);
    }
    
  } else {
    sendMessage(chatId, "❌ Неверный тип рассылки. Используйте: video, text, ru, или tj");
  }
}

/**********************
 *  ВЕБХУК
 *********************/
function setWebhook() {
  const scriptUrl = ScriptApp.getService().getUrl();
  const url = `https://api.telegram.org/bot${TOKEN}/setWebhook?url=${encodeURIComponent(scriptUrl)}`;
  const resp = UrlFetchApp.fetch(url);
  Logger.log(resp.getContentText());
}

/**********************
 *  УВЕДОМЛЕНИЕ АДМИНА
 *********************/
function notifyAdmins_(text) {
  const ids = getAdminIds_();
  ids.forEach(id => {
    try { sendMessage(id, text); } catch (_) { }
  });
}
