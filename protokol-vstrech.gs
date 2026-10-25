/**
 * Протокол встреч: запись → расшифровка (Yandex SpeechKit) → итоги и задачи (YandexGPT)
 * → документы Word в папке встречи на Google Диске → письмо участникам из Google Календаря.
 *
 * Как работает:
 * 1. Вы кладёте запись встречи (MP3, WAV или OGG) в папку «Встречи / Входящие» на Google Диске.
 * 2. Раз в 10 минут скрипт находит новый файл и отправляет его на расшифровку.
 * 3. Когда расшифровка готова, скрипт:
 *    — находит встречу в Google Календаре (по времени файла) и берёт оттуда название и участников;
 *    — создаёт папку «Встречи / ГГГГ-ММ-ДД ЧЧ:ММ — Название» и переносит туда запись;
 *    — сохраняет туда «Расшифровка.docx» и «Итоги и задачи.docx»;
 *    — отправляет письмо с итогами и задачами (сначала только вам — см. SEND_TO_PARTICIPANTS).
 *
 * Настройка (один раз):
 * — Настройки проекта (шестерёнка слева) → «Свойства скрипта» → добавить:
 *     YANDEX_API_KEY   — секретный ключ сервисного аккаунта Yandex Cloud
 *     YANDEX_FOLDER_ID — идентификатор каталога Yandex Cloud
 * — Выбрать функцию setup и нажать «Выполнить», затем разрешить доступ.
 *
 * Fireflies (бот сам записывает встречи):
 * — В свойства скрипта добавить FIREFLIES_API_KEY — ключ из Fireflies (Settings → Developer Settings).
 * — Выполнить testFireflies (проверка), затем setupFireflies (запуск каждые 15 минут).
 * — Новые встречи из Fireflies сами появятся в папке «Встречи» с Word-файлами, а итоги придут на почту.
 */

const CONFIG = {
  ROOT_FOLDER_NAME: 'Встречи',
  INBOX_FOLDER_NAME: 'Входящие',
  ERROR_FOLDER_NAME: 'Не обработано',
  SEND_TO_PARTICIPANTS: false,   // false — письмо приходит только вам (для проверки). Потом поставьте true.
  LOOKBACK_HOURS: 24,            // в каком окне до появления файла искать встречу в календаре
  MAX_FILE_MB: 25,               // больше — не влезает в запрос; сожмите в MP3 моно 32–64 кбит/с
  GPT_MODEL: 'yandexgpt/latest',
  CHECK_EVERY_MINUTES: 10,
  TIMEZONE: 'Europe/Moscow',
  // Fireflies
  SUMMARY_SOURCE: 'yandex',      // 'yandex' — итоги в нашем формате через YandexGPT; 'fireflies' — взять итоги Fireflies
  SAVE_AUDIO: true,              // сохранять аудио из Fireflies в папку встречи (если тариф Fireflies даёт ссылку)
  FIREFLIES_FIRST_DAYS: 1,       // при первом запуске забрать встречи за последние N дней
};

/* ===================== Запуск ===================== */

function setup() {
  apiKey_(); folderId_(); // проверяем, что свойства заданы
  const root = getOrCreateFolder_(DriveApp.getRootFolder(), CONFIG.ROOT_FOLDER_NAME);
  getOrCreateFolder_(root, CONFIG.INBOX_FOLDER_NAME);
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'processInbox')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processInbox').timeBased().everyMinutes(CONFIG.CHECK_EVERY_MINUTES).create();
  Logger.log('Готово. Кладите записи в папку «%s / %s». Проверка каждые %s минут.',
    CONFIG.ROOT_FOLDER_NAME, CONFIG.INBOX_FOLDER_NAME, CONFIG.CHECK_EVERY_MINUTES);
}

function processInbox() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    const root = getOrCreateFolder_(DriveApp.getRootFolder(), CONFIG.ROOT_FOLDER_NAME);
    const inbox = getOrCreateFolder_(root, CONFIG.INBOX_FOLDER_NAME);
    const files = inbox.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      try {
        handleFile_(file, root);
      } catch (e) {
        props_().deleteProperty('op_' + file.getId());
        file.moveTo(getOrCreateFolder_(root, CONFIG.ERROR_FOLDER_NAME));
        notifyOwner_('Встреча не обработана: ' + file.getName(),
          'Файл перенесён в папку «' + CONFIG.ERROR_FOLDER_NAME + '».\n\nПричина: ' + e.message +
          '\n\nИсправьте проблему и положите файл в «' + CONFIG.INBOX_FOLDER_NAME + '» ещё раз.');
      }
    }
  } finally {
    lock.releaseLock();
  }
}

/* ===================== Обработка файла ===================== */

function handleFile_(file, root) {
  const p = props_();
  const opKey = 'op_' + file.getId();
  let opId = p.getProperty(opKey);

  if (!opId) {
    const type = audioType_(file);
    if (!type) throw new Error('Формат не подходит. Нужен MP3, WAV или OGG. Запись в MP4 сначала сконвертируйте в MP3.');
    if (file.getSize() > CONFIG.MAX_FILE_MB * 1024 * 1024) {
      throw new Error('Файл больше ' + CONFIG.MAX_FILE_MB + ' МБ. Сожмите его в MP3 моно 32–64 кбит/с или разделите на части.');
    }
    opId = startRecognition_(file, type);
    p.setProperty(opKey, opId);
    return; // результат заберём на следующем запуске
  }

  const op = JSON.parse(callApi_('https://operation.api.cloud.yandex.net/operations/' + opId));
  if (!op.done) return;
  if (op.error) throw new Error('SpeechKit: ' + JSON.stringify(op.error));

  const transcript = getTranscript_(opId);
  p.deleteProperty(opKey);
  if (!transcript.trim()) throw new Error('Расшифровка пустая: в записи не распознана речь.');
  finishMeeting_(file, root, transcript);
}

function finishMeeting_(file, root, transcript) {
  const event = findEvent_(file);
  deliverMeeting_({
    root: root,
    title: event ? event.getTitle() : file.getName().replace(/\.[^.]+$/, ''),
    when: event ? event.getStartTime() : meetingTime_(file).time,
    transcript: transcript,
    guests: event ? event.getGuestList().map(g => g.getEmail()) : [],
    audioFile: file,
  });
}

/**
 * Общая выдача результата: папка встречи, Word-файлы, письмо.
 * opts: root, title, when, transcript, guests[], audioFile? , audioBlob?, summary?, sourceLink?
 */
function deliverMeeting_(opts) {
  const stamp = Utilities.formatDate(opts.when, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm');
  const title = opts.title || 'Встреча';
  const meetingFolder = opts.root.createFolder(stamp + ' — ' + title);
  if (opts.audioFile) opts.audioFile.moveTo(meetingFolder);
  if (opts.audioBlob) meetingFolder.createFile(opts.audioBlob);

  const summary = opts.summary || summarize_(opts.transcript, title);

  saveDocx_(meetingFolder, 'Расшифровка — ' + title, body => {
    body.appendParagraph('Расшифровка встречи').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph(title + ' · ' + stamp);
    if (opts.sourceLink) body.appendParagraph('Запись и расшифровка в Fireflies: ' + opts.sourceLink);
    opts.transcript.split('\n').forEach(line => body.appendParagraph(line));
  });
  const summaryFile = saveDocx_(meetingFolder, 'Итоги и задачи — ' + title, body => {
    body.appendParagraph('Итоги и задачи').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph(title + ' · ' + stamp);
    summary.split('\n').forEach(line => body.appendParagraph(line.replace(/^#+\s*/, '').replace(/\*\*/g, '')));
  });

  const owner = Session.getEffectiveUser().getEmail();
  const guests = (opts.guests || []).filter(e => e && !/fireflies|fred@/i.test(e));
  let recipients = [owner];
  if (CONFIG.SEND_TO_PARTICIPANTS) recipients = recipients.concat(guests);
  recipients = Array.from(new Set(recipients.filter(Boolean)));

  const subject = 'Итоги встречи: ' + title + ' (' + stamp + ')';
  const text = summary.replace(/\*\*/g, '') + '\n\nПапка встречи (доступна организатору): ' + meetingFolder.getUrl();
  GmailApp.sendEmail(recipients.join(','), subject, text, {
    htmlBody: toHtml_(summary) + '<p style="color:#666">Документ с итогами — во вложении.</p>',
    attachments: [summaryFile.getBlob()],
    name: 'Протокол встречи',
  });
  if (!CONFIG.SEND_TO_PARTICIPANTS && guests.length) {
    notifyOwner_('Проверьте итоги: ' + title,
      'Режим проверки: письмо с итогами пришло только вам. Участники: ' + guests.join(', ') +
      '.\nКогда убедитесь, что итоги точные, поставьте SEND_TO_PARTICIPANTS: true.');
  }
}

/* ===================== Fireflies ===================== */

/**
 * Забирает новые встречи из Fireflies (бот сам записал и расшифровал их),
 * делает итоги в нашем формате, Word-файлы в папке встречи и письмо участникам.
 */
function syncFireflies() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    const p = props_();
    const root = getOrCreateFolder_(DriveApp.getRootFolder(), CONFIG.ROOT_FOLDER_NAME);
    const since = p.getProperty('ff_since') ||
      new Date(Date.now() - CONFIG.FIREFLIES_FIRST_DAYS * 24 * 3600 * 1000).toISOString();
    const done = JSON.parse(p.getProperty('ff_done') || '[]');

    const data = firefliesQuery_(
      'query($from: DateTime, $limit: Int) { transcripts(fromDate: $from, limit: $limit, mine: true) {' +
      ' id title date duration transcript_url audio_url' +
      ' meeting_attendees { displayName email name }' +
      ' sentences { speaker_name text start_time }' +
      ' summary { overview action_items shorthand_bullet } } }',
      { from: new Date(new Date(since).getTime() - 24 * 3600 * 1000).toISOString(), limit: 20 }); // запас сутки: Fireflies обрабатывает записи с задержкой

    const list = (data.transcripts || []).sort((a, b) => a.date - b.date);
    let newest = since;
    list.forEach(t => {
      if (done.indexOf(t.id) >= 0) return;
      if (!t.sentences || !t.sentences.length) return; // ещё обрабатывается — заберём позже
      try {
        handleFireflies_(t, root);
      } catch (e) {
        notifyOwner_('Встреча из Fireflies не обработана: ' + t.title, 'Причина: ' + e.message);
      }
      done.push(t.id);
      const iso = new Date(Number(t.date)).toISOString();
      if (iso > newest) newest = iso;
    });
    p.setProperty('ff_done', JSON.stringify(done.slice(-200)));
    p.setProperty('ff_since', newest);
  } finally {
    lock.releaseLock();
  }
}

function handleFireflies_(t, root) {
  const transcript = t.sentences.map(s =>
    '[' + mmss_(Number(s.start_time || 0) * 1000) + '] ' + (s.speaker_name ? s.speaker_name + ': ' : '') + s.text
  ).join('\n');

  let summary = null;
  if (CONFIG.SUMMARY_SOURCE === 'fireflies' || !props_().getProperty('YANDEX_API_KEY')) {
    summary = firefliesSummary_(t);
  } else {
    try { summary = summarize_(transcript, t.title); }
    catch (e) { summary = firefliesSummary_(t); } // если YandexGPT недоступен — берём итоги Fireflies
  }

  let audioBlob = null;
  if (CONFIG.SAVE_AUDIO && t.audio_url) {
    try {
      const r = UrlFetchApp.fetch(t.audio_url, { muteHttpExceptions: true });
      if (r.getResponseCode() < 300) audioBlob = r.getBlob().setName('Запись — ' + t.title + '.mp3');
    } catch (e) { /* запись слишком большая или недоступна — оставим ссылку на Fireflies */ }
  }

  deliverMeeting_({
    root: root,
    title: t.title,
    when: new Date(Number(t.date)),
    transcript: transcript,
    guests: (t.meeting_attendees || []).map(a => a.email),
    audioBlob: audioBlob,
    summary: summary,
    sourceLink: t.transcript_url,
  });
}

function firefliesSummary_(t) {
  const s = t.summary || {};
  const out = ['ТЕМА', s.overview || '—', '', 'ГЛАВНЫЕ ИТОГИ'];
  (s.shorthand_bullet || '').split('\n').filter(Boolean).forEach(x => out.push('- ' + x.replace(/^[-•*]\s*/, '')));
  out.push('', 'ЗАДАЧИ');
  (s.action_items || '').split('\n').filter(Boolean).forEach(x => out.push('- ' + x.replace(/^[-•*]\s*/, '')));
  return out.join('\n');
}

function firefliesQuery_(query, variables) {
  const key = props_().getProperty('FIREFLIES_API_KEY');
  if (!key) throw new Error('Не задано свойство скрипта FIREFLIES_API_KEY (Настройки проекта → Свойства скрипта).');
  const r = UrlFetchApp.fetch('https://api.fireflies.ai/graphql', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + key.trim() },
    payload: JSON.stringify({ query: query, variables: variables || {} }),
    muteHttpExceptions: true,
  });
  const res = JSON.parse(r.getContentText());
  if (r.getResponseCode() >= 300 || res.errors) {
    throw new Error('Fireflies ответил ' + r.getResponseCode() + ': ' + JSON.stringify(res.errors || res).slice(0, 400));
  }
  return res.data;
}

/* Проверка подключения к Fireflies: выберите testFireflies и нажмите «Выполнить». */
function testFireflies() {
  const d = firefliesQuery_('query { transcripts(limit: 5, mine: true) { title date } }');
  Logger.log('Последние встречи в Fireflies: ' +
    (d.transcripts || []).map(t => t.title + ' (' + Utilities.formatDate(new Date(Number(t.date)), CONFIG.TIMEZONE, 'dd.MM HH:mm') + ')').join('; '));
}

/* Запуск автоматической выгрузки из Fireflies каждые 15 минут. */
function setupFireflies() {
  firefliesQuery_('query { user { email } }');
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'syncFireflies')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('syncFireflies').timeBased().everyMinutes(15).create();
  getOrCreateFolder_(DriveApp.getRootFolder(), CONFIG.ROOT_FOLDER_NAME);
  Logger.log('Готово: новые встречи из Fireflies будут забираться каждые 15 минут.');
}

/* ===================== Yandex SpeechKit ===================== */

function startRecognition_(file, type) {
  const body = {
    content: Utilities.base64Encode(file.getBlob().getBytes()),
    recognitionModel: {
      model: 'general',
      audioFormat: { containerAudio: { containerAudioType: type } },
      textNormalization: { textNormalization: 'TEXT_NORMALIZATION_ENABLED', profanityFilter: false, literatureText: true },
      languageRestriction: { restrictionType: 'WHITELIST', languageCode: ['ru-RU'] },
    },
    speakerLabeling: { speakerLabeling: 'SPEAKER_LABELING_ENABLED' },
  };
  const url = 'https://stt.api.cloud.yandex.net/stt/v3/recognizeFileAsync';
  try {
    return JSON.parse(callApi_(url, body)).id;
  } catch (e) {
    // если разметка спикеров недоступна для этого файла — пробуем без неё
    delete body.speakerLabeling;
    return JSON.parse(callApi_(url, body)).id;
  }
}

function getTranscript_(opId) {
  const raw = callApi_('https://stt.api.cloud.yandex.net/stt/v3/getRecognition?operationId=' + opId);
  const refined = [], finals = [];
  raw.split('\n').forEach(line => {
    line = line.trim();
    if (!line) return;
    let o;
    try { o = JSON.parse(line); } catch (e) { return; }
    const r = o.result || o;
    const tag = r.channelTag || '0';
    const ref = r.finalRefinement && r.finalRefinement.normalizedText;
    if (ref && ref.alternatives && ref.alternatives[0] && ref.alternatives[0].text) {
      refined.push({ tag: tag, text: ref.alternatives[0].text, ms: Number(ref.alternatives[0].startTimeMs || 0) });
    }
    const fin = r.final;
    if (fin && fin.alternatives && fin.alternatives[0] && fin.alternatives[0].text) {
      finals.push({ tag: tag, text: fin.alternatives[0].text, ms: Number(fin.alternatives[0].startTimeMs || 0) });
    }
  });
  const parts = (refined.length ? refined : finals).sort((a, b) => a.ms - b.ms);
  const tags = Array.from(new Set(parts.map(x => x.tag)));
  const many = tags.length > 1;
  const lines = [];
  let last = null;
  parts.forEach(x => {
    if (many && x.tag === last && lines.length) {
      lines[lines.length - 1] += ' ' + x.text;
    } else {
      const who = many ? 'Спикер ' + (tags.indexOf(x.tag) + 1) + ': ' : '';
      lines.push('[' + mmss_(x.ms) + '] ' + who + x.text);
    }
    last = x.tag;
  });
  return lines.join('\n');
}

/* ===================== YandexGPT ===================== */

const SUMMARY_PROMPT =
  'Ты — опытный бизнес-ассистент руководителя. По расшифровке рабочей встречи составь протокол на русском языке строго в формате:\n\n' +
  'ТЕМА\nОдно предложение — о чём была встреча.\n\n' +
  'ГЛАВНЫЕ ИТОГИ\n- 3–7 пунктов: что решили и о чём договорились.\n\n' +
  'ЗАДАЧИ\n- Ответственный — что сделать — срок\n\n' +
  'ОТКРЫТЫЕ ВОПРОСЫ\n- что осталось нерешённым (если нет — «нет»).\n\n' +
  'Правила: пиши кратко и по делу; ничего не придумывай; если ответственный не назван — пиши «ответственный не определён», ' +
  'если срок не назван — «срок не обсуждался»; имена бери из расшифровки, а не «Спикер 1».';

function summarize_(transcript, title) {
  const LIMIT = 50000;
  if (transcript.length <= LIMIT) {
    return callGpt_(SUMMARY_PROMPT, 'Название встречи: ' + title + '\n\nРасшифровка:\n' + transcript);
  }
  // длинная встреча: сначала конспектируем частями, потом собираем протокол
  const chunks = [];
  for (let i = 0; i < transcript.length; i += LIMIT) chunks.push(transcript.slice(i, i + LIMIT));
  const notes = chunks.map((c, i) => callGpt_(
    'Кратко законспектируй фрагмент рабочей встречи: решения, задачи с ответственными и сроками, открытые вопросы. Ничего не придумывай.',
    'Фрагмент ' + (i + 1) + ' из ' + chunks.length + ':\n' + c, 1500));
  return callGpt_(SUMMARY_PROMPT, 'Название встречи: ' + title + '\n\nКонспект встречи по частям:\n' + notes.join('\n\n'));
}

function callGpt_(system, user, maxTokens) {
  const body = {
    modelUri: 'gpt://' + folderId_() + '/' + CONFIG.GPT_MODEL,
    completionOptions: { stream: false, temperature: 0.2, maxTokens: String(maxTokens || 2000) },
    messages: [{ role: 'system', text: system }, { role: 'user', text: user }],
  };
  const res = JSON.parse(callApi_('https://llm.api.cloud.yandex.net/foundationModels/v1/completion', body));
  return res.result.alternatives[0].message.text.trim();
}

/* ===================== Вспомогательное ===================== */

function callApi_(url, body) {
  const opt = {
    method: body ? 'post' : 'get',
    muteHttpExceptions: true,
    headers: { Authorization: 'Api-Key ' + apiKey_(), 'x-folder-id': folderId_() },
  };
  if (body) { opt.contentType = 'application/json'; opt.payload = JSON.stringify(body); }
  const r = UrlFetchApp.fetch(url, opt);
  const code = r.getResponseCode();
  const text = r.getContentText();
  if (code >= 300) throw new Error('Yandex Cloud ответил ' + code + ': ' + text.slice(0, 400));
  return text;
}

function meetingTime_(file) {
  // время встречи берём из названия файла («Запись встречи 18.10.2026 22_11_23», «2026-10-18 22-11»),
  // а если его там нет — из даты загрузки файла на Диск
  const n = file.getName();
  let m = n.match(/(\d{2})\.(\d{2})\.(\d{4})[ _T-]+(\d{2})[_:.\-](\d{2})/);
  if (m) return { time: Utilities.parseDate(m[3] + '-' + m[2] + '-' + m[1] + ' ' + m[4] + ':' + m[5], CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm'), fromName: true };
  m = n.match(/(\d{4})-(\d{2})-(\d{2})[ _T]+(\d{2})[_:.\-](\d{2})/);
  if (m) return { time: Utilities.parseDate(m[1] + '-' + m[2] + '-' + m[3] + ' ' + m[4] + ':' + m[5], CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm'), fromName: true };
  return { time: file.getDateCreated(), fromName: false };
}

function findEvent_(file) {
  const mt = meetingTime_(file);
  const t = mt.time;
  const before = (mt.fromName ? 6 : CONFIG.LOOKBACK_HOURS) * 3600 * 1000;
  const after = (mt.fromName ? 2 : 0) * 3600 * 1000;
  const events = CalendarApp.getDefaultCalendar()
    .getEvents(new Date(t.getTime() - before), new Date(t.getTime() + after))
    .filter(e => !e.isAllDayEvent());
  const withGuests = events.filter(e => e.getGuestList().length > 0);
  const list = (withGuests.length ? withGuests : events)
    .sort((a, b) => Math.abs(a.getStartTime() - t) - Math.abs(b.getStartTime() - t));
  return list[0] || null;
}

function saveDocx_(folder, name, build) {
  const doc = DocumentApp.create(name);
  build(doc.getBody());
  doc.saveAndClose();
  const mime = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  const blob = UrlFetchApp.fetch(
    'https://www.googleapis.com/drive/v3/files/' + doc.getId() + '/export?mimeType=' + encodeURIComponent(mime),
    { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }
  ).getBlob().setName(name + '.docx');
  const file = folder.createFile(blob);
  DriveApp.getFileById(doc.getId()).setTrashed(true);
  return file;
}

function audioType_(file) {
  const mt = file.getMimeType();
  const n = file.getName().toLowerCase();
  if (mt === 'audio/mpeg' || mt === 'audio/mp3' || n.endsWith('.mp3')) return 'MP3';
  if (mt === 'audio/wav' || mt === 'audio/x-wav' || mt === 'audio/wave' || n.endsWith('.wav')) return 'WAV';
  if (mt === 'audio/ogg' || mt === 'audio/opus' || n.endsWith('.ogg') || n.endsWith('.opus')) return 'OGG_OPUS';
  return null;
}

function toHtml_(text) {
  const esc = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  let html = '', inList = false;
  text.split('\n').forEach(raw => {
    const line = raw.replace(/^#+\s*/, '').replace(/\*\*/g, '').trim();
    if (/^[-•*]\s+/.test(line)) {
      if (!inList) { html += '<ul>'; inList = true; }
      html += '<li>' + esc(line.replace(/^[-•*]\s+/, '')) + '</li>';
      return;
    }
    if (inList) { html += '</ul>'; inList = false; }
    if (!line) return;
    html += /^[A-ZА-ЯЁ\s]+:?$/.test(line) ? '<h3 style="margin:16px 0 6px">' + esc(line) + '</h3>' : '<p>' + esc(line) + '</p>';
  });
  if (inList) html += '</ul>';
  return '<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.5">' + html + '</div>';
}

function mmss_(ms) {
  const s = Math.floor(ms / 1000);
  const h = Math.floor(s / 3600), m = Math.floor((s % 3600) / 60), sec = s % 60;
  const pad = x => (x < 10 ? '0' : '') + x;
  return (h ? h + ':' + pad(m) : pad(m)) + ':' + pad(sec);
}

function getOrCreateFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function notifyOwner_(subject, body) {
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(), subject, body);
}

function props_() { return PropertiesService.getScriptProperties(); }

function apiKey_() {
  const k = props_().getProperty('YANDEX_API_KEY');
  if (!k) throw new Error('Не задано свойство скрипта YANDEX_API_KEY (Настройки проекта → Свойства скрипта).');
  return k.trim();
}

function folderId_() {
  const f = props_().getProperty('YANDEX_FOLDER_ID');
  if (!f) throw new Error('Не задано свойство скрипта YANDEX_FOLDER_ID (Настройки проекта → Свойства скрипта).');
  return f.trim();
}

/* Проверка подключения к YandexGPT: выберите testYandex и нажмите «Выполнить» — в журнале появится ответ. */
function testYandex() {
  Logger.log(callGpt_('Отвечай одним предложением.', 'Скажи «Подключение работает» и назови сегодняшний день недели, если знаешь.', 50));
}
