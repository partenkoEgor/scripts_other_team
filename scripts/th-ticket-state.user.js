// ==UserScript==
// @name         TH Management — Состояние тикетов
// @namespace    th-management-ticket-state
// @version      1.2
// @description  Кнопка «Состояние тикетов» на странице списка платёжных запросов: по списку Ticket ID (или по тикетам на экране) спрашивает сайт и показывает текущее состояние каждого тикета — External Status, Transaction ID, Inside comment, User ID, ГЕО, реферала, агента и субагента — со сводкой, таблицей для копирования и файлом Excel (общий лист и по листу на каждый External Status). Только читает: окна Edit не открываются, в тикетах ничего не меняется.
// @match        https://th-managment.com/en/admin/backoffice/paymentsupport*
// @match        https://managment.io/en/admin/backoffice/paymentsupport*
// @match        https://my-managment.com/en/admin/backoffice/paymentsupport*
// @grant        none
// @updateURL    https://raw.githubusercontent.com/partenkoEgor/scripts_other_team/main/scripts/th-ticket-state.user.js
// @downloadURL  https://raw.githubusercontent.com/partenkoEgor/scripts_other_team/main/scripts/th-ticket-state.user.js
// ==/UserScript==

// Отдельный скрипт из кнопки «Состояние тикетов» беты Bulk Approve (0.31,
// репозиторий partenkoEgor/Monitoring) — для команд, которым остальной Bulk
// Approve не нужен. Функции перенесены из беты как есть: запрос списка,
// разбор ответа, сборка Excel и окна. Поменялись только префиксы (id и
// классы на странице, консоль) и перехват запросов — здесь он только
// запоминает запрос списка. Исправление в одной копии стоит перенести и в
// другую.
//
// Скрипт уживается с Bulk Approve на одной странице: у него свои id и
// классы, своя кнопка (слева внизу), а свои запросы он шлёт мимо журнала
// запросов Bulk Approve.

(function () {
  'use strict';

  // ------------------------------------------------------------------
  // НАСТРОЙКИ
  // ------------------------------------------------------------------
  const CONFIG = {
    // Сколько Ticket ID спрашивать у сайта одним запросом (страница выдачи)
    precheckChunkSize: 100,
    // Пауза между запросами к сайту (мс) — чтобы не нагружать его
    precheckBetweenRequestsDelay: 300,
    // Сколько раз повторить запрос, если сайт перегружен (429, 5xx, 529)
    precheckRetries: 3,
    // Пауза перед повтором (мс), растёт с каждой попыткой
    precheckRetryDelay: 1000,
    // Страховка от бесконечной выдачи: больше страниц на пачку не листаем
    precheckMaxPages: 20,
    // Сколько ждать, пока страница сама отправит запрос списка после Apply (мс)
    precheckTemplateTimeout: 10000,
  };

  // ------------------------------------------------------------------
  // ЗАПРОС СПИСКА ТИКЕТОВ.
  //
  // Таблицу страница грузит одним POST-запросом: в теле все фильтры, номера
  // Ticket ID — строкой через перенос (idTicket), страница (page) и размер
  // (perPage). Ответ — JSON со строками таблицы.
  //
  // Сам запрос скрипт не собирает: он запоминает тот, что отправила
  // страница, — со всеми заголовками и защитным токеном — и повторяет его,
  // меняя только номера и страницу. Через этот запрос скрипт только ЧИТАЕТ.
  // ------------------------------------------------------------------
  let listRequestTemplate = null; // { url, body, headers, capturedAt }
  let pageFetch = null;

  // Чем слать свои запросы: настоящим fetch страницы. Если на странице есть
  // Bulk Approve, он хранит его у себя — берём оттуда, чтобы свои запросы не
  // попадали в его журнал (по журналу он ждёт, когда страница догрузится).
  function nativeFetch() {
    return typeof window.__bulkApproveNativeFetch === 'function'
      ? window.__bulkApproveNativeFetch
      : pageFetch;
  }

  // Перехват только запоминает запрос списка: адрес, тело и заголовки.
  // Сами запросы страницы идут как шли.
  function installListCapture() {
    if (window.__thTicketStateCaptureInstalled) return;
    window.__thTicketStateCaptureInstalled = true;

    const origOpen = XMLHttpRequest.prototype.open;
    const origSend = XMLHttpRequest.prototype.send;
    const origSetHeader = XMLHttpRequest.prototype.setRequestHeader;

    XMLHttpRequest.prototype.open = function (method, url, ...rest) {
      try {
        this.__thTicketStateUrl = url;
        this.__thTicketStateMethod = method;
      } catch (e) {
        /* перехват не имеет права мешать запросу */
      }
      return origOpen.call(this, method, url, ...rest);
    };
    if (typeof origSetHeader === 'function') {
      XMLHttpRequest.prototype.setRequestHeader = function (name, value) {
        try {
          if (!this.__thTicketStateHeaders) this.__thTicketStateHeaders = {};
          this.__thTicketStateHeaders[name] = value;
        } catch (e) {
          /* см. выше */
        }
        return origSetHeader.call(this, name, value);
      };
    }
    XMLHttpRequest.prototype.send = function (...args) {
      try {
        noteListRequest(this.__thTicketStateMethod, this.__thTicketStateUrl, args[0], this.__thTicketStateHeaders);
      } catch (e) {
        /* см. выше */
      }
      return origSend.apply(this, args);
    };

    const origFetch = window.fetch;
    if (typeof origFetch === 'function') {
      pageFetch = origFetch;
      window.fetch = function (input, ...rest) {
        try {
          const init = rest[0] || {};
          noteListRequest(
            init.method || (input && input.method) || 'GET',
            input && input.url ? input.url : input,
            init.body,
            headersToObject(init.headers)
          );
        } catch (e) {
          /* см. выше */
        }
        return origFetch.call(this, input, ...rest);
      };
    }
  }

  // ------------------------------------------------------------------
  // Вспомогательные функции
  // ------------------------------------------------------------------
  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  function fireClick(el) {
    if (!el) return;
    el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
    el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true }));
    el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
  }

  function getTicketRows() {
    return Array.from(document.querySelectorAll('.table-wrapper tbody tr[data-table-row]'));
  }

  // Номера колонок по названию из шапки таблицы — чтобы читать Ticket ID с
  // экрана, даже если колонки переставлены
  let columnIndexMap = null; // { 'ticket id': 9, ... } (1-based, как в nth-child)

  function refreshColumnIndexMap() {
    const headerCells = document.querySelectorAll('.table-wrapper thead th');
    if (headerCells.length === 0) {
      console.warn('[TicketState] Не найдены заголовки таблицы — использую резервные номера колонок.');
      columnIndexMap = null;
      return;
    }

    const map = {};
    headerCells.forEach((th, idx) => {
      const label = th.textContent.trim().toLowerCase();
      if (label) map[label] = idx + 1; // nth-child считается с 1
    });
    columnIndexMap = map;

    const required = ['ticket id'];
    const missing = required.filter((name) => !map[name]);
    if (missing.length > 0) {
      console.warn(
        `[TicketState] Не найдены колонки по названию: ${missing.join(', ')} — для них будут использованы резервные номера.`
      );
    }
  }

  // Возвращает номер колонки по названию, либо резервный номер, если название не нашлось
  function getColumnIndex(name, fallbackIndex) {
    if (!columnIndexMap) return fallbackIndex;
    return columnIndexMap[name.toLowerCase()] || fallbackIndex;
  }

  function getTicketIdFromRow(row) {
    const idx = getColumnIndex('ticket id', 9);
    const cell = row.querySelector(`td:nth-child(${idx})`);
    return cell ? cell.textContent.trim() : '(unknown)';
  }

  function getFilterApplyButton() {
    return (
      document.querySelector('#filter_form .btn-block button[type="submit"]') ||
      document.querySelector('#filter_form button[type="submit"]')
    );
  }

  function chunkArray(items, size) {
    const out = [];
    for (let i = 0; i < items.length; i += size) out.push(items.slice(i, i + size));
    return out;
  }

  // 1 тикет, 2 тикета, 5 тикетов
  function pluralRu(n, one, few, many) {
    const n100 = Math.abs(n) % 100;
    const n10 = n100 % 10;
    if (n100 > 10 && n100 < 20) return many;
    if (n10 === 1) return one;
    if (n10 >= 2 && n10 <= 4) return few;
    return many;
  }

  // «Статус — сколько» по убыванию: для сводки, а не для поштучного списка
  function countByStatus(items) {
    const counts = new Map();
    items.forEach((w) => {
      const name = w.externalStatus || '(пусто)';
      counts.set(name, (counts.get(name) || 0) + 1);
    });
    return [...counts.entries()].sort((a, b) => b[1] - a[1]);
  }

  function headersToObject(h) {
    if (!h) return {};
    if (Array.isArray(h)) return Object.fromEntries(h);
    if (typeof h.forEach === 'function') {
      const out = {};
      h.forEach((value, key) => {
        out[key] = value;
      });
      return out;
    }
    return typeof h === 'object' ? { ...h } : {};
  }

  // Узнаёт запрос списка по телу, а не по адресу: адрес у него пустой
  // (страница шлёт его «на себя»).
  function noteListRequest(method, url, body, headers) {
    if (String(method || '').toUpperCase() !== 'POST' || typeof body !== 'string') return;
    if (body.indexOf('idTicket') === -1) return;
    let parsed;
    try {
      parsed = JSON.parse(body);
    } catch (e) {
      return;
    }
    if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) return;
    if (!('idTicket' in parsed) || !('page' in parsed) || !('perPage' in parsed)) return;
    let resolved;
    try {
      resolved = new URL(url == null ? '' : String(url), window.location.href);
    } catch (e) {
      return;
    }
    resolved.hash = '';
    listRequestTemplate = {
      url: resolved.href,
      body: parsed,
      headers: { ...(headers || {}) },
      capturedAt: Date.now(),
    };
  }

  function siteQueryAvailable() {
    return typeof nativeFetch() === 'function';
  }

  function listRequestHeaders() {
    const h = { ...listRequestTemplate.headers };
    const has = (name) => Object.keys(h).some((k) => k.toLowerCase() === name.toLowerCase());
    if (!has('Content-Type')) h['Content-Type'] = 'application/json';
    if (!has('Accept')) h.Accept = 'application/json, text/plain, */*';
    if (!has('X-Requested-With')) h['X-Requested-With'] = 'XMLHttpRequest';
    if (!has('X-CSRF-TOKEN')) {
      const meta = document.querySelector('meta[name="csrf-token"]');
      if (meta && meta.getAttribute('content')) h['X-CSRF-TOKEN'] = meta.getAttribute('content');
    }
    return h;
  }

  // Статусы, при которых сайт просто перегружен и запрос стоит повторить
  const RETRYABLE_HTTP = [429, 500, 502, 503, 504, 520, 521, 522, 523, 524, 529];

  // Колонки ГЕО, Реферал и Агент — поля nameCountry, nameRef и nameAgent.
  // undefined — поля в ответе нет вовсе (колонка выключена в настройках
  // таблицы): это не «пусто», а «нет данных».
  function optionalText(r, key) {
    if (!(key in r)) return undefined;
    return r[key] == null ? '' : String(r[key]).trim();
  }

  // Одна страница выдачи для набора Ticket ID. Возвращает:
  //   { kind: 'ok', rows, hasMore } — данные;
  //   { kind: 'expired', detail }   — вместо данных страница входа;
  //   { kind: 'retry', detail }     — сайт перегружен или ещё считает;
  //   { kind: 'error', detail }     — что-то непонятное, дальше не идём.
  async function fetchListPage(ids, page) {
    const body = {
      ...listRequestTemplate.body,
      idTicket: ids.join('\n'),
      page,
      perPage: CONFIG.precheckChunkSize,
    };
    let resp;
    try {
      resp = await nativeFetch().call(window, listRequestTemplate.url, {
        method: 'POST',
        credentials: 'same-origin',
        cache: 'no-store',
        headers: listRequestHeaders(),
        body: JSON.stringify(body),
      });
    } catch (e) {
      return { kind: 'retry', detail: `запрос не дошёл до сайта (${(e && e.message) || e})` };
    }
    let text = '';
    try {
      text = await resp.text();
    } catch (e) {
      return { kind: 'retry', detail: 'ответ сайта оборвался' };
    }
    if (resp.status === 401 || resp.status === 419) {
      return { kind: 'expired', detail: `сайт ответил ${resp.status}` };
    }
    if (!resp.ok) {
      return RETRYABLE_HTTP.includes(resp.status)
        ? { kind: 'retry', detail: `сайт ответил ${resp.status}` }
        : { kind: 'error', detail: `сайт ответил ${resp.status}` };
    }
    // Данные приходят JSON'ом. HTML с кодом 200 — это страница входа, на
    // которую сайт перенаправляет, когда сессия кончилась.
    if (/^\s*</.test(text)) {
      return { kind: 'expired', detail: 'вместо данных пришла страница входа' };
    }
    let json;
    try {
      json = JSON.parse(text);
    } catch (e) {
      return { kind: 'error', detail: 'ответ сайта не читается' };
    }
    if (json && json.is_polling) {
      return { kind: 'retry', detail: 'сайт ещё собирает выдачу' };
    }
    if (!json || json.success === false || !Array.isArray(json.data)) {
      return { kind: 'error', detail: 'сайт ответил не в том виде, который знает скрипт' };
    }
    const rows = json.data
      .flat()
      .filter((r) => r && typeof r === 'object')
      .map((r) => ({
        ticketId: String(r.idTicket == null ? '' : r.idTicket).trim(),
        externalStatus: String(r.nameExternalStatus == null ? '' : r.nameExternalStatus).trim(),
        transactionId: r.idTransaction == null ? '' : String(r.idTransaction).trim(),
        // Колонка Agent's files. undefined — поля в ответе нет вовсе: это
        // не «пусто», а «не знаю».
        agentFiles: 'documentsExternal' in r ? (r.documentsExternal == null ? '' : String(r.documentsExternal)) : undefined,
        userId: r.idUser == null ? '' : String(r.idUser).trim(),
        subagent: String(r.nameSubagent == null ? '' : r.nameSubagent).trim(),
        geo: optionalText(r, 'nameCountry'),
        referral: optionalText(r, 'nameRef'),
        agent: optionalText(r, 'nameAgent'),
        // Inside comment
        insideComment: String(r.commentSupport == null ? '' : r.commentSupport).trim(),
      }))
      .filter((r) => r.ticketId !== '');
    return { kind: 'ok', rows, hasMore: !!(json.otherData && json.otherData.hasMore) };
  }

  // Строки сайта для списка Ticket ID — пачками, со страницами и повторами.
  // { ok: true, rows: Map ticketId → строка } или { ok: false, expired, detail }.
  async function queryTicketsOnSite(ids, { onProgress, isCancelled } = {}) {
    const rows = new Map();
    const chunks = chunkArray(ids, CONFIG.precheckChunkSize);
    let done = 0;
    for (let c = 0; c < chunks.length; c++) {
      for (let page = 1; ; page++) {
        if (page > CONFIG.precheckMaxPages) {
          return { ok: false, detail: 'сайт отдаёт слишком много страниц выдачи' };
        }
        let res;
        for (let attempt = 0; ; attempt++) {
          if (isCancelled && isCancelled()) return { ok: false, cancelled: true };
          res = await fetchListPage(chunks[c], page);
          if (res.kind !== 'retry' || attempt >= CONFIG.precheckRetries) break;
          await sleep(CONFIG.precheckRetryDelay * (attempt + 1));
        }
        if (res.kind === 'expired') return { ok: false, expired: true, detail: res.detail };
        if (res.kind !== 'ok') return { ok: false, detail: res.detail };
        res.rows.forEach((r) => rows.set(r.ticketId, r));
        if (!res.hasMore) break;
        await sleep(CONFIG.precheckBetweenRequestsDelay);
      }
      done += chunks[c].length;
      if (onProgress) onProgress(done, ids.length);
      if (c < chunks.length - 1) await sleep(CONFIG.precheckBetweenRequestsDelay);
    }
    return { ok: true, rows };
  }

  // Своего запроса у скрипта нет, пока страница не отправила свой. Таблица
  // при открытии страницы грузится раньше, чем запускается скрипт, поэтому
  // первый раз скрипт сам жмёт Apply у фильтров — ровно то, что сделал бы
  // человек. Фильтры страницы при этом не меняются.
  async function ensureListRequestTemplate(isCancelled) {
    if (listRequestTemplate) return true;
    if (!siteQueryAvailable()) return false;
    const btn = getFilterApplyButton();
    if (!btn) return false;
    fireClick(btn);
    const end = Date.now() + CONFIG.precheckTemplateTimeout;
    while (!listRequestTemplate && Date.now() < end) {
      if (isCancelled && isCancelled()) return false;
      await sleep(100);
    }
    return !!listRequestTemplate;
  }

  // Разбирает список Ticket ID — по номеру в строке. Номера — ключи pairs
  // (значения пустые: так устроен разбор в Bulk Approve, откуда он взят).
  //
  // Строго один столбец: так решено сознательно. Список сюда попадает из
  // выгрузки, и лишняя колонка обычно значит, что скопировали не тот кусок
  // таблицы, — молча взять из неё первое число было бы худшим из исходов.
  function parseTicketIdList(text) {
    const pairs = new Map();
    const badLines = [];
    let duplicates = 0;

    String(text == null ? '' : text)
      .split(/\r?\n/)
      .forEach((raw, idx) => {
        const lineNo = idx + 1;
        const line = raw.trim();
        if (line === '') return;

        if (/[\t;,]/.test(line) || /\s/.test(line)) {
          badLines.push({
            line: lineNo,
            text: line,
            why: 'в строке больше одного значения — нужен только Ticket ID',
          });
          return;
        }
        if (!/^\d+$/.test(line)) {
          badLines.push({ line: lineNo, text: line, why: 'не похоже на Ticket ID' });
          return;
        }

        if (pairs.has(line)) {
          duplicates++;
          return;
        }
        // null вместо номера транзакции: в этом режиме его просто нет
        pairs.set(line, null);
      });

    return { pairs, badLines, duplicateTickets: [], duplicateTransactions: [], duplicates };
  }

  // ------------------------------------------------------------------
  // ФАЙЛ EXCEL С ДУБЛЯМИ.
  //
  // Дубли разбирают руками, и удобнее всего — в таблице: на каждый дубль два
  // Ticket ID друг под другом (наш тикет, затем обращение, у которого эта
  // транзакция уже есть), а справа — их общий Transaction ID в объединённой
  // ячейке.
  //
  // Файл .xlsx собирается здесь же, без сторонних библиотек: это zip из
  // нескольких XML. Тянуть библиотеку с CDN ради одной таблицы — лишняя
  // зависимость и лишний чужой код на странице с платёжными данными.
  // ------------------------------------------------------------------
  const CRC32_TABLE = (() => {
    const table = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      table[n] = c >>> 0;
    }
    return table;
  })();

  function crc32(bytes) {
    let c = 0xffffffff;
    for (let i = 0; i < bytes.length; i++) c = CRC32_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  }

  // Zip без сжатия (method 0 — «stored»). Excel такие открывает без вопросов,
  // а писать deflate руками ради файла в десятки килобайт незачем.
  function buildZip(files) {
    const enc = new TextEncoder();
    const now = new Date();
    const dosTime = (now.getHours() << 11) | (now.getMinutes() << 5) | Math.floor(now.getSeconds() / 2);
    const dosDate = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
    const chunks = [];
    const central = [];
    let offset = 0;

    files.forEach(({ name, text }) => {
      const nameBytes = enc.encode(name);
      const data = enc.encode(text);
      const crc = crc32(data);

      const local = new DataView(new ArrayBuffer(30));
      local.setUint32(0, 0x04034b50, true);
      local.setUint16(4, 20, true);
      local.setUint16(6, 0x0800, true); // имена в UTF-8
      local.setUint16(8, 0, true);
      local.setUint16(10, dosTime, true);
      local.setUint16(12, dosDate, true);
      local.setUint32(14, crc, true);
      local.setUint32(18, data.length, true);
      local.setUint32(22, data.length, true);
      local.setUint16(26, nameBytes.length, true);
      local.setUint16(28, 0, true);
      chunks.push(new Uint8Array(local.buffer), nameBytes, data);

      const entry = new DataView(new ArrayBuffer(46));
      entry.setUint32(0, 0x02014b50, true);
      entry.setUint16(4, 20, true);
      entry.setUint16(6, 20, true);
      entry.setUint16(8, 0x0800, true);
      entry.setUint16(10, 0, true);
      entry.setUint16(12, dosTime, true);
      entry.setUint16(14, dosDate, true);
      entry.setUint32(16, crc, true);
      entry.setUint32(20, data.length, true);
      entry.setUint32(24, data.length, true);
      entry.setUint16(28, nameBytes.length, true);
      entry.setUint32(42, offset, true);
      central.push(new Uint8Array(entry.buffer), nameBytes);

      offset += 30 + nameBytes.length + data.length;
    });

    const centralSize = central.reduce((sum, part) => sum + part.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true);
    end.setUint16(8, files.length, true);
    end.setUint16(10, files.length, true);
    end.setUint32(12, centralSize, true);
    end.setUint32(16, offset, true);

    return new Blob([...chunks, ...central, new Uint8Array(end.buffer)], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
  }

  function xmlEscape(value) {
    return String(value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  // Номера, состоящие только из цифр, пишем числом с форматом «0»: так Excel
  // не ставит зелёные треугольники «число как текст» и не превращает длинный
  // номер в 2,39E+10. Дольше 15 цифр Excel число не удержит — такое и всё
  // нецифровое пишем текстом. Номер с ведущим нулём (кошельки вида
  // 01080969968) тоже текстом: числом он потерял бы этот ноль.
  function xlsxCell(ref, value, style) {
    // Формула с уже посчитанным значением: { f: 'SUM(B2:B3)', v: 5 } — Excel
    // покажет число сразу и пересчитает, если строки на листе поправят
    if (value && typeof value === 'object' && 'f' in value) {
      return `<c r="${ref}" s="${style.number}"><f>${xmlEscape(value.f)}</f><v>${Number(value.v) || 0}</v></c>`;
    }
    const v = value == null ? '' : String(value);
    if (v === '') return `<c r="${ref}" s="${style.text}"/>`;
    if (/^(0|[1-9]\d{0,14})$/.test(v)) return `<c r="${ref}" s="${style.number}"><v>${v}</v></c>`;
    // Текст из тикета (комментарии) может нести служебные символы, которых
    // XML не допускает: с ними Excel не откроет файл вовсе. Выкидываем их и
    // режем по пределу ячейки Excel (32 767 символов).
    const clean = v.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, '').slice(0, 32767);
    // Пробелы по краям и переносы строк Excel сохраняет только с этой пометкой
    const keep = /^\s|\s$|\n/.test(clean) ? ' xml:space="preserve"' : '';
    return `<c r="${ref}" s="${style.text}" t="inlineStr"><is><t${keep}>${xmlEscape(clean)}</t></is></c>`;
  }

  // Буква колонки Excel по номеру с нуля: 0 → A, 25 → Z, 26 → AA
  function xlsxColumn(n) {
    let out = '';
    for (let i = n + 1; i > 0; i = Math.floor((i - 1) / 26)) {
      out = String.fromCharCode(65 + ((i - 1) % 26)) + out;
    }
    return out;
  }

  // Имя листа по правилам Excel: до 31 символа, без []:*?/\ и уникальное
  function xlsxSheetNames(names) {
    const used = new Set();
    return names.map((raw) => {
      const base = (String(raw || '').replace(/[\[\]:*?\/\\]/g, ' ').replace(/\s+/g, ' ').trim() || 'Без страны')
        .slice(0, 31)
        .replace(/^'+|'+$/g, '') || 'Лист';
      let name = base;
      for (let k = 2; used.has(name.toLowerCase()); k++) {
        const tail = ` (${k})`;
        name = base.slice(0, 31 - tail.length) + tail;
      }
      used.add(name.toLowerCase());
      return name;
    });
  }

  // Книга из нескольких листов: [{ name, header: [...], widths: [...], rows: [[...], ...],
  // side? }]. Шапка жирная и закреплена, на основной таблице автофильтр. Стили —
  // как у файла дублей.
  //
  // side — таблица справа от основной через один пустой столбец (сводка):
  // { header: [...], widths: [...], rows: [[...]], boldLastRow? }. Её строки идут
  // в те же строки листа, что и основная таблица, начиная с шапки.
  function buildSheetsXlsx(sheets) {
    const style = { header: 1, number: 2, text: 3 };
    const names = xlsxSheetNames(sheets.map((sh) => sh.name));
    const sheetXml = sheets.map((sh) => {
      const lastCol = xlsxColumn(Math.max(0, sh.header.length - 1));
      const side = sh.side || null;
      const sideStart = sh.header.length + 1; // один пустой столбец между таблицами
      const sideCells = (values, r, isHeader, isBold) =>
        values
          .map((v, c) =>
            xlsxCell(`${xlsxColumn(sideStart + c)}${r}`, v, isHeader || (isBold && typeof v !== 'object')
              ? { text: style.header, number: style.number }
              : style)
          )
          .join('');
      const total = Math.max(sh.rows.length, side ? side.rows.length : 0);
      const rows = [
        `<row r="1">${sh.header.map((h, c) => xlsxCell(`${xlsxColumn(c)}1`, h, { text: style.header })).join('')}` +
          (side ? sideCells(side.header, 1, true) : '') +
          '</row>',
      ];
      for (let i = 0; i < total; i++) {
        const r = i + 2;
        const main = sh.rows[i] ? sh.rows[i].map((v, c) => xlsxCell(`${xlsxColumn(c)}${r}`, v, style)).join('') : '';
        const extra =
          side && side.rows[i]
            ? sideCells(side.rows[i], r, false, side.boldLastRow && i === side.rows.length - 1)
            : '';
        rows.push(`<row r="${r}">${main}${extra}</row>`);
      }
      const widths = [...(sh.widths || [])];
      if (side) {
        while (widths.length < sideStart) widths.push(4);
        (side.widths || []).forEach((w) => widths.push(w));
      }
      const cols = widths
        .map((w, c) => `<col min="${c + 1}" max="${c + 1}" width="${w}" customWidth="1"/>`)
        .join('');
      return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
        (cols ? `<cols>${cols}</cols>` : '') +
        `<sheetData>${rows.join('')}</sheetData>` +
        `<autoFilter ref="A1:${lastCol}${sh.rows.length + 1}"/>` +
        '</worksheet>'
      );
    });
    const styles =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>' +
      '<font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
      '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>' +
      '<border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="4">' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/>' +
      '<xf numFmtId="1" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>' +
      '</cellXfs>' +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '</styleSheet>';
    const ws = (i) => `xl/worksheets/sheet${i + 1}.xml`;
    return buildZip([
      {
        name: '[Content_Types].xml',
        text:
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
          '<Default Extension="xml" ContentType="application/xml"/>' +
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
          sheets
            .map((_, i) => `<Override PartName="/${ws(i)}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`)
            .join('') +
          '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
          '</Types>',
      },
      {
        name: '_rels/.rels',
        text:
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
          '</Relationships>',
      },
      {
        name: 'xl/workbook.xml',
        text:
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' +
          names.map((n, i) => `<sheet name="${xmlEscape(n)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') +
          '</sheets>' +
          // Автофильтру Excel нужно скрытое имя на каждом листе
          '<definedNames>' +
          sheets
            .map(
              (sh, i) =>
                `<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">` +
                `'${xmlEscape(names[i].replace(/'/g, "''"))}'!$A$1:$${xlsxColumn(Math.max(0, sh.header.length - 1))}$${sh.rows.length + 1}` +
                '</definedName>'
            )
            .join('') +
          '</definedNames></workbook>',
      },
      {
        name: 'xl/_rels/workbook.xml.rels',
        text:
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          sheets
            .map(
              (_, i) =>
                `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`
            )
            .join('') +
          `<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
          '</Relationships>',
      },
      ...sheetXml.map((text, i) => ({ name: ws(i), text })),
      { name: 'xl/styles.xml', text: styles },
    ]);
  }

  function stampedFileName(prefix) {
    const d = new Date();
    const pad = (n) => String(n).padStart(2, '0');
    return (
      `${prefix}_${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}` +
      `_${pad(d.getHours())}-${pad(d.getMinutes())}.xlsx`
    );
  }

  function downloadBlob(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    link.remove();
    // Сразу отзывать нельзя: часть браузеров ещё не успела начать загрузку
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  }

  // Что лежит в колонке Agent's files (documentsExternal). Сайт отдаёт её
  // HTML-ссылками на файлы, как соседнюю колонку файлов поддержки; на случай
  // другого вида понимаем и JSON-массив, и просто текст.
  function parseAgentFiles(raw) {
    if (raw === undefined) return { state: 'missing', count: 0, names: [] };
    const text = String(raw).trim();
    if (text === '' || text === '[]' || text.toLowerCase() === 'null') {
      return { state: 'empty', count: 0, names: [] };
    }
    const baseName = (v) => String(v).split('/').pop().trim();
    if (text.startsWith('[')) {
      try {
        const list = JSON.parse(text);
        if (Array.isArray(list)) {
          const names = list.map((x) => baseName(typeof x === 'string' ? x : JSON.stringify(x))).filter(Boolean);
          return names.length > 0
            ? { state: 'filled', count: names.length, names }
            : { state: 'empty', count: 0, names: [] };
        }
      } catch (e) {
        /* не JSON — разбираем как HTML/текст ниже */
      }
    }
    const anchors = text.match(/<a\b[\s\S]*?<\/a>/gi) || [];
    if (anchors.length > 0) {
      const names = anchors.map((a) => baseName(a.replace(/<[^>]*>/g, ''))).filter(Boolean);
      return { state: 'filled', count: anchors.length, names };
    }
    const plain = text.replace(/<[^>]*>/g, '').trim();
    return plain === ''
      ? { state: 'empty', count: 0, names: [] }
      : { state: 'filled', count: 1, names: [baseName(plain)] };
  }

  function runSiteCheck({ title: titleText, hint: hintText, logName, onDone }) {
    // Номера тикетов с экрана читаются по карте колонок — она должна быть свежей
    refreshColumnIndexMap();
    const existing = document.querySelector('.th-ticket-state-check-overlay');
    if (existing) existing.remove();

    let busy = false;
    let closed = false;
    const overlay = document.createElement('div');
    // Не .modal_wrap и не .swal2-popup: Bulk Approve на той же странице
    // принял бы такое окно за окно сайта
    overlay.className = 'th-ticket-state-check-overlay';
    overlay.style.cssText = [
      'position:fixed', 'inset:0', 'z-index:2147483600',
      'background:rgba(0,0,0,.5)', 'display:flex',
      'align-items:center', 'justify-content:center', 'padding:20px',
    ].join(';');
    const panel = document.createElement('div');
    panel.style.cssText = [
      'background:#fff', 'color:#222', 'border-radius:8px',
      'box-shadow:0 10px 40px rgba(0,0,0,.35)', 'max-width:640px', 'width:100%',
      'max-height:90vh', 'display:flex', 'flex-direction:column',
      'font-family:system-ui,Arial,sans-serif', 'font-size:14px', 'line-height:1.45',
    ].join(';');
    const body = document.createElement('div');
    body.style.cssText = 'padding:18px 20px;overflow:auto;flex:1';
    const title = document.createElement('div');
    title.textContent = titleText;
    title.style.cssText = 'font-weight:600;font-size:16px;margin-bottom:6px';
    const hint = document.createElement('div');
    hint.style.cssText = 'color:#555;margin-bottom:10px';
    hint.textContent = hintText;
    const area = document.createElement('textarea');
    area.rows = 10;
    area.placeholder = '22708318\n22707583';
    area.style.cssText = [
      'width:100%', 'box-sizing:border-box', 'font-family:Consolas,monospace',
      'font-size:13px', 'padding:8px', 'border:1px solid #ccc', 'border-radius:4px',
      'resize:vertical', 'white-space:pre',
    ].join(';');
    const status = document.createElement('div');
    status.className = 'th-ticket-state-check-status';
    status.style.cssText = 'margin-top:10px;white-space:pre-wrap;color:#555';
    body.appendChild(title);
    body.appendChild(hint);
    body.appendChild(area);
    body.appendChild(status);

    const footer = document.createElement('div');
    footer.style.cssText =
      'padding:12px 20px;border-top:1px solid #eee;display:flex;gap:8px;justify-content:flex-end;flex-wrap:wrap';
    const mkBtn = (text, primary) => {
      const b = document.createElement('button');
      b.textContent = text;
      b.style.cssText = [
        'padding:8px 18px', 'border-radius:4px', 'font-size:14px', 'cursor:pointer',
        primary ? 'border:none;background:#6A5ACD;color:#fff' : 'border:1px solid #bbb;background:#f5f5f5',
      ].join(';');
      return b;
    };
    const cancelBtn = mkBtn('Отмена', false);
    const screenIds = () => getTicketRows().map((r) => getTicketIdFromRow(r)).filter(Boolean);
    const screenBtn = mkBtn(`Проверить экран (${screenIds().length})`, false);
    screenBtn.className = 'th-ticket-state-check-screen';
    const listBtn = mkBtn('Проверить список', true);
    listBtn.className = 'th-ticket-state-check-list';
    footer.appendChild(cancelBtn);
    footer.appendChild(screenBtn);
    footer.appendChild(listBtn);
    panel.appendChild(body);
    panel.appendChild(footer);
    overlay.appendChild(panel);

    const close = () => {
      closed = true;
      document.removeEventListener('keydown', onKey);
      overlay.remove();
    };
    const onKey = (e) => {
      if (e.key === 'Escape' && !busy) close();
    };
    document.addEventListener('keydown', onKey);
    cancelBtn.addEventListener('click', close); // во время проверки — прерывает её

    const setBusy = (on) => {
      busy = on;
      [screenBtn, listBtn, area].forEach((el) => {
        el.disabled = on;
        el.style.opacity = on ? '0.5' : '1';
      });
    };
    const say = (text, color) => {
      status.textContent = text;
      status.style.color = color || '#555';
    };

    const check = async (ids, sourceLabel) => {
      if (ids.length === 0) {
        say('Проверять нечего.', '#B00020');
        return;
      }
      if (!siteQueryAvailable()) {
        say('Проверка через сайт в этом браузере недоступна.', '#B00020');
        return;
      }
      setBusy(true);
      const cancelled = () => closed;
      try {
        if (!listRequestTemplate) say('Нажимаю Apply на странице, чтобы подхватить запрос сайта…');
        const hasTemplate = await ensureListRequestTemplate(cancelled);
        if (cancelled()) return;
        if (!hasTemplate) {
          say('Не удалось подхватить запрос, которым страница грузит таблицу. Нажми Apply у фильтров и попробуй снова.', '#B00020');
          return;
        }
        say(`Проверяю на сайте… 0 из ${ids.length}`);
        const res = await queryTicketsOnSite(ids, {
          isCancelled: cancelled,
          onProgress: (done, total) => {
            if (!cancelled()) say(`Проверяю на сайте… ${done} из ${total}`);
          },
        });
        if (cancelled()) return;
        if (!res.ok) {
          say(
            res.expired
              ? `СЕССИЯ ИСТЕКЛА: сайт вместо данных отвечает страницей входа (${res.detail}). Обнови страницу и войди заново.`
              : `Проверка не удалась: ${res.detail}.`,
            '#B00020'
          );
          return;
        }
        close();
        onDone(ids, res.rows, sourceLabel);
      } catch (e) {
        console.error(`[TicketState] ${logName}: проверка упала:`, e);
        if (!closed) say(`Проверка упала: ${(e && e.message) || e}`, '#B00020');
      } finally {
        if (!closed) setBusy(false);
      }
    };

    screenBtn.addEventListener('click', () => {
      if (!busy) check(screenIds(), 'тикеты на экране');
    });
    listBtn.addEventListener('click', () => {
      if (busy) return;
      const parsed = parseTicketIdList(area.value);
      if (parsed.badLines.length > 0) {
        say(
          `Не разобрал строк: ${parsed.badLines.length}\n` +
            parsed.badLines.slice(0, 10).map((b) => `  строка ${b.line}: ${b.why} — «${b.text}»`).join('\n'),
          '#B00020'
        );
        return;
      }
      check([...parsed.pairs.keys()], 'список');
    });

    document.body.appendChild(overlay);
    area.focus();
  }

  // ------------------------------------------------------------------
  // ОТЧЁТ «СОСТОЯНИЕ ТИКЕТОВ».
  //
  // Кнопка, которая только ЧИТАЕТ: по списку Ticket ID (или по экрану)
  // спрашивает сайт тем же запросом списка и показывает, в каком тикет
  // состоянии сейчас, — с файлом Excel. Строки идут в порядке списка;
  // тикеты, которых сайт не нашёл, тоже попадают в файл — с пометкой.
  // ------------------------------------------------------------------
  function showTicketStateReport(ids, rows, sourceLabel) {
    const found = ids.filter((id) => rows.has(id));
    const notFound = ids.filter((id) => !rows.has(id));
    // ГЕО, Реферал, Агент: undefined — такого поля в ответе сайта нет
    const extraNote = (v) => (v === undefined ? 'нет данных' : v);
    const header = [
      'Ticket ID', 'External Status', 'Transaction ID', 'Inside comment',
      'User ID', 'ГЕО', 'Реферал', 'Агент', 'Субагент',
    ];
    const widths = [14, 28, 18, 40, 14, 14, 24, 24, 32];
    const NOT_FOUND = 'сайт не нашёл по фильтру';
    const tableRows = ids.map((id) => {
      const row = rows.get(id);
      if (!row) return [id, NOT_FOUND, ...header.slice(2).map(() => '')];
      return [
        id, row.externalStatus, row.transactionId, row.insideComment,
        row.userId, extraNote(row.geo), extraNote(row.referral),
        extraNote(row.agent), row.subagent,
      ];
    });
    window.__thTicketState = tableRows.map((r) =>
      Object.fromEntries(header.map((h, i) => [h, r[i]]))
    );

    const byStatus = countByStatus(found.map((id) => ({ externalStatus: rows.get(id).externalStatus })));
    const withTx = found.filter((id) => rows.get(id).transactionId !== '').length;
    const states = found.map((id) => parseAgentFiles(rows.get(id).agentFiles).state);
    const filesFilled = states.filter((st) => st === 'filled').length;
    const filesEmpty = states.filter((st) => st === 'empty').length;
    const filesMissing = states.filter((st) => st === 'missing').length;
    const lines = [
      'СОСТОЯНИЕ ТИКЕТОВ',
      `Проверено: ${ids.length} (${sourceLabel}), сайт нашёл: ${found.length}`,
    ];
    if (notFound.length > 0) {
      lines.push(
        `  • сайт не нашёл по фильтру: ${notFound.length} — проверь, не стоит ли на странице ` +
          'ещё какой-то фильтр (даты, статус, отдел)'
      );
    }
    if (found.length > 0) {
      lines.push('', 'External Status:');
      byStatus.forEach(([name, count]) => lines.push(`  • ${name || '(пусто)'}: ${count}`));
      lines.push(
        '',
        `Transaction ID: заполнен ${withTx}, пусто ${found.length - withTx}`,
        `Agent's files: есть файлы ${filesFilled}, пусто ${filesEmpty}` +
          (filesMissing > 0 ? `, нет данных ${filesMissing} (колонки нет в ответе сайта)` : '')
      );
    }
    lines.push('', 'Скрипт только читал — в тикетах ничего не менялось.');

    // Та же таблица текстом — вставить в свою таблицу без файла
    const tsvCell = (v) => String(v == null ? '' : v).replace(/[\t\r\n]+/g, ' ').trim();
    const blocks = [{
      label: `Таблица (${ids.length}) — ${header.join(', ')}:`,
      text: tableRows.map((r) => r.map(tsvCell).join('\t')).join('\n'),
    }];
    if (notFound.length > 0) {
      blocks.push({ label: `Сайт не нашёл (${notFound.length}) — Ticket ID:`, text: notFound.join('\n') });
    }

    const options = {
      downloads: [{
        kind: 'state',
        label: `Excel (${ids.length} ${pluralRu(ids.length, 'тикет', 'тикета', 'тикетов')})`,
        fileName: () => stampedFileName('sostoyanie-tiketov'),
        // Первый лист — все тикеты в порядке списка. Дальше по листу на
        // каждый External Status (от самого частого), в конце — тикеты,
        // которых сайт не нашёл. Строки на листах — в том же порядке списка.
        build: () =>
          buildSheetsXlsx([
            { name: 'Тикеты', header, widths, rows: tableRows },
            ...byStatus.map(([name]) => ({
              name: name === '(пусто)' ? 'Без статуса' : name,
              header,
              widths,
              rows: tableRows.filter((r, i) => rows.has(ids[i]) && (r[1] || '(пусто)') === name),
            })),
            ...(notFound.length > 0
              ? [{ name: 'Сайт не нашёл', header, widths, rows: tableRows.filter((_, i) => !rows.has(ids[i])) }]
              : []),
          ]),
      }],
    };
    showReportWindow(lines.join('\n'), blocks, options);
  }

  // Одна колодка «подпись + поле + кнопка Скопировать». Блоков в окне может
  // быть несколько: таблица и список тикетов, которых сайт не нашёл.
  function appendCopyBlock(body, block) {
    const label = document.createElement('div');
    label.textContent = block.label;
    label.style.cssText = 'margin:18px 0 6px;font-weight:600';
    body.appendChild(label);

    const area = document.createElement('textarea');
    area.className = 'th-ticket-state-report-copy';
    area.readOnly = true;
    area.value = block.text;
    area.rows = Math.min(block.text.split('\n').length + 1, 12);
    area.style.cssText = [
      'width:100%', 'box-sizing:border-box', 'font-family:Consolas,monospace',
      'font-size:13px', 'padding:8px', 'border:1px solid #ccc',
      'border-radius:4px', 'resize:vertical', 'white-space:pre',
    ].join(';');
    area.addEventListener('focus', () => area.select());
    body.appendChild(area);

    const copyBtn = document.createElement('button');
    copyBtn.textContent = 'Скопировать';
    copyBtn.style.cssText = [
      'margin-top:8px', 'padding:8px 16px', 'border:none',
      'border-radius:4px', 'background:#2ABFCF', 'color:#fff',
      'font-size:14px', 'cursor:pointer',
    ].join(';');
    copyBtn.addEventListener('click', async () => {
      // Выделяем в любом случае: даже если запись в буфер не пройдёт,
      // человек сможет нажать Ctrl+C сам.
      area.focus();
      area.select();
      let ok = false;
      try {
        if (navigator.clipboard && navigator.clipboard.writeText) {
          await navigator.clipboard.writeText(block.text);
          ok = true;
        }
      } catch (e) {
        ok = false;
      }
      if (!ok) {
        try {
          ok = document.execCommand('copy');
        } catch (e) {
          ok = false;
        }
      }
      copyBtn.textContent = ok ? 'Скопировано ✓' : 'Выделено — нажми Ctrl+C';
      copyBtn.style.background = ok ? '#3BA55D' : '#E0A030';
      setTimeout(() => {
        copyBtn.textContent = 'Скопировать';
        copyBtn.style.background = '#2ABFCF';
      }, 2500);
    });
    body.appendChild(copyBtn);
  }

  function showReportWindow(reportText, copyBlocks, options) {
    // Если что-то пойдёт не так с вёрсткой — отчёт всё равно должен дойти
    // до человека, поэтому любой сбой откатывается на обычный alert.
    try {
      const existing = document.querySelector('.th-ticket-state-report-overlay');
      if (existing) existing.remove();

      const overlay = document.createElement('div');
      overlay.className = 'th-ticket-state-report-overlay';
      overlay.style.cssText = [
        'position:fixed', 'inset:0', 'z-index:2147483600',
        'background:rgba(0,0,0,.5)', 'display:flex',
        'align-items:center', 'justify-content:center', 'padding:20px',
      ].join(';');

      const panel = document.createElement('div');
      panel.style.cssText = [
        'background:#fff', 'color:#222', 'border-radius:8px',
        'box-shadow:0 10px 40px rgba(0,0,0,.35)', 'max-width:820px',
        'width:100%', 'max-height:90vh', 'display:flex',
        'flex-direction:column', 'font-family:system-ui,Arial,sans-serif',
        'font-size:14px', 'line-height:1.45',
      ].join(';');

      const body = document.createElement('div');
      body.style.cssText = 'padding:18px 20px;overflow:auto;flex:1';

      const report = document.createElement('pre');
      report.className = 'th-ticket-state-report-text';
      report.textContent = reportText;
      report.style.cssText = [
        'margin:0', 'white-space:pre-wrap', 'word-break:break-word',
        'font-family:inherit', 'font-size:14px', 'user-select:text',
      ].join(';');
      body.appendChild(report);

      (copyBlocks || [])
        .filter((block) => block && block.text !== '')
        .forEach((block) => appendCopyBlock(body, block));

      const footer = document.createElement('div');
      footer.style.cssText =
        'padding:12px 20px;border-top:1px solid #eee;display:flex;gap:8px;' +
        'justify-content:flex-end;flex-shrink:0;flex-wrap:wrap';


      // Другие файлы Excel к отчёту: [{ kind, label, fileName, build: () => Blob }]
      const downloads = (options && options.downloads) || [];
      downloads.forEach((download) => {
        const fileBtn = document.createElement('button');
        fileBtn.className =
          'th-ticket-state-report-download' +
          (download.kind ? ` th-ticket-state-report-download-${download.kind}` : '');
        fileBtn.textContent = download.label;
        fileBtn.style.cssText = [
          'padding:8px 20px', 'border:none', 'border-radius:4px',
          'background:#E8590C', 'color:#fff', 'font-size:14px', 'font-weight:600',
          'cursor:pointer', 'box-shadow:0 2px 6px rgba(232,89,12,.4)',
        ].join(';');
        fileBtn.addEventListener('click', () => {
          try {
            downloadBlob(download.build(), download.fileName());
          } catch (e) {
            console.error('[TicketState] Не удалось собрать файл Excel:', e);
            alert('Не удалось собрать файл Excel — подробности в консоли (F12). Данные есть в блоках для копирования ниже.');
          }
        });
        footer.appendChild(fileBtn);
      });

      const closeBtn = document.createElement('button');
      closeBtn.textContent = 'Закрыть';
      closeBtn.style.cssText = [
        'padding:8px 20px', 'border:1px solid #bbb', 'border-radius:4px',
        'background:#f5f5f5', 'font-size:14px', 'cursor:pointer',
      ].join(';');
      footer.appendChild(closeBtn);

      // Закрывается ТОЛЬКО кнопкой и только после подтверждения. Клик мимо
      // окна и Esc больше не закрывают: при выделении текста мышь легко
      // уезжает за край панели, и отчёт пропадал вместе с несохранённым.
      closeBtn.addEventListener('click', () => {
        if (!confirm('Точно хотите закрыть отчёт?')) return;
        overlay.remove();
      });

      panel.appendChild(body);
      panel.appendChild(footer);
      overlay.appendChild(panel);
      document.body.appendChild(overlay);
      closeBtn.focus();
    } catch (e) {
      console.error('[TicketState] Не удалось показать окно отчёта, показываю обычный alert:', e);
      alert(reportText);
    }
  }

  // ------------------------------------------------------------------
  // КНОПКА — слева внизу, рядом с шестерёнкой Team Helper, чтобы не
  // спорить за правый угол с Bulk Approve.
  // ------------------------------------------------------------------
  function addButton() {
    if (document.getElementById('th-ticket-state-btn')) return;
    const btn = document.createElement('button');
    btn.id = 'th-ticket-state-btn';
    btn.type = 'button';
    btn.textContent = 'Состояние тикетов';
    btn.title = 'Текущее состояние тикетов по списку Ticket ID, с файлом Excel';
    Object.assign(btn.style, {
      position: 'fixed',
      left: '72px',
      bottom: '24px',
      zIndex: 99999,
      padding: '8px 14px',
      background: '#4F6D8F',
      color: '#fff',
      border: 'none',
      borderRadius: '999px',
      fontFamily: 'system-ui,Arial,sans-serif',
      fontSize: '13px',
      fontWeight: '600',
      cursor: 'pointer',
      boxShadow: '0 4px 12px rgba(0,0,0,0.25)',
    });
    btn.addEventListener('click', () => {
      try {
        runTicketStateCheck();
      } catch (e) {
        console.error('[TicketState] Состояние тикетов:', e);
      }
    });
    document.body.appendChild(btn);
  }

  function runTicketStateCheck() {
    runSiteCheck({
      title: 'Состояние тикетов',
      hint:
        'Скрипт только читает: спросит сайт о каждом тикете и покажет его текущее состояние — ' +
        'External Status, Transaction ID, Inside comment, User ID, ГЕО, реферала, ' +
        'агента и субагента — с файлом Excel. Окна Edit не открываются, в тикетах ничего не меняется. Вставь список ' +
        'Ticket ID — по одному в строке — или проверь тикеты, которые сейчас на экране.',
      logName: 'Состояние тикетов',
      onDone: showTicketStateReport,
    });
  }

  installListCapture();
  if (document.body) addButton();
  else document.addEventListener('DOMContentLoaded', addButton);
})();
