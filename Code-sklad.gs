/**
 * Склад — сервер общей таблицы (Google Apps Script).
 * Развернуть → Новое развёртывание → Веб-приложение,
 * «Выполнять от имени»: Я, «У кого есть доступ»: Все.
 * Лист «_данные» служебный, не трогай. Листы «Остатки», «Движения»,
 * «Инвентаризации» пересобираются при каждой отправке, правки в них руками затираются.
 */
const ACCESS_KEY = 'ВПИШИ_КЛЮЧ';

const DATA_SHEET = '_данные';
const COLLECTIONS = ['items', 'moves', 'invs'];
const CHUNK = 45000;

function doGet() {
  return out_({ ok: true, message: 'Склад: сервер работает' });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); }
  catch (err) { return out_({ ok: false, error: 'bad_json' }); }
  if (!ACCESS_KEY || ACCESS_KEY === 'ВПИШИ_КЛЮЧ' || ACCESS_KEY === 'ЗАМЕНИ_НА_СВОЙ_КЛЮЧ') return out_({ ok: false, error: 'server', message: 'в коде таблицы не задан ключ доступа' });
  if (String(req.key || '') !== ACCESS_KEY) return out_({ ok: false, error: 'bad_key' });
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return out_({ ok: false, error: 'server', message: 'таблица занята, повтори через минуту' });
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = dataSheet_(ss);
    const docs = Array.isArray(req.docs) ? req.docs : [];
    if (docs.length) upsert_(sheet, docs);
    const all = readAll_(sheet);
    if (docs.length) rebuild_(ss, all);
    return out_({ ok: true, data: all, at: new Date().toISOString() });
  } catch (err) {
    return out_({ ok: false, error: 'server', message: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function dataSheet_(ss) {
  let sh = ss.getSheetByName(DATA_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DATA_SHEET);
    sh.getRange(1, 1, 1, 4).setValues([['раздел', 'id', 'обновлено', 'данные']]);
    sh.setFrozenRows(1);
    sh.hideSheet();
  }
  return sh;
}

function upsert_(sheet, docs) {
  const last = sheet.getLastRow();
  const index = {};
  if (last > 1) sheet.getRange(2, 1, last - 1, 2).getValues().forEach((r, i) => { index[r[0] + '|' + r[1]] = i + 2; });
  let next = Math.max(last, 1) + 1;
  const now = new Date().toISOString();
  const fresh = [];
  docs.forEach(d => {
    if (!d || COLLECTIONS.indexOf(d.c) < 0 || !d.id || !d.data) return;
    const id = String(d.id), json = JSON.stringify(d.data), parts = [];
    for (let i = 0; i < json.length; i += CHUNK) parts.push(json.slice(i, i + CHUNK));
    const row = [d.c, id, now].concat(parts);
    const r = index[d.c + '|' + id];
    if (r) {
      const width = Math.max(sheet.getLastColumn(), row.length);
      sheet.getRange(r, 1, 1, width).clearContent();
      sheet.getRange(r, 1, 1, row.length).setNumberFormat('@').setValues([row]);
    } else {
      index[d.c + '|' + id] = next++;
      fresh.push(row);
    }
  });
  if (fresh.length) {
    const w = Math.max.apply(null, fresh.map(r => r.length));
    const block = fresh.map(r => r.concat(new Array(w - r.length).fill('')));
    sheet.getRange(Math.max(last, 1) + 1, 1, block.length, w).setNumberFormat('@').setValues(block);
  }
}

function readAll_(sheet) {
  const res = {};
  COLLECTIONS.forEach(c => res[c] = []);
  const last = sheet.getLastRow();
  if (last < 2) return res;
  sheet.getRange(2, 1, last - 1, Math.max(4, sheet.getLastColumn())).getValues().forEach(r => {
    if (!res[r[0]]) return;
    try { res[r[0]].push(JSON.parse(r.slice(3).join(''))); } catch (e) { }
  });
  return res;
}

function rebuild_(ss, all) {
  const tz = ss.getSpreadsheetTimeZone();
  const t = iso => { if (!iso) return ''; const d = new Date(iso); return isNaN(d) ? '' : Utilities.formatDate(d, tz, 'dd.MM.yyyy HH:mm'); };
  const n = v => { const x = parseFloat(String(v === undefined || v === null ? '' : v).replace(',', '.')); return isFinite(x) ? x : 0; };
  const items = {};
  all.items.forEach(i => { if (i && i.code && !i.deleted) items[i.code] = i; });
  const stock = {};
  all.moves.forEach(m => { stock[m.code] = (stock[m.code] || 0) + (m.type === 'out' ? -n(m.qty) : m.type === 'move' ? 0 : n(m.qty)); });
  const TYPES = { in: 'Приход', out: 'Расход', move: 'Перемещение', adj: 'Инвентаризация' };

  const rows = Object.keys(items).map(c => {
    const i = items[c], s = Math.round((stock[c] || 0) * 1000) / 1000, min = n(i.min);
    return [i.code, i.name || '', i.art || '', i.unit || 'шт', i.loc || '', i.cat || '', s, min || '', min > 0 && s < min ? 'ЗАКАЗАТЬ' : ''];
  }).sort((a, b) => String(a[4]).localeCompare(String(b[4])) || String(a[1]).localeCompare(String(b[1])));
  write_(ss, 'Остатки', ['Код', 'Наименование', 'Артикул', 'Ед.', 'Адрес', 'Группа', 'Остаток', 'Мин. остаток', 'Ниже минимума'], rows);

  const moves = all.moves.slice().sort((a, b) => String(b.at || '').localeCompare(String(a.at || '')));
  write_(ss, 'Движения', ['Дата', 'Операция', 'Код', 'Наименование', 'Кол-во', 'Ед.', 'Кому / от кого', 'Откуда', 'Куда', 'Примечание', 'Кто'],
    moves.map(m => {
      const i = items[m.code] || {};
      return [t(m.at), TYPES[m.type] || m.type, m.code, i.name || '', m.type === 'move' ? '' : (m.type === 'out' ? -n(m.qty) : n(m.qty)),
        i.unit || 'шт', m.whom || '', m.from || '', m.to || '', m.note || '', m.by || ''];
    }));

  const inv = [];
  all.invs.filter(v => v.closed && !v.cancelled).sort((a, b) => String(b.closedAt || '').localeCompare(String(a.closedAt || ''))).forEach(v => {
    Object.keys(v.counts || {}).forEach(c => {
      const i = items[c] || {}, x = v.counts[c];
      inv.push([v.name || '', t(v.closedAt), c, i.name || '', i.loc || '', n(x.q), x.by || '', t(x.at)]);
    });
  });
  write_(ss, 'Инвентаризации', ['Инвентаризация', 'Проведена', 'Код', 'Наименование', 'Адрес', 'Факт', 'Посчитал', 'Время подсчёта'], inv);
}

function write_(ss, name, head, rows) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clearContents();
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  if (rows.length) sh.getRange(2, 1, rows.length, head.length).setValues(rows);
  sh.setFrozenRows(1);
}
