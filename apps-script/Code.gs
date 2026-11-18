/**
 * TRA CỨU SINH HOẠT KHOA HỌC / HỘI THẢO KHOA HỌC 2026
 * Backend (Google Apps Script).
 *
 * Sheet "Tra cứu"          : A STT | B Ngày tổ chức | C Tên hội thảo | D Họ và tên | E Email | F Khoa/phòng
 * Sheet "DS các buổi event": A STT | B Ngày tổ chức | C Tên hội thảo   (danh sách đầy đủ các buổi -> mẫu số XX/30)
 * Sheet "DS Khoa/phòng"    : A STT | B Khoa/phòng
 * (Mã tự dò cột theo tiêu đề dòng 1; không dò được thì dùng đúng thứ tự trên.)
 *
 * API (GET):
 *   ?action=suggest&q=...                -> gợi ý (tối thiểu 5 ký tự)
 *   ?action=search&type=name|email&key=  -> các buổi đã / chưa tham dự của 1 người
 *   ?action=config                       -> {deptLocked}
 *   ?action=depts[&pass=]                -> danh sách khoa/phòng + số người
 *   ?action=dept&name=...[&pass=]        -> người của 1 khoa/phòng + số buổi
 */

const SPREADSHEET_ID = '16MUn8yW0lqvdiGRc98q5OYKmiCHx3LoBLgozwyj9YHc';
const SHEET_NAME     = 'Tra cứu';
const EVENTS_SHEET   = 'DS các buổi event';
const DEPTS_SHEET    = 'DS Khoa/phòng';
const MIN_CHARS      = 5;      // chỉ gợi ý khi gõ từ ký tự thứ 5
const MAX_SUGGEST    = 8;      // số gợi ý tối đa
const CACHE_SECONDS  = 600;    // dữ liệu nhớ tạm 10 phút (sửa sheet -> tối đa 10 phút sau web mới cập nhật)
const DEPT_PASSCODE  = '';     // để trống = ai cũng xem được danh sách theo khoa/phòng; đặt mã (vd 'abc123') để khóa
const CODE_VERSION   = '2026-11-18.4';
const CK             = 'v4_';  // tiền tố cache (đổi khi đổi cấu trúc dữ liệu)

/* ------------------------------ ENTRY ------------------------------ */

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    switch (p.action) {
      case 'suggest': return output(suggest_(p.q));
      case 'search':  return output(search_(p.type, p.key));
      case 'config':  return output({ deptLocked: !!DEPT_PASSCODE });
      case 'depts':   return output(authorized_(p.pass) ? depts_() : { needPass: true });
      case 'dept':    return output(authorized_(p.pass) ? dept_(p.name) : { needPass: true });
    }
    return output({ status: 'ok', version: CODE_VERSION });
  } catch (err) {
    return output({ error: String(err && err.message || err) });
  }
}

function output(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function authorized_(pass) {
  return !DEPT_PASSCODE || String(pass || '') === DEPT_PASSCODE;
}

/* ------------------------------ TEXT ------------------------------ */

// "Lê Thị Ngọc Châm" -> "le thi ngoc cham"
function norm_(s) {
  return String(s == null ? '' : s)
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/đ/g, 'd').replace(/Đ/g, 'd')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

// Chỉ giữ chữ + số, dùng để so khớp tên sự kiện (bỏ dấu câu, ngoặc kép, khoảng trắng thừa)
function alnum_(s) { return norm_(s).replace(/[^a-z0-9]/g, ''); }

// "email@gmail.com" -> "Em**l@gmail.com"
function maskEmail_(email) {
  email = String(email || '');
  const i = email.indexOf('@');
  const local = i < 0 ? email : email.slice(0, i);
  const domain = i < 0 ? '' : email.slice(i);
  if (!local) return '';
  let head;
  if (local.length <= 3) head = local.slice(0, 1) + '**';
  else head = local.slice(0, 2) + '**' + local.slice(-1);
  return head.charAt(0).toUpperCase() + head.slice(1) + domain;
}

// Khoa/phòng: ưu tiên giá trị cụ thể hơn "Khác", rồi đến giá trị xuất hiện nhiều nhất
function topDept_(counts) {
  let best = '', n = 0, bestOther = '', nOther = 0;
  for (const k in counts) {
    if (counts[k] > n) { best = k; n = counts[k]; }
    if (norm_(k) !== 'khac' && counts[k] > nOther) { bestOther = k; nOther = counts[k]; }
  }
  return bestOther || best;
}

function topEmail_(counts) {
  let best = '', n = 0;
  for (const k in counts) if (counts[k] > n) { best = k; n = counts[k]; }
  return best;
}

/* ------------------------------ DATE ------------------------------ */

function fmtDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy');
  const s = String(v == null ? '' : v).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  return m ? ('0' + m[1]).slice(-2) + '/' + ('0' + m[2]).slice(-2) + '/' + m[3] : s;
}

function dateKey_(s) { // dd/mm/yyyy (hoặc "26-27/09/2026" -> lấy ngày cuối) -> yyyymmdd
  const m = String(s).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  return m ? (+m[3]) * 10000 + (+m[2]) * 100 + (+m[1]) : 0;
}

function todayKey_() {
  return +Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyyMMdd');
}

/* ------------------------------ SHEET HELPERS ------------------------------ */

function findCol_(header, patterns, fallback) {
  for (let i = 0; i < header.length; i++) {
    const h = norm_(header[i]);
    if (patterns.some(function (p) { return h.indexOf(p) >= 0; })) return i;
  }
  return fallback;
}

function cacheGet_(name) {
  const cache = CacheService.getScriptCache();
  const meta = cache.get(CK + name + '_n');
  if (!meta) return null;
  const keys = [];
  for (let i = 0; i < +meta; i++) keys.push(CK + name + '_' + i);
  const parts = cache.getAll(keys);
  if (!keys.every(function (k) { return parts[k]; })) return null;
  return JSON.parse(keys.map(function (k) { return parts[k]; }).join(''));
}

function cachePut_(name, obj) { // chia mảnh <= 90KB (giới hạn 100KB/khóa)
  try {
    const json = JSON.stringify(obj), size = 90000, put = {};
    let n = 0;
    for (let i = 0; i < json.length; i += size) put[CK + name + '_' + n++] = json.slice(i, i + size);
    put[CK + name + '_n'] = String(n);
    CacheService.getScriptCache().putAll(put, CACHE_SECONDS);
  } catch (err) { /* không cache được vẫn chạy bình thường */ }
}

/* ------------------------------ DATA ------------------------------ */

// Danh sách đầy đủ các buổi: [{date, name}] (đúng thứ tự trong sheet)
function loadEvents_() {
  const hit = cacheGet_('events');
  if (hit) return hit;

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(EVENTS_SHEET);
  const out = [];
  if (sheet && sheet.getLastRow() >= 2) {
    const data = sheet.getRange(1, 1, sheet.getLastRow(), Math.max(sheet.getLastColumn(), 3)).getValues();
    const cDate = findCol_(data[0], ['ngay'], 1);
    const cName = findCol_(data[0], ['ten hoi thao', 'hoi thao', 'sinh hoat'], 2);
    for (let r = 1; r < data.length; r++) {
      const name = String(data[r][cName] || '').replace(/\s+/g, ' ').trim();
      if (name) out.push({ date: fmtDate_(data[r][cDate]), name: name });
    }
  }
  cachePut_('events', out);
  return out;
}

// Ghép 1 dòng điểm danh với danh sách buổi: theo tên (bỏ dấu câu), không được thì theo ngày (nếu ngày đó chỉ có 1 buổi)
function buildEventIndex_(events) {
  const byName = {}, byDate = {};
  events.forEach(function (ev, i) {
    byName[alnum_(ev.name)] = i;
    byDate[ev.date] = byDate[ev.date] === undefined ? i : -1;
  });
  return function (date, name) {
    const a = alnum_(name);
    if (byName[a] !== undefined) return byName[a];
    const d = byDate[date];
    return d === undefined ? -1 : d;
  };
}

// Mảng gọn: [ngày, tên sự kiện, họ tên, email(thường), họ tên không dấu, khoa/phòng, chỉ số buổi (-1 nếu không khớp)]
function loadRows_() {
  const hit = cacheGet_('rows');
  if (hit) return hit;

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const data = sheet.getRange(1, 1, last, Math.max(sheet.getLastColumn(), 5)).getValues();

  const h = data[0];
  const cDate  = findCol_(h, ['ngay'], 1);
  const cEvent = findCol_(h, ['ten hoi thao', 'hoi thao', 'sinh hoat'], 2);
  const cName  = findCol_(h, ['ho va ten', 'ho ten'], 3);
  const cMail  = findCol_(h, ['email', 'mail'], 4);
  const cDept  = findCol_(h, ['khoa', 'phong'], 5);
  const match  = buildEventIndex_(loadEvents_());

  const rows = [];
  for (let r = 1; r < data.length; r++) {
    const name = String(data[r][cName] || '').replace(/\s+/g, ' ').trim();
    const event = String(data[r][cEvent] || '').replace(/\s+/g, ' ').trim();
    if (!name || !event) continue;
    const date = fmtDate_(data[r][cDate]);
    rows.push([
      date,
      event,
      name,
      String(data[r][cMail] || '').trim().toLowerCase(),
      norm_(name),
      String(data[r][cDept] || '').replace(/\s+/g, ' ').trim(),
      match(date, event)
    ]);
  }
  cachePut_('rows', rows);
  return rows;
}

// id duy nhất của 1 buổi (để đếm không trùng)
function idOf_(r) { return r[6] >= 0 ? r[6] : 'u' + r[0] + '|' + r[1]; }

/* ------------------------------ SUGGEST ------------------------------ */

function suggest_(q) {
  const raw = String(q || '').trim().toLowerCase();
  const nq = norm_(q);
  const total = loadEvents_().length;
  if (nq.replace(/\s/g, '').length < MIN_CHARS) return { suggestions: [], minChars: MIN_CHARS, total: total };

  const rows = loadRows_();
  const isEmail = raw.indexOf('@') >= 0;
  const found = {};

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    let f;
    if (isEmail) {
      const mail = r[3];
      if (!mail || mail.indexOf(raw) < 0) continue;
      f = found[mail] || (found[mail] = {
        type: 'email', key: mail, label: maskEmail_(mail), sub: r[2], depts: {}, ids: {},
        rank: mail.indexOf(raw) === 0 ? 0 : 1
      });
    } else {
      const pos = r[4].indexOf(nq);
      if (pos < 0) continue;
      f = found[r[4]] || (found[r[4]] = {
        type: 'name', key: r[4], label: r[2], sub: '', mails: {}, depts: {}, ids: {},
        rank: pos === 0 ? 0 : (r[4].indexOf(' ' + nq) >= 0 ? 1 : 2)
      });
      if (r[3]) f.mails[r[3]] = (f.mails[r[3]] || 0) + 1;
    }
    if (r[6] >= 0) f.ids[r[6]] = 1; // chỉ đếm các buổi có trong DS (khớp mẫu số)
    if (r[5]) f.depts[r[5]] = (f.depts[r[5]] || 0) + 1;
  }

  const list = Object.keys(found).map(function (k) {
    const f = found[k];
    f.count = Object.keys(f.ids).length;
    return f;
  })
    .sort(function (a, b) { return a.rank - b.rank || b.count - a.count || a.label.localeCompare(b.label); })
    .slice(0, MAX_SUGGEST)
    .map(function (f) {
      if (f.type === 'name') { f.email = maskEmail_(topEmail_(f.mails)); delete f.mails; }
      f.dept = topDept_(f.depts);
      delete f.depts; delete f.ids;
      return f;
    });

  return { suggestions: list, minChars: MIN_CHARS, total: total };
}

/* ------------------------------ SEARCH ------------------------------ */

function search_(type, key) {
  const rows = loadRows_();
  const events = loadEvents_();
  const k = type === 'email' ? String(key || '').trim().toLowerCase() : norm_(key);
  if (!k) return { person: '', events: [], missing: [], count: 0, total: events.length };

  const col = type === 'email' ? 3 : 4;
  const attended = [], seen = {}, mails = {}, depts = {}, got = {};
  let person = '';

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (r[col] !== k) continue;
    if (!person) person = r[2];
    if (r[3]) mails[r[3]] = (mails[r[3]] || 0) + 1;
    if (r[5]) depts[r[5]] = (depts[r[5]] || 0) + 1;
    const id = idOf_(r);
    if (seen[id]) continue; // bỏ trùng cùng buổi
    seen[id] = 1;
    if (r[6] >= 0) got[r[6]] = 1;
    attended.push({ date: r[0], event: r[1] });
  }

  attended.sort(function (a, b) { return dateKey_(b.date) - dateKey_(a.date); });

  const today = todayKey_(), missing = [];
  events.forEach(function (ev, i) {
    if (!got[i]) missing.push({ date: ev.date, event: ev.name, upcoming: dateKey_(ev.date) > today });
  });

  const emails = Object.keys(mails).sort(function (a, b) { return mails[b] - mails[a]; })
    .slice(0, 3).map(maskEmail_);

  return {
    person: person, emails: emails, dept: topDept_(depts),
    events: attended, missing: missing,
    count: Object.keys(got).length, total: events.length
  };
}

/* ------------------------------ KHOA / PHÒNG ------------------------------ */

// Mỗi người (gộp theo họ tên không dấu): {key, name, dept, count}
function people_() {
  const rows = loadRows_(), m = {};
  rows.forEach(function (r) {
    const p = m[r[4]] || (m[r[4]] = { name: r[2], depts: {}, ids: {} });
    if (r[5]) p.depts[r[5]] = (p.depts[r[5]] || 0) + 1;
    if (r[6] >= 0) p.ids[r[6]] = 1;
  });
  return Object.keys(m).map(function (k) {
    const p = m[k];
    return { key: k, name: p.name, dept: topDept_(p.depts), count: Object.keys(p.ids).length };
  });
}

function depts_() {
  const people = people_(), counts = {}, names = {};
  people.forEach(function (p) {
    if (!p.dept) return;
    const n = norm_(p.dept);
    counts[n] = (counts[n] || 0) + 1;
    names[n] = names[n] || p.dept;
  });

  const list = [], used = {};
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(DEPTS_SHEET);
  if (sheet && sheet.getLastRow() >= 2) {
    const data = sheet.getRange(1, 1, sheet.getLastRow(), Math.max(sheet.getLastColumn(), 2)).getValues();
    const c = findCol_(data[0], ['khoa', 'phong'], 1);
    for (let r = 1; r < data.length; r++) {
      const name = String(data[r][c] || '').replace(/\s+/g, ' ').trim();
      if (!name || used[norm_(name)]) continue;
      used[norm_(name)] = 1;
      list.push({ name: name, people: counts[norm_(name)] || 0 });
    }
  }
  Object.keys(counts).forEach(function (n) { // khoa/phòng có trong dữ liệu nhưng chưa có trong DS
    if (!used[n]) list.push({ name: names[n], people: counts[n] });
  });
  return { depts: list, total: loadEvents_().length };
}

function dept_(name) {
  const n = norm_(name);
  const list = people_().filter(function (p) { return p.dept && norm_(p.dept) === n; })
    .sort(function (a, b) { return b.count - a.count || a.name.localeCompare(b.name, 'vi'); })
    .map(function (p) { return { key: p.key, name: p.name, count: p.count }; });
  return { dept: String(name || ''), people: list, total: loadEvents_().length };
}

/* ------------------------------ KIỂM TRA ------------------------------ */

// Chạy tay trong editor để xem tên sự kiện nào ở sheet "Tra cứu" chưa khớp với "DS các buổi event"
function checkEventMatching() {
  const rows = loadRows_(), bad = {};
  rows.forEach(function (r) { if (r[6] < 0) bad[r[0] + ' | ' + r[1]] = (bad[r[0] + ' | ' + r[1]] || 0) + 1; });
  const keys = Object.keys(bad);
  Logger.log(keys.length ? 'Chưa khớp (' + keys.length + '):\n' + keys.map(function (k) { return k + '  x' + bad[k]; }).join('\n')
                         : 'Tất cả các dòng đều khớp với DS các buổi event.');
}
