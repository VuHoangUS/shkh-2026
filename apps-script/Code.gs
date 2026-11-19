/**
 * TRA CỨU SINH HOẠT KHOA HỌC / HỘI THẢO KHOA HỌC 2026
 * Backend (Google Apps Script).
 *
 * Sheet "Tra cứu"          : A STT | B Ngày tổ chức | C Tên hội thảo | D Họ và tên | E Email | F Khoa/phòng
 * Sheet "DS các buổi event": A STT | B Ngày tổ chức | C Tên hội thảo | D Số tiết | E Số tiết cho BCV/chủ tọa
 *                              (danh sách đầy đủ các buổi -> mẫu số XX/30; dòng thiếu ngày như "Ghi chú" bị bỏ qua)
 * Sheet "DS Khoa/phòng"    : A STT | B Khoa/phòng
 * Sheet "DS người tham dự theo khoa/phòng": A STT | B Khoa/phòng | C Họ và tên | D Số buổi | E % | F Số tiết tích lũy
 *   -> KHÔNG dùng làm nguồn cho web; web tự tính từ "Tra cứu" + "DS các buổi event". Chạy rebuildPeopleSheet() để ghi lại sheet này.
 * Sheet "Tra cứu" có thêm cột G "Vai trò" (trống = Đại biểu tham dự; Chủ tọa; Báo cáo viên)
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
const PEOPLE_SHEET   = 'DS người tham dự theo khoa/phòng'; // sheet kết quả: chạy rebuildPeopleSheet() để tính lại
const TIET_DELEGATE  = 3;      // mặc định (khi sheet buổi không ghi số tiết): Đại biểu tham dự
const TIET_SPEAKER   = 4;      // mặc định: Chủ tọa, Báo cáo viên
const MIN_CHARS      = 5;      // chỉ gợi ý khi gõ từ ký tự thứ 5
const MAX_SUGGEST    = 8;      // số gợi ý tối đa
const CACHE_SECONDS  = 600;    // dữ liệu nhớ tạm 10 phút (sửa sheet -> tối đa 10 phút sau web mới cập nhật)
const DEPT_PASSCODE  = '';     // để trống = ai cũng xem được danh sách theo khoa/phòng; đặt mã (vd 'abc123') để khóa
const CODE_VERSION   = '2026-11-19.4';
const CK             = 'v8_';  // tiền tố cache (đổi khi đổi cấu trúc dữ liệu)

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

/* ------------------------------ VAI TRÒ / TIẾT ------------------------------ */

// Ô "Vai trò" -> mảng nhãn. Trống = "Đại biểu tham dự".
function rolesOf_(raw) {
  const t = norm_(raw), out = [];
  if (!t) return ['Đại biểu tham dự'];
  if (t.indexOf('chu toa') >= 0) out.push('Chủ tọa');
  if (t.indexOf('bao cao') >= 0) out.push('Báo cáo viên');
  if (!out.length) out.push(t.indexOf('dai bieu') >= 0 ? 'Đại biểu tham dự' : String(raw).replace(/\s+/g, ' ').trim());
  return out;
}

// Số tiết của 1 buổi theo vai trò; ev = {tiet, tietSp} lấy từ sheet "DS các buổi event" (không có thì dùng mặc định)
function tietOf_(roles, ev) {
  const sp = roles.some(function (r) { return r === 'Chủ tọa' || r === 'Báo cáo viên'; });
  if (sp) return ev && ev.tietSp != null ? ev.tietSp : TIET_SPEAKER;
  return ev && ev.tiet != null ? ev.tiet : TIET_DELEGATE;
}

function num_(v) {
  const x = parseFloat(String(v == null ? '' : v).replace(',', '.'));
  return isNaN(x) ? null : x;
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
    let cTiet = 3, cTietSp = 4;
    data[0].forEach(function (x, i) {
      const t = norm_(x);
      if (t === 'so tiet') cTiet = i;
      else if (t.indexOf('tiet') >= 0 && (t.indexOf('bcv') >= 0 || t.indexOf('chu toa') >= 0 || t.indexOf('bao cao') >= 0)) cTietSp = i;
    });
    for (let r = 1; r < data.length; r++) {
      const name = String(data[r][cName] || '').replace(/\s+/g, ' ').trim();
      const date = fmtDate_(data[r][cDate]);
      if (!name || !date) continue; // bỏ dòng trống / dòng ghi chú (không có ngày)
      out.push({ date: date, name: name, tiet: num_(data[r][cTiet]), tietSp: num_(data[r][cTietSp]) });
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

// Mảng gọn: [ngày, tên sự kiện, họ tên, email(thường), họ tên không dấu, khoa/phòng, chỉ số buổi (-1 nếu không khớp), [vai trò]]
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
  const cRole  = findCol_(h, ['vai tro'], 6);
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
      match(date, event),
      rolesOf_(data[r][cRole])
    ]);
  }
  assignPeople_(rows);
  cachePut_('rows', rows);
  return rows;
}

// Tách người trùng tên: cùng họ tên nhưng khác khoa/phòng (cụ thể) => 2 người khác nhau.
// Dòng ghi "Khác"/trống được gán vào người có cùng email; nếu chỉ có 1 khoa cụ thể thì gán vào đó;
// nếu không xác định được (nhiều khoa, không khớp email) thì để riêng thành "Khác".
// Thêm r[8] = "họ tên không dấu|khoa/phòng không dấu".
function assignPeople_(rows) {
  const byName = {};
  rows.forEach(function (r) { (byName[r[4]] = byName[r[4]] || []).push(r); });
  Object.keys(byName).forEach(function (n) {
    const rs = byName[n], spec = {};
    rs.forEach(function (r) {
      const d = norm_(r[5]);
      if (d && d !== 'khac') { const o = spec[d] || (spec[d] = {}); if (r[3]) o[r[3]] = 1; }
    });
    const keys = Object.keys(spec);
    rs.forEach(function (r) {
      const d = norm_(r[5]);
      let id;
      if (d && d !== 'khac') id = d;
      else if (!keys.length) id = 'khac';
      else {
        const c = keys.filter(function (k) { return r[3] && spec[k][r[3]]; });
        id = c.length === 1 ? c[0] : (keys.length === 1 ? keys[0] : 'khac');
      }
      r.push(n + '|' + id);
    });
  });
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
      f = found[r[8]] || (found[r[8]] = {
        type: 'name', key: r[8], label: r[2], sub: '', mails: {}, depts: {}, ids: {},
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
  if (!k) return { person: '', events: [], missing: [], count: 0, total: events.length, tiet: 0 };

  let col = 4, kk = k;
  if (type === 'email') { // email -> người (tên + khoa/phòng) đầu tiên dùng email đó
    col = 8; kk = '';
    for (let i = 0; i < rows.length; i++) if (rows[i][3] === k) { kk = rows[i][8]; break; }
  } else if (String(key).indexOf('|') >= 0) { col = 8; kk = String(key); }
  const byId = {}, order = [], mails = {}, depts = {}, got = {};
  let person = '';

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (!kk || r[col] !== kk) continue;
    if (!person) person = r[2];
    if (r[3]) mails[r[3]] = (mails[r[3]] || 0) + 1;
    if (r[5]) depts[r[5]] = (depts[r[5]] || 0) + 1;
    const id = idOf_(r);
    let e = byId[id];
    if (!e) { e = byId[id] = { date: r[0], event: r[1], idx: r[6], roles: [] }; order.push(id); }
    (r[7] || ['Đại biểu tham dự']).forEach(function (x) { if (e.roles.indexOf(x) < 0) e.roles.push(x); });
    if (r[6] >= 0) got[r[6]] = 1;
  }

  const rank = { 'Chủ tọa': 0, 'Báo cáo viên': 1, 'Đại biểu tham dự': 2 };
  const attended = order.map(function (id) {
    const e = byId[id];
    e.roles.sort(function (x, y) { return (rank[x] === undefined ? 3 : rank[x]) - (rank[y] === undefined ? 3 : rank[y]); });
    return { date: e.date, event: e.event, role: e.roles.join(', '), tiet: tietOf_(e.roles, e.idx >= 0 ? events[e.idx] : null) };
  }).sort(function (a, b) { return dateKey_(b.date) - dateKey_(a.date); });

  const today = todayKey_(), missing = [];
  events.forEach(function (ev, i) {
    if (!got[i]) missing.push({ date: ev.date, event: ev.name, upcoming: dateKey_(ev.date) > today });
  });

  const emails = Object.keys(mails).sort(function (a, b) { return mails[b] - mails[a]; })
    .slice(0, 3).map(maskEmail_);

  const dept = topDept_(depts);
  const tiet = attended.reduce(function (s, e) { return s + e.tiet; }, 0); // tự tính: số tiết từng buổi theo vai trò

  return {
    person: person, emails: emails, dept: dept,
    events: attended, missing: missing,
    count: Object.keys(got).length, total: events.length,
    tiet: tiet
  };
}

/* ------------------------------ KHOA / PHÒNG ------------------------------ */

// Mỗi người (họ tên + khoa/phòng): {key, name, dept, count (số buổi khớp DS), tiet (tổng số tiết)}
function people_() {
  const rows = loadRows_(), events = loadEvents_(), m = {};
  rows.forEach(function (r) {
    const p = m[r[8]] || (m[r[8]] = { name: r[2], depts: {}, ids: {}, ev: {} });
    if (r[5]) p.depts[r[5]] = (p.depts[r[5]] || 0) + 1;
    if (r[6] >= 0) p.ids[r[6]] = 1;
    const id = idOf_(r), e = p.ev[id] || (p.ev[id] = { idx: r[6], roles: [] });
    (r[7] || ['Đại biểu tham dự']).forEach(function (x) { if (e.roles.indexOf(x) < 0) e.roles.push(x); });
  });
  return Object.keys(m).map(function (k) {
    const p = m[k];
    const tiet = Object.keys(p.ev).reduce(function (s, id) {
      const e = p.ev[id];
      return s + tietOf_(e.roles, e.idx >= 0 ? events[e.idx] : null);
    }, 0);
    return { key: k, name: p.name, dept: topDept_(p.depts), count: Object.keys(p.ids).length, tiet: tiet };
  });
}

function depts_() {
  const people = people_(), counts = {};
  people.forEach(function (p) {
    if (!p.dept) return;
    const n = norm_(p.dept);
    counts[n] = (counts[n] || 0) + 1;
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
  return { depts: list, total: loadEvents_().length };
}

function dept_(name) {
  const n = norm_(name);
  const list = people_().filter(function (p) { return p.dept && norm_(p.dept) === n; })
    .sort(function (a, b) { return b.count - a.count || a.name.localeCompare(b.name, 'vi'); })
    .map(function (p) { return { key: p.key, name: p.name, count: p.count }; });
  return { dept: String(name || ''), people: list, total: loadEvents_().length };
}

/* ------------------------------ GHI LẠI SHEET THỐNG KÊ ------------------------------ */

function flushCache_() {
  const cache = CacheService.getScriptCache();
  ['rows', 'events'].forEach(function (name) {
    const n = +cache.get(CK + name + '_n') || 0, keys = [CK + name + '_n'];
    for (let i = 0; i < n; i++) keys.push(CK + name + '_' + i);
    cache.removeAll(keys);
  });
}

/**
 * Tính lại sheet "DS người tham dự theo khoa/phòng" từ "Tra cứu" + "DS các buổi event":
 *  - cùng họ tên nhưng khác khoa/phòng => 2 người riêng (cùng cách tách như trên web)
 *  - số buổi = số buổi khác nhau (khớp DS các buổi event), % = số buổi / tổng số buổi
 *  - số tiết = tổng số tiết từng buổi theo vai trò (Đại biểu / Chủ tọa / Báo cáo viên)
 * Chỉ ghi đè vùng dữ liệu A2:F của sheet đó (giữ dòng tiêu đề). Chạy tay từ menu "SHKH" hoặc trong editor.
 */
function rebuildPeopleSheet() {
  flushCache_();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const total = loadEvents_().length;
  const people = people_();
  const order = depts_().depts.map(function (d) { return d.name; });

  const byDept = {};
  people.forEach(function (p) { const k = norm_(p.dept) || 'khac'; (byDept[k] = byDept[k] || []).push(p); });
  const cmp = function (a, b) { return b.count - a.count || b.tiet - a.tiet || a.name.localeCompare(b.name, 'vi'); };

  const out = [], used = {};
  function add(deptName, list) {
    list.sort(cmp).forEach(function (p) { out.push(['=ROW()-1', deptName, p.name, p.count, total ? p.count / total : 0, p.tiet]); });
  }
  order.forEach(function (name) { const k = norm_(name); used[k] = 1; add(name, byDept[k] || []); });
  Object.keys(byDept).sort().forEach(function (k) { // khoa/phòng chưa có trong DS Khoa/phòng
    if (!used[k]) add(byDept[k][0].dept || 'Khác', byDept[k]);
  });

  let sh = ss.getSheetByName(PEOPLE_SHEET) || ss.insertSheet(PEOPLE_SHEET);
  if (!sh.getRange(1, 1).getValue()) {
    sh.getRange(1, 1, 1, 6).setValues([['STT', 'Khoa/phòng', 'Họ và tên', 'Số buổi tham dự', '% số buổi tham dự', 'Số tiết tích lũy']]);
    sh.setFrozenRows(1);
  }
  if (sh.getMaxRows() < out.length + 1) sh.insertRowsAfter(sh.getMaxRows(), out.length + 1 - sh.getMaxRows());
  sh.getRange(2, 1, sh.getMaxRows() - 1, 6).clearContent();
  if (out.length) {
    sh.getRange(2, 1, out.length, 6).setValues(out);
    sh.getRange(2, 4, out.length, 1).setNumberFormat('0"/' + total + '"');
    sh.getRange(2, 5, out.length, 1).setNumberFormat('0.0%');
  }
  Logger.log('Đã ghi ' + out.length + ' người (' + total + ' buổi) vào sheet "' + PEOPLE_SHEET + '".');
  return out.length;
}

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('SHKH')
      .addItem('Tính lại sheet "DS người tham dự theo khoa/phòng"', 'rebuildPeopleSheet')
      .addToUi();
  } catch (e) { /* script không gắn với sheet thì bỏ qua */ }
}

// Chạy tay trong editor để xem tên sự kiện nào ở sheet "Tra cứu" chưa khớp với "DS các buổi event"
function checkEventMatching() {
  const rows = loadRows_(), bad = {};
  rows.forEach(function (r) { if (r[6] < 0) bad[r[0] + ' | ' + r[1]] = (bad[r[0] + ' | ' + r[1]] || 0) + 1; });
  const keys = Object.keys(bad);
  Logger.log(keys.length ? 'Chưa khớp (' + keys.length + '):\n' + keys.map(function (k) { return k + '  x' + bad[k]; }).join('\n')
                         : 'Tất cả các dòng đều khớp với DS các buổi event.');
}
