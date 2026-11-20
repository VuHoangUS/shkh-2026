/**
 * TRA CỨU SINH HOẠT KHOA HỌC / HỘI THẢO KHOA HỌC 2026
 * Backend (Google Apps Script).
 *
 * Sheet "Tra cứu"          : A STT | B Ngày tổ chức | C Tên hội thảo | D Họ và tên | E Email | F Khoa/phòng
 * Sheet "DS các buổi event": A STT | B Ngày tổ chức | C Tên hội thảo | D Số tiết | E Số tiết cho BCV/chủ tọa
 *                              (danh sách đầy đủ các buổi -> mẫu số XX/30; dòng thiếu ngày như "Ghi chú" bị bỏ qua)
 * Sheet "DS Khoa/phòng"    : A STT | B Khoa/phòng
 * Sheet "DS người tham dự theo khoa/phòng": A STT | B Khoa/phòng | C Họ và tên | D Số buổi | E % | F Số tiết tích lũy
 *   (+ cột G Email, tùy chọn) -> web ĐỌC TRỰC TIẾP Số buổi / % / Số tiết tích lũy từ sheet này.
 *   Danh sách buổi đã/chưa tham dự, vai trò, số tiết từng buổi vẫn lấy từ "Tra cứu" + "DS các buổi event".
 *   Người không tìm thấy trong sheet này sẽ tự tính từ "Tra cứu".
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
const PEOPLE_SHEET   = 'DS người tham dự theo khoa/phòng'; // nguồn của Số buổi / % / Số tiết tích lũy (cập nhật theo thời gian thực)
const STATS_CACHE    = 60;     // sheet thống kê chỉ nhớ tạm 60 giây
const TIET_DELEGATE  = 3;      // mặc định (khi sheet buổi không ghi số tiết): Đại biểu tham dự
const TIET_SPEAKER   = 4;      // mặc định: Chủ tọa, Báo cáo viên
const MIN_CHARS      = 5;      // chỉ gợi ý khi gõ từ ký tự thứ 5
const MAX_SUGGEST    = 8;      // số gợi ý tối đa
const CACHE_SECONDS  = 60;     // dữ liệu nhớ tạm 60 giây (sửa sheet -> tối đa ~1 phút sau web mới cập nhật); cùng thời hạn với STATS_CACHE để các con số khớp nhau
const DEPT_PASSCODE  = '';     // để trống = ai cũng xem được danh sách theo khoa/phòng; đặt mã (vd 'abc123') để khóa
const CODE_VERSION   = '2026-11-20.2';
const CK             = 'v9_';  // tiền tố cache (đổi khi đổi cấu trúc dữ liệu)

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

function cachePut_(name, obj, ttl) { // chia mảnh <= 90KB (giới hạn 100KB/khóa)
  try {
    const json = JSON.stringify(obj), size = 90000, put = {};
    let n = 0;
    for (let i = 0; i < json.length; i += size) put[CK + name + '_' + n++] = json.slice(i, i + size);
    put[CK + name + '_n'] = String(n);
    CacheService.getScriptCache().putAll(put, ttl || CACHE_SECONDS);
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

/* ------------------------------ THỐNG KÊ (đọc từ sheet DS người tham dự theo khoa/phòng) ------------------------------ */

// Mảng: [khoa/phòng, họ tên, số buổi, % (0-100 hoặc null), số tiết, email(thường)]
function loadStats_() {
  const hit = cacheGet_('stats');
  if (hit) return hit;
  const out = [];
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(PEOPLE_SHEET);
  if (sheet && sheet.getLastRow() >= 2) {
    const data = sheet.getRange(1, 1, sheet.getLastRow(), Math.max(sheet.getLastColumn(), 6)).getValues();
    let cDept = 1, cName = 2, cCount = 3, cPct = 4, cTiet = 5, cMail = -1;
    data[0].forEach(function (x, i) {
      const t = norm_(x);
      if (t.indexOf('khoa') >= 0 || t.indexOf('phong') >= 0) cDept = i;
      else if (t.indexOf('ho va ten') >= 0 || t === 'ho ten') cName = i;
      else if (t.indexOf('%') >= 0) cPct = i;
      else if (t.indexOf('so buoi') >= 0) cCount = i;
      else if (t.indexOf('tiet') >= 0) cTiet = i;
      else if (t.indexOf('mail') >= 0) cMail = i;
    });
    for (let r = 1; r < data.length; r++) {
      const name = String(data[r][cName] || '').replace(/\s+/g, ' ').trim();
      if (!name) continue;
      const cv = data[r][cCount];
      let count = typeof cv === 'number' ? cv : (String(cv).match(/^\s*(\d+)/) || [])[1];
      count = count === undefined || count === '' ? null : +count;
      let pct = null;
      const pv = data[r][cPct];
      if (typeof pv === 'number') pct = pv <= 1 ? pv * 100 : pv;
      else if (String(pv).trim()) { const x = num_(String(pv).replace('%', '')); pct = x === null ? null : x; }
      out.push([
        String(data[r][cDept] || '').replace(/\s+/g, ' ').trim(),
        name,
        count,
        pct,
        num_(data[r][cTiet]),
        cMail >= 0 ? String(data[r][cMail] || '').trim().toLowerCase() : ''
      ]);
    }
  }
  cachePut_('stats', out, STATS_CACHE);
  return out;
}

// Tìm dòng thống kê của 1 người: cùng họ tên + cùng khoa/phòng (nhiều dòng thì ưu tiên trùng email);
// nếu không trùng khoa nhưng họ tên chỉ có đúng 1 dòng thì dùng dòng đó. Không có -> undefined.
function statFor_(nname, dept, emails) {
  const all = loadStats_().filter(function (x) { return norm_(x[1]) === nname; });
  if (!all.length) return undefined;
  const d = norm_(dept);
  let c = all.filter(function (x) { return norm_(x[0]) === d; });
  if (!c.length) c = all.length === 1 ? all : [];
  if (c.length > 1 && emails && emails.length) {
    const m = c.filter(function (x) { return x[5] && emails.indexOf(x[5]) >= 0; });
    if (m.length) c = m;
  }
  return c.length ? { count: c[0][2], pct: c[0][3], tiet: c[0][4] } : undefined;
}

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
      f.dept = topDept_(f.depts);
      const emails = f.type === 'name' ? Object.keys(f.mails) : [f.key];
      const st = statFor_(f.type === 'name' ? norm_(f.label) : norm_(f.sub), f.dept, emails);
      if (st && st.count !== null) f.count = st.count; // ưu tiên số buổi trong sheet thống kê
      if (f.type === 'name') { f.email = maskEmail_(topEmail_(f.mails)); delete f.mails; }
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
  } else if (String(key).indexOf('|') >= 0) { // "họ tên|khoa/phòng" (mã người từ gợi ý hoặc danh sách khoa/phòng)
    col = 8; kk = String(key);
    if (!rows.some(function (r) { return r[8] === kk; })) { // không có đúng mã -> thử theo họ tên + khoa/phòng, rồi họ tên
      const nm = kk.split('|')[0], dp = kk.split('|')[1];
      const hit = rows.filter(function (r) { return r[4] === nm && norm_(r[5]) === dp; })[0]
        || rows.filter(function (r) { return r[4] === nm; })[0];
      if (hit) kk = hit[8];
    }
  }
  const byId = {}, order = [], mails = {}, depts = {}, got = {};
  let person = '', nname = '';

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (!kk || r[col] !== kk) continue;
    if (!person) { person = r[2]; nname = r[4]; }
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
  const calcTiet = attended.reduce(function (s, e) { return s + e.tiet; }, 0); // số tiết tự tính (dự phòng)
  const calcCount = Object.keys(got).length;
  const listed = attended.length; // số buổi đang liệt kê bên dưới (gồm cả buổi chưa khớp DS)
  // Số buổi / % / số tiết: đọc trực tiếp từ sheet "DS người tham dự theo khoa/phòng"; không có thì dùng số tự tính
  const st = nname ? statFor_(nname, dept, Object.keys(mails)) : undefined;
  const count = st && st.count !== null ? st.count : calcCount;
  const pct = st && st.pct !== null ? st.pct : (events.length ? calcCount * 100 / events.length : 0);
  const tiet = st && st.tiet !== null ? st.tiet : calcTiet;

  return {
    person: person, emails: emails, dept: dept,
    events: attended, missing: missing,
    count: count, pct: pct, total: events.length,
    sync: count === calcCount && calcCount === listed, // false: số trong sheet thống kê lệch với danh sách buổi (vd. dữ liệu vừa sửa)
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
  const stats = loadStats_(), counts = {};
  if (stats.length) {
    stats.forEach(function (x) { const n = norm_(x[0]); if (n) counts[n] = (counts[n] || 0) + 1; });
  } else {
    people_().forEach(function (p) { const n = norm_(p.dept); if (n) counts[n] = (counts[n] || 0) + 1; });
  }

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

// Người của 1 khoa/phòng: đọc trực tiếp từ sheet "DS người tham dự theo khoa/phòng" (không có thì tự tính)
function dept_(name) {
  const n = norm_(name), total = loadEvents_().length;
  const stats = loadStats_();
  let list;
  if (stats.length) {
    list = stats.filter(function (x) { return norm_(x[0]) === n; }).map(function (x) {
      const count = x[2] === null ? 0 : x[2];
      return {
        key: norm_(x[1]) + '|' + n, name: x[1], count: count,
        pct: x[3] !== null ? x[3] : (total ? count * 100 / total : 0), tiet: x[4]
      };
    });
  } else {
    list = people_().filter(function (p) { return p.dept && norm_(p.dept) === n; }).map(function (p) {
      return { key: p.key, name: p.name, count: p.count, pct: total ? p.count * 100 / total : 0, tiet: p.tiet };
    });
  }
  list.sort(function (a, b) { return b.count - a.count || a.name.localeCompare(b.name, 'vi'); });
  return { dept: String(name || ''), people: list, total: total };
}

// Chạy tay trong editor để xem tên sự kiện nào ở sheet "Tra cứu" chưa khớp với "DS các buổi event"
function checkEventMatching() {
  const rows = loadRows_(), bad = {};
  rows.forEach(function (r) { if (r[6] < 0) bad[r[0] + ' | ' + r[1]] = (bad[r[0] + ' | ' + r[1]] || 0) + 1; });
  const keys = Object.keys(bad);
  Logger.log(keys.length ? 'Chưa khớp (' + keys.length + '):\n' + keys.map(function (k) { return k + '  x' + bad[k]; }).join('\n')
                         : 'Tất cả các dòng đều khớp với DS các buổi event.');
}
