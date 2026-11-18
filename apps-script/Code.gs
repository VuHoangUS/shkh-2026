/**
 * TRA CỨU SINH HOẠT KHOA HỌC / HỘI THẢO KHOA HỌC 2026
 * Backend (Google Apps Script) – đọc sheet "Tra cứu".
 *
 * Cột: A STT | B Ngày tổ chức | C Tên hội thảo | D Họ và tên | E Email
 * (Mã tự dò theo tiêu đề dòng 1; nếu không dò được sẽ dùng đúng thứ tự A–E.)
 *
 * API (GET):
 *   ?action=suggest&q=...            -> gợi ý (tối thiểu 5 ký tự)
 *   ?action=search&type=name&key=... -> danh sách sự kiện của 1 người (theo họ tên không dấu)
 *   ?action=search&type=email&key=.. -> danh sách sự kiện theo email
 */

const SPREADSHEET_ID = '16MUn8yW0lqvdiGRc98q5OYKmiCHx3LoBLgozwyj9YHc';
const SHEET_NAME     = 'Tra cứu';
const MIN_CHARS      = 5;      // chỉ gợi ý khi gõ từ ký tự thứ 5
const MAX_SUGGEST    = 8;      // số gợi ý tối đa
const CACHE_SECONDS  = 600;    // dữ liệu được nhớ tạm 10 phút (sửa sheet -> tối đa 10 phút sau web mới cập nhật)
const CODE_VERSION   = '2026-11-18.2';

/* ------------------------------ ENTRY ------------------------------ */

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === 'suggest') return output(suggest_(p.q));
    if (p.action === 'search')  return output(search_(p.type, p.key));
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

// "email@gmail.com" -> "Em**l@gmail.com" (giữ 2 ký tự đầu + 1 ký tự cuối phần tên)
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

// email xuất hiện nhiều nhất trong {email: số lần}
function topEmail_(counts) {
  let best = '', n = 0;
  for (const k in counts) if (counts[k] > n) { best = k; n = counts[k]; }
  return best;
}

/* ------------------------------ DATA ------------------------------ */

function fmtDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy');
  return String(v == null ? '' : v).trim();
}

function dateKey_(s) { // dd/mm/yyyy -> yyyymmdd (để sắp xếp)
  const m = String(s).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  return m ? (+m[3]) * 10000 + (+m[2]) * 100 + (+m[1]) : 0;
}

function findCol_(header, patterns, fallback) {
  for (let i = 0; i < header.length; i++) {
    const h = norm_(header[i]);
    if (patterns.some(function (p) { return h.indexOf(p) >= 0; })) return i;
  }
  return fallback;
}

// Trả về mảng gọn: [ngày, tên sự kiện, họ tên, email(thường), họ tên không dấu]
function loadRows_() {
  const cache = CacheService.getScriptCache();
  const meta = cache.get('rows_n');
  if (meta) {
    const keys = [];
    for (let i = 0; i < +meta; i++) keys.push('rows_' + i);
    const parts = cache.getAll(keys);
    if (keys.every(function (k) { return parts[k]; })) {
      return JSON.parse(keys.map(function (k) { return parts[k]; }).join(''));
    }
  }

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const width = Math.max(sheet.getLastColumn(), 5);
  const data = sheet.getRange(1, 1, last, width).getValues();

  const h = data[0];
  const cDate  = findCol_(h, ['ngay'], 1);
  const cEvent = findCol_(h, ['ten hoi thao', 'hoi thao', 'sinh hoat'], 2);
  const cName  = findCol_(h, ['ho va ten', 'ho ten'], 3);
  const cMail  = findCol_(h, ['email', 'mail'], 4);

  const rows = [];
  for (let r = 1; r < data.length; r++) {
    const name = String(data[r][cName] || '').replace(/\s+/g, ' ').trim();
    const event = String(data[r][cEvent] || '').replace(/\s+/g, ' ').trim();
    if (!name || !event) continue;
    rows.push([
      fmtDate_(data[r][cDate]),
      event,
      name,
      String(data[r][cMail] || '').trim().toLowerCase(),
      norm_(name)
    ]);
  }

  try { // cache theo từng mảnh <= 90KB (giới hạn 100KB/khóa)
    const json = JSON.stringify(rows), size = 90000, put = {};
    let n = 0;
    for (let i = 0; i < json.length; i += size) put['rows_' + n++] = json.slice(i, i + size);
    put['rows_n'] = String(n);
    cache.putAll(put, CACHE_SECONDS);
  } catch (err) { /* không cache được vẫn chạy bình thường */ }

  return rows;
}

/* ------------------------------ SUGGEST ------------------------------ */

function suggest_(q) {
  const raw = String(q || '').trim().toLowerCase();
  const nq = norm_(q);
  if (nq.replace(/\s/g, '').length < MIN_CHARS) return { suggestions: [], minChars: MIN_CHARS };

  const rows = loadRows_();
  const isEmail = raw.indexOf('@') >= 0;
  const found = {}; // key -> {label, sub, type, key, count, rank}

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (isEmail) {
      const mail = r[3];
      if (!mail || mail.indexOf(raw) < 0) continue;
      const f = found[mail] || (found[mail] = {
        type: 'email', key: mail, label: maskEmail_(mail), sub: r[2],
        count: 0, rank: mail.indexOf(raw) === 0 ? 0 : 1
      });
      f.count++;
    } else {
      const pos = r[4].indexOf(nq);
      if (pos < 0) continue;
      const f = found[r[4]] || (found[r[4]] = {
        type: 'name', key: r[4], label: r[2], sub: '', mails: {},
        count: 0, rank: pos === 0 ? 0 : (r[4].indexOf(' ' + nq) >= 0 ? 1 : 2)
      });
      f.count++;
      if (r[3]) f.mails[r[3]] = (f.mails[r[3]] || 0) + 1;
    }
  }

  const list = Object.keys(found).map(function (k) { return found[k]; })
    .sort(function (a, b) { return a.rank - b.rank || b.count - a.count || a.label.localeCompare(b.label); })
    .slice(0, MAX_SUGGEST)
    .map(function (f) {
      if (f.type === 'name') { f.email = maskEmail_(topEmail_(f.mails)); delete f.mails; }
      return f;
    });

  return { suggestions: list, minChars: MIN_CHARS };
}

/* ------------------------------ SEARCH ------------------------------ */

function search_(type, key) {
  const rows = loadRows_();
  const k = type === 'email' ? String(key || '').trim().toLowerCase() : norm_(key);
  if (!k) return { person: '', events: [] };

  const col = type === 'email' ? 3 : 4;
  const events = [], seen = {};
  let person = '';
  const mails = {};

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (r[col] !== k) continue;
    if (!person) person = r[2];
    if (r[3]) mails[r[3]] = (mails[r[3]] || 0) + 1;
    const id = r[0] + '|' + r[1];
    if (seen[id]) continue; // bỏ trùng cùng ngày + cùng sự kiện
    seen[id] = 1;
    events.push({ date: r[0], event: r[1] });
  }

  events.sort(function (a, b) { return dateKey_(b.date) - dateKey_(a.date); });
  const emails = Object.keys(mails).sort(function (a, b) { return mails[b] - mails[a]; })
    .slice(0, 3).map(maskEmail_);
  return { person: person, emails: emails, events: events };
}
