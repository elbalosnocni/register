/**
 * ĐĂNG KÝ NGƯỜI THÂN CỔ VŨ - HỘI THAO LIWAYWAY 2026
 * Triển khai Web App: Execute as Me; Who has access: Anyone.
 * Chạy setupAdminPassword() trong Apps Script một lần để khởi tạo mật khẩu Admin mặc định.
 */
const SPREADSHEET_ID = '1g3Lc4RCiE23EzjLT1yshjUHX5GsYOJWgmJJn7ultesA';
const SHEET_ATHLETES = 'DanhSachVDV';
const SHEET_REGISTRATIONS = 'DangKyNguoiThan';
const PROP_ADMIN_HASH = 'ADMIN_PASSWORD_HASH';
const PROP_DEADLINE = 'REGISTRATION_DEADLINE';
const DEFAULT_ADMIN_PASSWORD = 'Hoadaquy'; // Đổi ngay sau lần đăng nhập đầu tiên.
const MAX_RELATIVES = 30;

/** Chạy thủ công một lần trong Apps Script Editor để khởi tạo/đặt lại Admin. */
function setupAdminPassword() {
  PropertiesService.getScriptProperties().setProperty(PROP_ADMIN_HASH, hashPassword_(DEFAULT_ADMIN_PASSWORD));
  Logger.log('Đã đặt mật khẩu Admin mặc định. Đăng nhập bằng Hoadaquy và đổi mật khẩu ngay.');
}

function doGet(e) {
  e = e || { parameter: {} };
  const p = e.parameter || {};
  let result;
  try {
    const action = String(p.action || 'lookup');
    if (action === 'lookup') result = lookupAthlete_(p.cccd);
    else if (action === 'listRelatives') result = listRelatives_(p.cccd);
    else if (action === 'register') result = registerRelatives_(JSON.parse(String(p.payload || '{}')));
    else if (action === 'adminLogin') result = adminLogin_(String(p.password || ''));
    else if (action === 'updateConfig') result = updateConfig_(p);
    else result = { success: false, message: 'Yêu cầu không hợp lệ.' };
  } catch (err) {
    result = { success: false, status: 'error', message: safeError_(err) };
  }
  return output_(result, p.callback);
}

// Giữ khả năng tương thích với các client gửi POST JSON. Client GitHub Pages dùng JSONP GET
// để đọc được phản hồi xuyên miền và chỉ báo thành công khi backend trả success=true.
function doPost(e) {
  let result;
  try {
    const body = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
    if (body.action === 'updateConfig') result = updateConfig_(body);
    else result = registerRelatives_(body);
  } catch (err) {
    result = { success: false, status: 'error', message: safeError_(err) };
  }
  return output_(result, e && e.parameter && e.parameter.callback);
}

function lookupAthlete_(cccd) {
  cccd = normalizeCccd_(cccd);
  if (!/^\d{12}$/.test(cccd)) return { found: false, message: 'CCCD phải gồm đúng 12 chữ số.' };
  const sheet = getSheet_(SHEET_ATHLETES, ['CCCD', 'Họ tên VĐV', 'Bộ phận']);
  const rows = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < rows.length; i++) {
    if (normalizeCccd_(rows[i][0]) === cccd) {
      return { found: true, name: String(rows[i][1] || '').trim(), dept: String(rows[i][2] || '').trim() };
    }
  }
  return { found: false, message: 'Không tìm thấy CCCD trong danh sách VĐV.' };
}

/** Trả về danh sách người thân theo CCCD VĐV đã xác thực; CCCD người thân luôn được che, chỉ hiện 3 số cuối. */
function listRelatives_(cccd) {
  cccd = normalizeCccd_(cccd);
  if (!/^\d{12}$/.test(cccd)) return { success: false, message: 'CCCD phải gồm đúng 12 chữ số.' };
  const athlete = lookupAthlete_(cccd);
  if (!athlete.found) return { success: false, message: 'Không tìm thấy CCCD trong danh sách VĐV.' };
  const headers = ['Thời gian cập nhật', 'CCCD VĐV', 'Họ tên VĐV', 'Bộ phận', 'Họ tên người thân', 'Quan hệ', 'Số điện thoại người thân', 'CCCD người thân', 'Đồng ý điều khoản', 'Thời điểm đồng ý'];
  const sheet = getSheet_(SHEET_REGISTRATIONS, headers);
  const data = sheet.getDataRange().getDisplayValues();
  const relatives = [];
  let updatedAt = '';
  for (let i = 1; i < data.length; i++) {
    if (normalizeCccd_(data[i][1]) !== cccd || !String(data[i][4] || '').trim()) continue;
    const rawCccd = normalizeCccd_(data[i][7]);
    relatives.push({
      name: String(data[i][4] || '').trim(),
      relation: String(data[i][5] || '').trim(),
      phone: String(data[i][6] || '').trim().replace(/^'/, ''),
      cccdMasked: rawCccd.length >= 3 ? '•••••••••' + rawCccd.slice(-3) : '••••••••••••'
    });
    updatedAt = String(data[i][0] || updatedAt);
  }
  return { success: true, athlete: { name: athlete.name, dept: athlete.dept, cccdMasked: '•••••••••' + cccd.slice(-3) }, relatives: relatives, count: relatives.length, updatedAt: updatedAt };
}

function registerRelatives_(data) {
  if (!data || typeof data !== 'object') throw new Error('Dữ liệu đăng ký không hợp lệ.');
  if (data.consent !== true) throw new Error('Bạn cần đồng ý điều khoản bảo mật trước khi gửi đăng ký.');
  const cccd = normalizeCccd_(data.cccd);
  if (!/^\d{12}$/.test(cccd)) throw new Error('CCCD phải gồm đúng 12 chữ số.');
  const athlete = lookupAthlete_(cccd);
  if (!athlete.found) throw new Error('CCCD không tồn tại trong danh sách VĐV.');
  const relatives = Array.isArray(data.relatives) ? data.relatives : [];
  if (relatives.length < 1 || relatives.length > MAX_RELATIVES) throw new Error('Vui lòng đăng ký từ 1 đến ' + MAX_RELATIVES + ' người thân.');
  const allowedRelations = ['Bố/Mẹ', 'Vợ/Chồng', 'Con', 'Anh/Chị/Em', 'Khác'];
  const clean = relatives.map(function(r) {
    const name = cleanText_(r && r.name, 120);
    const relation = cleanText_(r && r.relation, 40);
    const phone = normalizePhone_(r && r.phone);
    const relativeCccd = normalizeCccd_(r && r.cccd);
    if (!name) throw new Error('Họ tên người thân không được để trống.');
    if (allowedRelations.indexOf(relation) < 0) throw new Error('Quan hệ người thân không hợp lệ.');
    if (!/^0\d{9,10}$/.test(phone)) throw new Error('Số điện thoại người thân phải gồm 10 hoặc 11 chữ số và bắt đầu bằng số 0.');
    if (!/^\d{12}$/.test(relativeCccd)) throw new Error('CCCD người thân phải gồm đúng 12 chữ số.');
    return { name: name, relation: relation, phone: phone, cccd: relativeCccd };
  });
  // Không cho trùng CCCD người thân trong cùng danh sách.
  const seen = {};
  clean.forEach(function(r) {
    if (seen[r.cccd]) throw new Error('CCCD người thân bị trùng trong danh sách: ' + r.cccd + '.');
    seen[r.cccd] = true;
  });

  const props = PropertiesService.getScriptProperties();
  const deadline = props.getProperty(PROP_DEADLINE);
  if (deadline && new Date().getTime() > new Date(deadline).getTime()) throw new Error('Đã hết hạn đăng ký người thân.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const headers = ['Thời gian cập nhật', 'CCCD VĐV', 'Họ tên VĐV', 'Bộ phận', 'Họ tên người thân', 'Quan hệ', 'Số điện thoại người thân', 'CCCD người thân', 'Đồng ý điều khoản', 'Thời điểm đồng ý'];
    const sheet = getSheet_(SHEET_REGISTRATIONS, headers);
    ensureHeaders_(sheet, headers);
    const values = sheet.getDataRange().getDisplayValues();
    // Ghi danh sách mới trước để tránh mất danh sách cũ nếu thao tác ghi thất bại.
    // Sau khi ghi thành công mới xóa các dòng cũ của CCCD này, từ dưới lên.
    const oldRows = [];
    for (let i = 1; i < values.length; i++) {
      if (normalizeCccd_(values[i][1]) === cccd) oldRows.push(i + 1);
    }
    const timestamp = new Date();
    const rows = clean.map(function(r) { return [timestamp, "'" + cccd, athlete.name, athlete.dept, r.name, r.relation, "'" + r.phone, "'" + r.cccd, 'Có', timestamp]; });
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    for (let j = oldRows.length - 1; j >= 0; j--) sheet.deleteRow(oldRows[j]);
    return { success: true, status: 'success', message: 'Đã cập nhật danh sách người thân.', count: rows.length };
  } finally {
    lock.releaseLock();
  }
}

function adminLogin_(password) {
  const props = PropertiesService.getScriptProperties();
  let storedHash = props.getProperty(PROP_ADMIN_HASH);
  if (!storedHash) {
    // Không tự động mở quyền quản trị: yêu cầu chủ dự án chạy setupAdminPassword() trong Editor.
    return { success: false, message: 'Admin chưa được khởi tạo. Chủ dự án hãy chạy setupAdminPassword() trong Apps Script Editor.' };
  }
  if (hashPassword_(password) !== storedHash) return { success: false, message: 'Mật khẩu quản trị không đúng.' };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const reg = getSheet_(SHEET_REGISTRATIONS, ['Thời gian cập nhật', 'CCCD VĐV', 'Họ tên VĐV', 'Bộ phận', 'Họ tên người thân', 'Quan hệ', 'Số điện thoại người thân', 'CCCD người thân', 'Đồng ý điều khoản', 'Thời điểm đồng ý']);
  const data = reg.getDataRange().getDisplayValues();
  const deptStats = {};
  const athletes = {};
  let totalRelatives = 0;
  for (let i = 1; i < data.length; i++) {
    const cccd = normalizeCccd_(data[i][1]);
    const name = String(data[i][4] || '').trim();
    if (!cccd || !name) continue;
    totalRelatives++;
    athletes[cccd] = true;
    const dept = String(data[i][3] || 'Chưa phân bộ phận').trim() || 'Chưa phân bộ phận';
    deptStats[dept] = (deptStats[dept] || 0) + 1;
  }
  const deadline = props.getProperty(PROP_DEADLINE) || '';
  return { success: true, deadline: toDatetimeLocal_(deadline), stats: { totalRelatives: totalRelatives, totalAthletes: Object.keys(athletes).length, deptStats: deptStats } };
}

function updateConfig_(p) {
  const props = PropertiesService.getScriptProperties();
  const supplied = String(p.adminPassword || '');
  const storedHash = props.getProperty(PROP_ADMIN_HASH);
  if (!storedHash || hashPassword_(supplied) !== storedHash) return { success: false, message: 'Không có quyền cập nhật cấu hình.' };
  const deadline = String(p.newDeadline || '').trim();
  let deadlineIso = '';
  if (deadline) {
    const parsed = new Date(deadline);
    if (isNaN(parsed.getTime())) return { success: false, message: 'Hạn đăng ký không hợp lệ.' };
    deadlineIso = parsed.toISOString();
  }
  const newPassword = String(p.newPassword || '');
  if (newPassword && newPassword.length < 8) return { success: false, message: 'Mật khẩu mới phải có ít nhất 8 ký tự.' };
  if (deadlineIso) props.setProperty(PROP_DEADLINE, deadlineIso); else props.deleteProperty(PROP_DEADLINE);
  if (newPassword) props.setProperty(PROP_ADMIN_HASH, hashPassword_(newPassword));
  return { success: true, message: 'Đã lưu cấu hình.' };
}

function getSheet_(name, headers) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0 && headers && headers.length) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (headers && headers.length) ensureHeaders_(sheet, headers);
  return sheet;
}
function ensureHeaders_(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }
  const width = Math.max(sheet.getLastColumn(), headers.length);
  const current = sheet.getRange(1, 1, 1, width).getDisplayValues()[0];
  headers.forEach(function(header, i) {
    if (!String(current[i] || '').trim()) sheet.getRange(1, i + 1).setValue(header);
  });
}
function normalizePhone_(value) { return String(value == null ? '' : value).replace(/[\s().-]/g, '').trim(); }

function normalizeCccd_(value) { return String(value == null ? '' : value).replace(/^'/, '').replace(/\s+/g, '').trim(); }
function cleanText_(value, max) { return String(value == null ? '' : value).replace(/[<>\u0000-\u001F]/g, '').trim().slice(0, max); }
function hashPassword_(password) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(password), Utilities.Charset.UTF_8);
  return bytes.map(function(b) { const v = (b + 256) % 256; return ('0' + v.toString(16)).slice(-2); }).join('');
}
function toDatetimeLocal_(value) {
  if (!value) return '';
  const d = new Date(value);
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm");
}
function output_(obj, callback) {
  const json = JSON.stringify(obj);
  if (callback && /^[A-Za-z_$][\w.$]{0,100}$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + json + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
function safeError_(err) { return String(err && err.message ? err.message : err || 'Lỗi không xác định').slice(0, 300); }
