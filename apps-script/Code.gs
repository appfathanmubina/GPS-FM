/**
 * GPS FM — Code.gs
 * Backend untuk Web App Apps Script. Otorisasi selalu dilakukan di server.
 */
const GPSFM_SESSION_TTL = 21600; // 6 jam
const GPSFM_STATUSES = ['Hadir','Izin','Sakit','Alpa','Tugas'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('GPS FM — Absensi Guru Pembina Santri')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function gpsfmHashPassword_(password, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(salt) + ':' + String(password), Utilities.Charset.UTF_8);
  return bytes.map(function(b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
}

function gpsfmDb_() {
  // Saat berjalan dari spreadsheet induk, gunakan spreadsheet yang terikat langsung.
  // Saat Web App berjalan tanpa konteks spreadsheet aktif, gunakan ID yang disimpan
  // oleh setupGpsFmDatabase() dari spreadsheet induk yang sama.
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    const id = active.getId();
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('GPSFM_SPREADSHEET_ID') !== id) {
      props.setProperty('GPSFM_SPREADSHEET_ID', id);
    }
    return active;
  }

  const id = PropertiesService.getScriptProperties().getProperty('GPSFM_SPREADSHEET_ID');
  if (!id) {
    throw new Error('Spreadsheet induk belum terhubung. Buka Apps Script melalui Ekstensi → Apps Script dari spreadsheet GPS FM, lalu jalankan setupGpsFmDatabase().');
  }
  return SpreadsheetApp.openById(id);
}

function gpsfmSheet_(name) {
  const sh = gpsfmDb_().getSheetByName(name);
  if (!sh) throw new Error('Sheet database tidak ditemukan: ' + name);
  return sh;
}

function gpsfmRows_(name) {
  const sh = gpsfmSheet_(name), values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  return values.slice(1).map(function(row, i) {
    const obj = { _row: i + 2 };
    headers.forEach(function(h, j) { obj[h] = row[j]; });
    return obj;
  }).filter(function(o) { return o[headers[0]] !== '' && o[headers[0]] != null; });
}

function gpsfmAppend_(name, obj) {
  const sh = gpsfmSheet_(name), headers = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
  sh.appendRow(headers.map(function(h) { return Object.prototype.hasOwnProperty.call(obj,h) ? obj[h] : ''; }));
}

function gpsfmUpdateRow_(name, rowNumber, obj) {
  const sh = gpsfmSheet_(name), headers = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
  const old = sh.getRange(rowNumber,1,1,headers.length).getValues()[0];
  headers.forEach(function(h,i) { if (Object.prototype.hasOwnProperty.call(obj,h)) old[i] = obj[h]; });
  sh.getRange(rowNumber,1,1,headers.length).setValues([old]);
}

function gpsfmId_() { return Utilities.getUuid(); }
function gpsfmNow_() { return new Date().toISOString(); }
function gpsfmDateKey_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Jakarta', 'yyyy-MM-dd');
  }
  const raw = String(value == null ? '' : value).trim();
  if (!raw) return '';
  // ISO dates/timestamps are the preferred stored format.
  const iso = raw.match(/^(\d{4}-\d{2}-\d{2})/);
  if (iso) return iso[1];
  // Also support legacy date strings such as M/D/YYYY or D/M/YYYY.
  const local = raw.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);
  if (local) {
    const a = Number(local[1]), b = Number(local[2]), year = local[3];
    const month = a > 12 ? b : (b > 12 ? a : b);
    const day = a > 12 ? a : (b > 12 ? b : a);
    return year + '-' + ('0' + month).slice(-2) + '-' + ('0' + day).slice(-2);
  }
  const parsed = new Date(raw);
  if (!isNaN(parsed.getTime())) return Utilities.formatDate(parsed, Session.getScriptTimeZone() || 'Asia/Jakarta', 'yyyy-MM-dd');
  return raw.slice(0, 10);
}

function gpsfmText_(value, max) {
  const s = String(value == null ? '' : value).trim();
  if (s.length > (max || 250)) throw new Error('Input terlalu panjang.');
  return s;
}
function gpsfmBool_(v) { return v === true || String(v).toLowerCase() === 'true'; }
function gpsfmPublicUser_(u) {
  return {userId:String(u.userId), username:String(u.username), fullName:String(u.fullName), role:String(u.role)};
}
function gpsfmAudit_(actorId, action, entityType, entityId, details) {
  gpsfmAppend_('AuditLogs', {logId:gpsfmId_(), actorId:actorId || '', action:action, entityType:entityType || '', entityId:entityId || '', details:JSON.stringify(details || {}), createdAt:gpsfmNow_()});
}
function gpsfmTokenKey_(token) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(token), Utilities.Charset.UTF_8);
  return 'gpsfm_session_' + bytes.map(function(b) { return ('0' + ((b+256)%256).toString(16)).slice(-2); }).join('');
}
function gpsfmSession_(token) {
  if (!token || String(token).length < 30) throw new Error('Sesi tidak valid. Silakan login kembali.');
  const raw = CacheService.getScriptCache().get(gpsfmTokenKey_(token));
  if (!raw) throw new Error('Sesi berakhir. Silakan login kembali.');
  const session = JSON.parse(raw);
  const user = gpsfmRows_('Users').find(function(u) { return String(u.userId) === session.userId && gpsfmBool_(u.active); });
  if (!user) throw new Error('Akun tidak aktif atau tidak ditemukan.');
  return {user:user, token:token};
}
function gpsfmRequireAdmin_(session) {
  if (String(session.user.role) !== 'superadmin') throw new Error('Akses ditolak. Fitur ini hanya untuk superadmin.');
}
function gpsfmIsPJ_(userId, activityId) {
  return gpsfmRows_('ActivityPJ').some(function(a) { return String(a.userId) === String(userId) && String(a.activityId) === String(activityId) && gpsfmBool_(a.active); });
}
function gpsfmIsActivityTeacher_(teacherId, activityId) {
  return gpsfmRows_('ActivityTeachers').some(function(a) {
    return String(a.activityId) === String(activityId) && String(a.teacherId) === String(teacherId) && gpsfmBool_(a.active);
  });
}
function gpsfmActivity_(id) {
  return gpsfmRows_('Activities').find(function(a) { return String(a.activityId) === String(id) && gpsfmBool_(a.active); });
}

function apiLogin(username, password) {
  username = gpsfmText_(username, 80).toLowerCase();
  password = String(password || '');
  if (!username || !password || password.length > 200) throw new Error('Username dan password wajib diisi.');
  const cache = CacheService.getScriptCache();
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, username, Utilities.Charset.UTF_8);
  const rateKey = 'gpsfm_login_fail_' + digest.map(function(b) { return ('0' + ((b+256)%256).toString(16)).slice(-2); }).join('');
  const failures = Number(cache.get(rateKey) || 0);
  if (failures >= 5) throw new Error('Terlalu banyak percobaan login. Coba lagi sekitar 15 menit.');
  const users = gpsfmRows_('Users');
  const user = users.find(function(u) { return String(u.username).toLowerCase() === username; });
  if (!user || !gpsfmBool_(user.active) || gpsfmHashPassword_(password, user.passwordSalt) !== String(user.passwordHash)) {
    const nextFailures = failures + 1;
    cache.put(rateKey, String(nextFailures), nextFailures >= 5 ? 900 : 600);
    Utilities.sleep(350);
    throw new Error('Username atau password tidak sesuai.');
  }
  cache.remove(rateKey);
  const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g,'');
  cache.put(gpsfmTokenKey_(token), JSON.stringify({userId:String(user.userId), issuedAt:Date.now()}), GPSFM_SESSION_TTL);
  gpsfmAudit_(user.userId, 'LOGIN', 'Users', user.userId, {});
  return {token:token, user:gpsfmPublicUser_(user), expiresIn:GPSFM_SESSION_TTL};
}

function apiLogout(token) {
  try {
    const s = gpsfmSession_(token);
    CacheService.getScriptCache().remove(gpsfmTokenKey_(token));
    gpsfmAudit_(s.user.userId, 'LOGOUT', 'Users', s.user.userId, {});
  } catch (e) {}
  return {ok:true};
}

function apiGetBootstrap(token) {
  const s = gpsfmSession_(token), u = s.user;
  const activities = gpsfmRows_('Activities').filter(function(a) {
    return gpsfmBool_(a.active) && (u.role === 'superadmin' || gpsfmIsPJ_(u.userId, a.activityId));
  }).map(function(a) {
    return {activityId:String(a.activityId), name:String(a.name), description:String(a.description || ''), schedule:String(a.schedule || ''), isPJ:gpsfmIsPJ_(u.userId,a.activityId)};
  });
  return {user:gpsfmPublicUser_(u), activities:activities, statuses:GPSFM_STATUSES, serverDate:Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd')};
}

function apiListUsers(token) {
  const s = gpsfmSession_(token); gpsfmRequireAdmin_(s);
  return gpsfmRows_('Users').map(function(u) { return {userId:String(u.userId), username:String(u.username), fullName:String(u.fullName), role:String(u.role), active:gpsfmBool_(u.active)}; });
}

function apiSaveUser(token, data) {
  const s = gpsfmSession_(token); gpsfmRequireAdmin_(s);
  data = data || {};
  const username = gpsfmText_(data.username, 50).toLowerCase();
  const fullName = gpsfmText_(data.fullName, 120);
  const role = String(data.role || 'gps');
  if (!/^[a-z0-9._-]{3,50}$/.test(username)) throw new Error('Username 3–50 karakter: huruf kecil, angka, titik, garis bawah, atau tanda hubung.');
  if (!fullName) throw new Error('Nama lengkap wajib diisi.');
  if (['gps','superadmin'].indexOf(role) < 0) throw new Error('Role tidak valid.');
  const users = gpsfmRows_('Users');
  if (users.some(function(u) { return String(u.username).toLowerCase() === username && String(u.userId) !== String(data.userId || ''); })) throw new Error('Username sudah digunakan.');
  const now = gpsfmNow_(), existing = users.find(function(u) { return String(u.userId) === String(data.userId || ''); });
  let record;
  if (existing) {
    record = {username:username,fullName:fullName,role:role,active:data.active !== false,updatedAt:now};
    gpsfmUpdateRow_('Users', existing._row, record);
  } else {
    const password = String(data.password || '');
    if (password.length < 6) throw new Error('Password awal minimal 6 karakter.');
    const salt = gpsfmId_().replace(/-/g,'');
    record = {userId:gpsfmId_(),username:username,fullName:fullName,role:role,passwordSalt:salt,passwordHash:gpsfmHashPassword_(password,salt),active:true,createdAt:now,updatedAt:now};
    gpsfmAppend_('Users', record);
  }
  gpsfmAudit_(s.user.userId, existing ? 'UPDATE_USER' : 'CREATE_USER', 'Users', existing ? existing.userId : record.userId, {username:username,role:role});
  return {ok:true};
}

function apiSetUserPassword(token, userId, newPassword) {
  const s = gpsfmSession_(token); gpsfmRequireAdmin_(s);
  if (String(newPassword || '').length < 6) throw new Error('Password minimal 6 karakter.');
  const u = gpsfmRows_('Users').find(function(x) { return String(x.userId) === String(userId); });
  if (!u) throw new Error('Akun tidak ditemukan.');
  const salt = gpsfmId_().replace(/-/g,'');
  gpsfmUpdateRow_('Users', u._row, {passwordSalt:salt,passwordHash:gpsfmHashPassword_(newPassword,salt),updatedAt:gpsfmNow_()});
  gpsfmAudit_(s.user.userId, 'RESET_PASSWORD', 'Users', userId, {});
  return {ok:true};
}

function apiListActivities(token) {
  const s = gpsfmSession_(token);
  return gpsfmRows_('Activities').filter(function(a) { return gpsfmBool_(a.active) && (s.user.role === 'superadmin' || gpsfmIsPJ_(s.user.userId,a.activityId)); }).map(function(a) {
    return {activityId:String(a.activityId),name:String(a.name),description:String(a.description || ''),schedule:String(a.schedule || ''),active:gpsfmBool_(a.active)};
  });
}

function apiSaveActivity(token, data) {
  const s = gpsfmSession_(token); gpsfmRequireAdmin_(s);
  data = data || {};
  const name = gpsfmText_(data.name,120);
  if (!name) throw new Error('Nama kegiatan wajib diisi.');
  const now = gpsfmNow_(), list = gpsfmRows_('Activities'), existing = list.find(function(a) { return String(a.activityId) === String(data.activityId || ''); });
  if (existing) gpsfmUpdateRow_('Activities',existing._row,{name:name,description:gpsfmText_(data.description,500),schedule:gpsfmText_(data.schedule,120),active:data.active !== false,updatedAt:now});
  else gpsfmAppend_('Activities',{activityId:gpsfmId_(),name:name,description:gpsfmText_(data.description,500),schedule:gpsfmText_(data.schedule,120),active:true,createdAt:now,updatedAt:now});
  gpsfmAudit_(s.user.userId,existing?'UPDATE_ACTIVITY':'CREATE_ACTIVITY','Activities',existing?existing.activityId:name,{name:name});
  return {ok:true};
}

function apiAssignPJ(token, activityId, userId, assigned) {
  const s = gpsfmSession_(token); gpsfmRequireAdmin_(s);
  const activity = gpsfmRows_('Activities').find(function(a) { return String(a.activityId) === String(activityId); });
  const user = gpsfmRows_('Users').find(function(u) { return String(u.userId) === String(userId) && gpsfmBool_(u.active) && String(u.role) === 'gps'; });
  if (!activity || !user) throw new Error('Kegiatan atau akun guru tidak valid.');
  const list = gpsfmRows_('ActivityPJ'), old = list.find(function(a) { return String(a.activityId) === String(activityId) && String(a.userId) === String(userId); });
  if (old) gpsfmUpdateRow_('ActivityPJ',old._row,{active:!!assigned});
  else if (assigned) gpsfmAppend_('ActivityPJ',{assignmentId:gpsfmId_(),activityId:activityId,userId:userId,active:true,assignedAt:gpsfmNow_(),assignedBy:s.user.userId});
  gpsfmAudit_(s.user.userId,assigned?'ASSIGN_PJ':'UNASSIGN_PJ','ActivityPJ',activityId,{userId:userId});
  return {ok:true};
}

function apiGetPJAssignments(token, activityId) {
  const s = gpsfmSession_(token); gpsfmRequireAdmin_(s);
  return gpsfmRows_('ActivityPJ').filter(function(a) { return String(a.activityId) === String(activityId) && gpsfmBool_(a.active); }).map(function(a) {
    const u = gpsfmRows_('Users').find(function(x) { return String(x.userId) === String(a.userId); });
    return {userId:String(a.userId),fullName:u?String(u.fullName):'Akun tidak ditemukan'};
  });
}

function apiGetActivityTeachers(token, activityId) {
  const s = gpsfmSession_(token);
  if (s.user.role !== 'superadmin' && !gpsfmIsPJ_(s.user.userId, activityId)) throw new Error('Anda bukan PJ untuk kegiatan ini.');
  if (!gpsfmActivity_(activityId)) throw new Error('Kegiatan tidak aktif atau tidak ditemukan.');
  const assigned = gpsfmRows_('ActivityTeachers').filter(function(a) { return String(a.activityId) === String(activityId) && gpsfmBool_(a.active); });
  const assignedIds = {};
  assigned.forEach(function(a) { assignedIds[String(a.teacherId)] = true; });
  // Keep the order from the Users sheet; do not alphabetically sort attendance teachers.
  return gpsfmRows_('Users').filter(function(u) {
    return assignedIds[String(u.userId)] && gpsfmBool_(u.active) && String(u.role) === 'gps';
  }).map(function(u) {
    return {userId:String(u.userId),fullName:String(u.fullName)};
  });
}

function apiAssignActivityTeacher(token, activityId, teacherId, assigned) {
  const s = gpsfmSession_(token); gpsfmRequireAdmin_(s);
  const activity = gpsfmRows_('Activities').find(function(a) { return String(a.activityId) === String(activityId); });
  const teacher = gpsfmRows_('Users').find(function(u) { return String(u.userId) === String(teacherId) && gpsfmBool_(u.active) && String(u.role) === 'gps'; });
  if (!activity || !teacher) throw new Error('Kegiatan atau akun guru tidak valid.');
  const list = gpsfmRows_('ActivityTeachers');
  const old = list.find(function(a) { return String(a.activityId) === String(activityId) && String(a.teacherId) === String(teacherId); });
  if (old) gpsfmUpdateRow_('ActivityTeachers',old._row,{active:!!assigned,updatedAt:gpsfmNow_()});
  else if (assigned) gpsfmAppend_('ActivityTeachers',{assignmentId:gpsfmId_(),activityId:activityId,teacherId:teacherId,active:true,assignedAt:gpsfmNow_(),assignedBy:s.user.userId,updatedAt:gpsfmNow_()});
  gpsfmAudit_(s.user.userId,assigned?'ASSIGN_ACTIVITY_TEACHER':'UNASSIGN_ACTIVITY_TEACHER','ActivityTeachers',activityId,{teacherId:teacherId});
  return {ok:true};
}

function apiGetAttendanceForDate(token, activityId, date) {
  const s = gpsfmSession_(token);
  date = gpsfmText_(date,10);
  if (s.user.role !== 'superadmin' && !gpsfmIsPJ_(s.user.userId,activityId)) throw new Error('Anda bukan PJ untuk kegiatan ini.');
  if (!gpsfmActivity_(activityId)) throw new Error('Kegiatan tidak aktif atau tidak ditemukan.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Tanggal tidak valid.');
  const allowed = apiGetActivityTeachers(token, activityId).map(function(t){return t.userId;});
  return gpsfmRows_('Attendance').filter(function(a) {
    return String(a.activityId) === String(activityId) && gpsfmDateKey_(a.attendanceDate) === date && allowed.indexOf(String(a.teacherId)) >= 0;
  }).map(function(a){return {teacherId:String(a.teacherId),status:String(a.status),note:String(a.note||''),recordedBy:String(a.recordedBy||'')};});
}

function apiSaveAttendanceBatch(token, data) {
  const s = gpsfmSession_(token), u = s.user;
  data = data || {};
  const activityId = gpsfmText_(data.activityId,80), date = gpsfmText_(data.attendanceDate,10);
  const entries = Array.isArray(data.entries) ? data.entries : [];
  if (u.role !== 'superadmin' && !gpsfmIsPJ_(u.userId,activityId)) throw new Error('Anda bukan PJ untuk kegiatan ini.');
  if (!gpsfmActivity_(activityId)) throw new Error('Kegiatan tidak aktif atau tidak ditemukan.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Tanggal tidak valid.');
  if (!entries.length || entries.length > 200) throw new Error('Pilih minimal satu guru dan maksimal 200 guru untuk disimpan.');
  const period = date.slice(0,7).split('-');
  const p = gpsfmRows_('Periods').find(function(x) { return Number(x.year) === Number(period[0]) && Number(x.month) === Number(period[1]); });
  if (p && String(p.status) === 'locked') throw new Error('Periode laporan sudah dikunci.');
  const seen = {}, users = gpsfmRows_('Users');
  entries.forEach(function(e) {
    const teacherId = gpsfmText_(e.teacherId,80), status = gpsfmText_(e.status,20);
    if (seen[teacherId]) throw new Error('Daftar guru duplikat. Muat ulang halaman dan coba lagi.');
    seen[teacherId] = true;
    if (GPSFM_STATUSES.indexOf(status) < 0) throw new Error('Status kehadiran tidak valid.');
    if (!users.some(function(x){return String(x.userId)===teacherId && gpsfmBool_(x.active) && String(x.role)==='gps';})) throw new Error('Salah satu akun guru tidak aktif atau tidak ditemukan.');
    if (!gpsfmIsActivityTeacher_(teacherId,activityId)) throw new Error('Ada guru yang tidak ditugaskan pada kegiatan ini. Muat ulang daftar guru.');
    gpsfmText_(e.note,300);
  });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const records = gpsfmRows_('Attendance');
    const existing = entries.map(function(e) {
      return {entry:e,old:records.find(function(a){return String(a.activityId)===activityId && String(a.teacherId)===String(e.teacherId) && gpsfmDateKey_(a.attendanceDate)===date;})};
    });
    if (u.role !== 'superadmin' && existing.some(function(x){return x.old && String(x.old.recordedBy)!==String(u.userId);})) throw new Error('Sebagian absensi sudah dicatat PJ lain. Hubungi superadmin untuk koreksi; tidak ada data yang disimpan.');
    const now = gpsfmNow_();
    existing.forEach(function(x) {
      const e=x.entry, teacherId=String(e.teacherId), status=String(e.status), note=gpsfmText_(e.note,300);
      if (x.old) {
        gpsfmUpdateRow_('Attendance',x.old._row,{status:status,note:note,recordedBy:u.userId,updatedAt:now});
        gpsfmAudit_(u.userId,'UPDATE_ATTENDANCE','Attendance',x.old.attendanceId,{status:status,date:date,bulk:true});
      } else {
        const id=gpsfmId_();
        gpsfmAppend_('Attendance',{attendanceId:id,activityId:activityId,teacherId:teacherId,attendanceDate:date,status:status,note:note,recordedBy:u.userId,createdAt:now,updatedAt:now});
        gpsfmAudit_(u.userId,'CREATE_ATTENDANCE','Attendance',id,{status:status,date:date,bulk:true});
      }
    });
  } finally { lock.releaseLock(); }
  return {ok:true,saved:entries.length};
}

function apiSaveAttendance(token, data) {
  const s = gpsfmSession_(token), u = s.user;
  data = data || {};
  const activityId = gpsfmText_(data.activityId,80), teacherId = gpsfmText_(data.teacherId,80);
  const date = gpsfmText_(data.attendanceDate,10), status = gpsfmText_(data.status,20);
  if (u.role !== 'superadmin' && !gpsfmIsPJ_(u.userId,activityId)) throw new Error('Anda bukan PJ untuk kegiatan ini.');
  if (!gpsfmActivity_(activityId)) throw new Error('Kegiatan tidak aktif atau tidak ditemukan.');
  const teacher = gpsfmRows_('Users').find(function(x) { return String(x.userId) === teacherId && gpsfmBool_(x.active) && String(x.role) === 'gps'; });
  if (!teacher) throw new Error('Guru GPS tidak ditemukan atau tidak aktif.');
  if (!gpsfmIsActivityTeacher_(teacherId,activityId)) throw new Error('Guru tersebut belum ditetapkan sebagai guru yang terlibat dalam kegiatan ini.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Tanggal tidak valid.');
  if (GPSFM_STATUSES.indexOf(status) < 0) throw new Error('Status kehadiran tidak valid.');
  const period = date.slice(0,7).split('-'), p = gpsfmRows_('Periods').find(function(x) { return Number(x.year) === Number(period[0]) && Number(x.month) === Number(period[1]); });
  if (p && String(p.status) === 'locked') throw new Error('Periode laporan sudah dikunci.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const records = gpsfmRows_('Attendance');
    const old = records.find(function(a) { return String(a.activityId) === activityId && String(a.teacherId) === teacherId && gpsfmDateKey_(a.attendanceDate) === date; });
    const now = gpsfmNow_();
    if (old) {
      if (u.role !== 'superadmin' && String(old.recordedBy) !== String(u.userId)) throw new Error('Absensi ini sudah dicatat PJ lain; hubungi superadmin untuk koreksi.');
      gpsfmUpdateRow_('Attendance',old._row,{status:status,note:gpsfmText_(data.note,300),recordedBy:u.userId,updatedAt:now});
      gpsfmAudit_(u.userId,'UPDATE_ATTENDANCE','Attendance',old.attendanceId,{status:status,date:date});
    } else {
      const id = gpsfmId_();
      gpsfmAppend_('Attendance',{attendanceId:id,activityId:activityId,teacherId:teacherId,attendanceDate:date,status:status,note:gpsfmText_(data.note,300),recordedBy:u.userId,createdAt:now,updatedAt:now});
      gpsfmAudit_(u.userId,'CREATE_ATTENDANCE','Attendance',id,{status:status,date:date});
    }
  } finally { lock.releaseLock(); }
  return {ok:true};
}

function apiGetMonthlyReport(token, year, month) {
  const s = gpsfmSession_(token);
  year = Number(year); month = Number(month);
  if (!Number.isInteger(year) || year < 2020 || year > 2100 || !Number.isInteger(month) || month < 1 || month > 12) throw new Error('Periode laporan tidak valid.');
  const prefix = year + '-' + ('0'+month).slice(-2);
  const users = gpsfmRows_('Users').filter(function(u) { return String(u.role) === 'gps'; });
  const allActivities = gpsfmRows_('Activities').filter(function(a) { return gpsfmBool_(a.active); });
  const userMap = {}, actMap = {};
  users.forEach(function(u) { userMap[String(u.userId)] = String(u.fullName); });
  allActivities.forEach(function(a) { actMap[String(a.activityId)] = String(a.name); });

  // Guru hanya memperoleh pilihan kegiatan yang memang menugaskannya.
  // Superadmin tetap dapat melihat seluruh kegiatan aktif.
  let allowedActivityIds;
  if (s.user.role === 'superadmin') {
    allowedActivityIds = allActivities.map(function(a) { return String(a.activityId); });
  } else {
    allowedActivityIds = gpsfmRows_('ActivityTeachers').filter(function(a) {
      return String(a.teacherId) === String(s.user.userId) && gpsfmBool_(a.active);
    }).map(function(a) { return String(a.activityId); }).filter(function(id) {
      return Object.prototype.hasOwnProperty.call(actMap, id);
    });
  }
  const activities = allActivities.filter(function(a) {
    return allowedActivityIds.indexOf(String(a.activityId)) >= 0;
  }).map(function(a) {
    return {activityId:String(a.activityId), activityName:String(a.name)};
  }).sort(function(a,b) { return a.activityName.localeCompare(b.activityName); });

  const rows = gpsfmRows_('Attendance').filter(function(a) {
    const date = gpsfmDateKey_(a.attendanceDate);
    return date.slice(0,7) === prefix && allowedActivityIds.indexOf(String(a.activityId)) >= 0 && (s.user.role === 'superadmin' || String(a.teacherId) === String(s.user.userId));
  });
  const result = rows.map(function(a) {
    const date = gpsfmDateKey_(a.attendanceDate);
    return {date:date,teacherId:String(a.teacherId),teacherName:userMap[String(a.teacherId)] || 'Guru tidak ditemukan',activityId:String(a.activityId),activityName:actMap[String(a.activityId)] || 'Kegiatan tidak ditemukan',status:String(a.status),note:String(a.note || '')};
  }).sort(function(a,b) { return a.date.localeCompare(b.date) || a.activityName.localeCompare(b.activityName) || a.teacherName.localeCompare(b.teacherName); });
  const summary = {Hadir:0,Izin:0,Sakit:0,Alpa:0,Tugas:0};
  result.forEach(function(r) { summary[r.status] = (summary[r.status] || 0) + 1; });
  return {year:year,month:month,summary:summary,rows:result,activities:activities};
}

function apiGetTeachers(token) {
  const s = gpsfmSession_(token);
  if (s.user.role !== 'superadmin' && s.user.role !== 'gps') throw new Error('Akses ditolak.');
  return gpsfmRows_('Users').filter(function(u) { return gpsfmBool_(u.active) && String(u.role) === 'gps'; }).map(function(u) { return {userId:String(u.userId),fullName:String(u.fullName)}; });
}

function apiSetPeriodLock(token, year, month, locked) {
  const s = gpsfmSession_(token); gpsfmRequireAdmin_(s);
  year=Number(year); month=Number(month);
  if (!Number.isInteger(year) || year < 2020 || year > 2100 || !Number.isInteger(month) || month<1 || month>12) throw new Error('Periode tidak valid.');
  const list = gpsfmRows_('Periods'), old = list.find(function(p) { return Number(p.year)===year && Number(p.month)===month; });
  const record = {periodId:old?old.periodId:gpsfmId_(),year:year,month:month,status:locked?'locked':'open',lockedAt:locked?gpsfmNow_():'',lockedBy:locked?s.user.userId:''};
  if (old) gpsfmUpdateRow_('Periods',old._row,record); else gpsfmAppend_('Periods',record);
  gpsfmAudit_(s.user.userId,locked?'LOCK_PERIOD':'UNLOCK_PERIOD','Periods',record.periodId,{year:year,month:month});
  return {ok:true};
}

function apiChangeOwnPassword(token, oldPassword, newPassword) {
  const s = gpsfmSession_(token);
  if (String(newPassword || '').length < 6) throw new Error('Password baru minimal 6 karakter.');
  if (gpsfmHashPassword_(String(oldPassword || ''),s.user.passwordSalt) !== String(s.user.passwordHash)) throw new Error('Password lama tidak sesuai.');
  const salt = gpsfmId_().replace(/-/g,'');
  gpsfmUpdateRow_('Users',s.user._row,{passwordSalt:salt,passwordHash:gpsfmHashPassword_(newPassword,salt),updatedAt:gpsfmNow_()});
  gpsfmAudit_(s.user.userId,'CHANGE_PASSWORD','Users',s.user.userId,{});
  return {ok:true};
}
