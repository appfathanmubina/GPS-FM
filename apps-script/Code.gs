/**
 * GPS FM — Code.gs
 * Backend untuk Web App Apps Script. Otorisasi selalu dilakukan di server.
 */
const GPSFM_SESSION_TTL = 21600; // 6 jam
const GPSFM_STATUSES = ['Hadir','A. Sakit','B. Tugas Pondok','C. Izin Pribadi','D. Alfa'];

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
  const sh = gpsfmSheet_(name), lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(String);
  return values.slice(1).map(function(row, i) {
    const obj = { _row: i + 2 };
    headers.forEach(function(h, j) { obj[h] = row[j]; });
    return obj;
  }).filter(function(o) { return o[headers[0]] !== '' && o[headers[0]] != null; });
}

/**
 * Cari baris berdasarkan satu kolom kunci tanpa membaca seluruh tabel.
 * Mengembalikan objek lengkap agar pemanggil tetap menerima bentuk data lama.
 */
function gpsfmFindRowByValue_(sheetName, keyName, keyValue) {
  const sh = gpsfmSheet_(sheetName), lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return null;
  const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(String);
  const keyIndex = headers.indexOf(keyName);
  if (keyIndex < 0) throw new Error('Kolom database tidak ditemukan: ' + keyName + ' (' + sheetName + ').');
  const matches = sh.getRange(2, keyIndex + 1, lastRow - 1, 1).createTextFinder(String(keyValue))
    .matchEntireCell(true).matchCase(true).findAll();
  if (!matches.length) return null;
  const rowNumber = matches[0].getRow(), row = sh.getRange(rowNumber, 1, 1, lastCol).getValues()[0];
  const obj = { _row: rowNumber };
  headers.forEach(function(h, i) { obj[h] = row[i]; });
  return obj;
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
  return {userId:String(u.userId), username:String(u.username), fullName:String(u.fullName), role:String(u.role), mustChangePassword:gpsfmBool_(u.mustChangePassword)};
}
function gpsfmEnsurePasswordFlag_() {
  const sh=gpsfmSheet_('Users'),lastCol=sh.getLastColumn(),headers=sh.getRange(1,1,1,lastCol).getDisplayValues()[0];
  if(headers.indexOf('mustChangePassword')<0){
    sh.getRange(1,lastCol+1).setValue('mustChangePassword');
    const lastRow=sh.getLastRow();
    if(lastRow>1){
      const rows=sh.getRange(2,1,lastRow-1,lastCol).getValues();
      const roleIndex=headers.indexOf('role');
      const flags=rows.map(function(row){return [String(row[roleIndex])==='superadmin'?false:true];});
      sh.getRange(2,lastCol+1,flags.length,1).setValues(flags);
    }
  }
}
function gpsfmAudit_(actorId, action, entityType, entityId, details) {
  gpsfmAppend_('AuditLogs', {logId:gpsfmId_(), actorId:actorId || '', action:action, entityType:entityType || '', entityId:entityId || '', details:JSON.stringify(details || {}), createdAt:gpsfmNow_()});
}
function gpsfmTokenKey_(token) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(token), Utilities.Charset.UTF_8);
  return 'gpsfm_session_' + bytes.map(function(b) { return ('0' + ((b+256)%256).toString(16)).slice(-2); }).join('');
}
function gpsfmSession_(token, allowForcedPasswordChange) {
  if (!token || String(token).length < 30) throw new Error('Sesi tidak valid. Silakan login kembali.');
  const raw = CacheService.getScriptCache().get(gpsfmTokenKey_(token));
  if (!raw) throw new Error('Sesi berakhir. Silakan login kembali.');
  const session = JSON.parse(raw);
  const user = gpsfmFindRowByValue_('Users', 'userId', session.userId);
  if (user && !gpsfmBool_(user.active)) throw new Error('Akun tidak aktif atau tidak ditemukan.');
  if (!user) throw new Error('Akun tidak aktif atau tidak ditemukan.');
  if(gpsfmBool_(user.mustChangePassword)&&!allowForcedPasswordChange)throw new Error('Anda wajib mengganti password sebelum menggunakan aplikasi.');
  return {user:user, token:token};
}
function gpsfmRequireAdmin_(session) {
  if (String(session.user.role) !== 'superadmin') throw new Error('Akses ditolak. Fitur ini hanya untuk superadmin.');
}
function gpsfmIsPJ_(userId, activityId) {
  const sh = gpsfmSheet_('ActivityPJ'), lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return false;
  const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(String);
  const userCol = headers.indexOf('userId') + 1, activityCol = headers.indexOf('activityId') + 1, activeCol = headers.indexOf('active') + 1;
  if (!userCol || !activityCol || !activeCol) throw new Error('Kolom penugasan PJ tidak lengkap.');
  const matches = sh.getRange(2, userCol, lastRow - 1, 1).createTextFinder(String(userId))
    .matchEntireCell(true).matchCase(true).findAll();
  if (!matches.length) return false;
  // Baca kolom aktivitas dan status aktif masing-masing satu kali, bukan per baris hasil pencarian.
  const activityValues = sh.getRange(2, activityCol, lastRow - 1, 1).getValues();
  const activeValues = sh.getRange(2, activeCol, lastRow - 1, 1).getValues();
  for (let i = 0; i < matches.length; i++) {
    const index = matches[i].getRow() - 2;
    if (String(activityValues[index][0]) === String(activityId) && gpsfmBool_(activeValues[index][0])) return true;
  }
  return false;
}
function gpsfmIsActivityTeacher_(teacherId, activityId) {
  return gpsfmRows_('ActivityTeachers').some(function(a) {
    return String(a.activityId) === String(activityId) && String(a.teacherId) === String(teacherId) && gpsfmBool_(a.active);
  });
}
function gpsfmIsParticipantRole_(role) {
  return ['gps','supervisor','admin_administrasi'].indexOf(String(role)) >= 0;
}
function gpsfmActivity_(id) {
  const activity = gpsfmFindRowByValue_('Activities', 'activityId', id);
  return activity && gpsfmBool_(activity.active) ? activity : null;
}

function gpsfmFindUserByUsername_(username) {
  const sh = gpsfmSheet_('Users'), lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return null;
  const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(String);
  const keyIndex = headers.indexOf('username');
  if (keyIndex < 0) throw new Error('Kolom username tidak ditemukan pada database Users.');
  const matches = sh.getRange(2, keyIndex + 1, lastRow - 1, 1)
    .createTextFinder(String(username)).matchEntireCell(true).matchCase(false).findAll();
  if (!matches.length) return null;
  const rowNumber = matches[0].getRow();
  const values = sh.getRange(rowNumber, 1, 1, lastCol).getValues()[0];
  const user = {_row:rowNumber};
  headers.forEach(function(header, index) { user[header] = values[index]; });
  return user;
}

function apiLogin(username, password) {
  gpsfmEnsurePasswordFlag_();
  username = gpsfmText_(username, 80).toLowerCase();
  password = String(password || '');
  if (!username || !password || password.length > 200) throw new Error('Username dan password wajib diisi.');
  const cache = CacheService.getScriptCache();
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, username, Utilities.Charset.UTF_8);
  const rateKey = 'gpsfm_login_fail_' + digest.map(function(b) { return ('0' + ((b+256)%256).toString(16)).slice(-2); }).join('');
  const failures = Number(cache.get(rateKey) || 0);
  if (failures >= 5) throw new Error('Terlalu banyak percobaan login. Coba lagi sekitar 15 menit.');
  // Cari satu akun secara langsung; tidak perlu membaca seluruh tabel Users.
  const user = gpsfmFindUserByUsername_(username);
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
  // Revoke token directly so logout also works during mandatory password change
  // or when the account has just been disabled. Never affect other tab tokens.
  if (!token || String(token).length < 30) return {ok:true};
  const cache = CacheService.getScriptCache();
  const key = gpsfmTokenKey_(token);
  let session = null;
  try {
    const raw = cache.get(key);
    if (raw) session = JSON.parse(raw);
  } catch (e) {}
  // Revoke first; audit logging must never prevent the session from ending.
  cache.remove(key);
  try {
    if (session && session.userId) {
      const user = gpsfmFindRowByValue_('Users', 'userId', session.userId);
      if (user) gpsfmAudit_(user.userId, 'LOGOUT', 'Users', user.userId, {});
    }
  } catch (e) {}
  return {ok:true};
}

function apiGetBootstrap(token) {
  const s = gpsfmSession_(token, true), u = s.user;
  const canSeeAllActivities = ['superadmin','supervisor','admin_administrasi'].indexOf(String(u.role)) >= 0;
  const pjActivityIds = {};
  if (!canSeeAllActivities) {
    const assignments = gpsfmRows_('ActivityPJ');
    assignments.forEach(function(item) {
      if (String(item.userId) === String(u.userId) && gpsfmBool_(item.active)) pjActivityIds[String(item.activityId)] = true;
    });
  }
  const activities = gpsfmRows_('Activities').filter(function(a) {
    return gpsfmBool_(a.active) && (canSeeAllActivities || !!pjActivityIds[String(a.activityId)]);
  }).map(function(a) {
    return {activityId:String(a.activityId), name:String(a.name), description:String(a.description || ''), schedule:String(a.schedule || ''), isPJ:!!pjActivityIds[String(a.activityId)]};
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
  if (['gps','supervisor','admin_administrasi','superadmin'].indexOf(role) < 0) throw new Error('Role tidak valid.');
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
    record = {userId:gpsfmId_(),username:username,fullName:fullName,role:role,passwordSalt:salt,passwordHash:gpsfmHashPassword_(password,salt),mustChangePassword:true,active:true,createdAt:now,updatedAt:now};
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
  gpsfmUpdateRow_('Users', u._row, {passwordSalt:salt,passwordHash:gpsfmHashPassword_(newPassword,salt),mustChangePassword:true,updatedAt:gpsfmNow_()});
  gpsfmAudit_(s.user.userId, 'RESET_PASSWORD', 'Users', userId, {});
  return {ok:true};
}

function apiListActivities(token) {
  const s = gpsfmSession_(token);
  const activities = gpsfmRows_('Activities').filter(function(a) { return gpsfmBool_(a.active); });
  if (s.user.role === 'superadmin') return activities.map(function(a) {
    return {activityId:String(a.activityId),name:String(a.name),description:String(a.description || ''),schedule:String(a.schedule || ''),active:gpsfmBool_(a.active)};
  });
  const assigned = {};
  gpsfmRows_('ActivityPJ').forEach(function(p) {
    if (String(p.userId) === String(s.user.userId) && gpsfmBool_(p.active)) assigned[String(p.activityId)] = true;
  });
  return activities.filter(function(a) { return !!assigned[String(a.activityId)]; }).map(function(a) {
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
  const user = gpsfmRows_('Users').find(function(u) { return String(u.userId) === String(userId) && gpsfmBool_(u.active) && gpsfmIsParticipantRole_(u.role); });
  if (!activity || !user) throw new Error('Kegiatan atau akun peserta tidak valid.');
  const list = gpsfmRows_('ActivityPJ'), old = list.find(function(a) { return String(a.activityId) === String(activityId) && String(a.userId) === String(userId); });
  if (old) gpsfmUpdateRow_('ActivityPJ',old._row,{active:!!assigned});
  else if (assigned) gpsfmAppend_('ActivityPJ',{assignmentId:gpsfmId_(),activityId:activityId,userId:userId,active:true,assignedAt:gpsfmNow_(),assignedBy:s.user.userId});
  gpsfmAudit_(s.user.userId,assigned?'ASSIGN_PJ':'UNASSIGN_PJ','ActivityPJ',activityId,{userId:userId});
  return {ok:true};
}

function apiGetPJAssignments(token, activityId) {
  const s = gpsfmSession_(token); gpsfmRequireAdmin_(s);
  const usersById = {};
  gpsfmRows_('Users').forEach(function(u) { usersById[String(u.userId)] = u; });
  return gpsfmRows_('ActivityPJ').filter(function(a) { return String(a.activityId) === String(activityId) && gpsfmBool_(a.active); }).map(function(a) {
    const u = usersById[String(a.userId)];
    return {userId:String(a.userId),fullName:u?String(u.fullName):'Akun tidak ditemukan'};
  });
}

function apiGetActivityTeachers(token, activityId) {
  const s = gpsfmSession_(token);
  if (s.user.role !== 'superadmin' && s.user.role !== 'supervisor' && s.user.role !== 'admin_administrasi' && !gpsfmIsPJ_(s.user.userId, activityId)) throw new Error('Anda tidak memiliki akses laporan kegiatan ini.');
  if (!gpsfmActivity_(activityId)) throw new Error('Kegiatan tidak aktif atau tidak ditemukan.');
  const assigned = gpsfmRows_('ActivityTeachers').filter(function(a) { return String(a.activityId) === String(activityId) && gpsfmBool_(a.active); });
  const assignedIds = {};
  assigned.forEach(function(a) { assignedIds[String(a.teacherId)] = true; });
  // Keep the order from the Users sheet; do not alphabetically sort attendance teachers.
  return gpsfmRows_('Users').filter(function(u) {
    return assignedIds[String(u.userId)] && gpsfmBool_(u.active) && gpsfmIsParticipantRole_(u.role);
  }).map(function(u) {
    return {userId:String(u.userId),fullName:String(u.fullName)};
  });
}

function apiAssignActivityTeacher(token, activityId, teacherId, assigned) {
  const s = gpsfmSession_(token); gpsfmRequireAdmin_(s);
  const activity = gpsfmRows_('Activities').find(function(a) { return String(a.activityId) === String(activityId); });
  const teacher = gpsfmRows_('Users').find(function(u) { return String(u.userId) === String(teacherId) && gpsfmBool_(u.active) && gpsfmIsParticipantRole_(u.role); });
  if (!activity || !teacher) throw new Error('Kegiatan atau akun peserta tidak valid.');
  const list = gpsfmRows_('ActivityTeachers');
  const old = list.find(function(a) { return String(a.activityId) === String(activityId) && String(a.teacherId) === String(teacherId); });
  // Bekukan roster occurrence lampau sebelum perubahan penugasan guru.
  gpsfmSnapshotPastOccurrences_(activityId, gpsfmActiveTeacherIds_(activityId));
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
  const allowed = {};
  apiGetActivityTeachers(token, activityId).forEach(function(t) { allowed[String(t.userId)] = true; });
  if (!Object.keys(allowed).length) return [];
  const sh = gpsfmSheet_('Attendance'), lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(String), cols = {};
  headers.forEach(function(h, i) { cols[h] = i; });
  if (!Object.prototype.hasOwnProperty.call(cols, 'activityId') || !Object.prototype.hasOwnProperty.call(cols, 'attendanceDate') ||
      !Object.prototype.hasOwnProperty.call(cols, 'teacherId') || !Object.prototype.hasOwnProperty.call(cols, 'status')) {
    throw new Error('Kolom Attendance tidak lengkap.');
  }
  const result = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (String(row[cols.activityId]) !== String(activityId) || gpsfmDateKey_(row[cols.attendanceDate]) !== date) continue;
    const teacherId = String(row[cols.teacherId]);
    if (allowed[teacherId]) result.push({
      teacherId:teacherId,
      status:String(row[cols.status]),
      note:String(Object.prototype.hasOwnProperty.call(cols,'note') ? row[cols.note] || '' : ''),
      recordedBy:String(Object.prototype.hasOwnProperty.call(cols,'recordedBy') ? row[cols.recordedBy] || '' : '')
    });
  }
  return result;
}

function gpsfmValidatePJAttendanceMonth_(user, date) {
  // Superadmin tetap mengikuti mekanisme penguncian periode yang sudah ada.
  if (String(user.role) === 'superadmin') return;

  // PJ boleh mengisi bulan berjalan, serta bulan sebelumnya sampai tanggal 2
  // pada bulan berjalan (berdasarkan kalender Asia/Jakarta, bukan 48 jam).
  const todayKey = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
  const currentMonth = todayKey.slice(0, 7);
  const currentDay = Number(todayKey.slice(8, 10));
  const parts = currentMonth.split('-');
  const previousMonth = Utilities.formatDate(
    new Date(Date.UTC(Number(parts[0]), Number(parts[1]) - 2, 1)),
    'UTC',
    'yyyy-MM'
  );
  const attendanceMonth = String(date).slice(0, 7);

  if (attendanceMonth === currentMonth) return;
  if (attendanceMonth === previousMonth && currentDay <= 2) return;

  throw new Error('PJ dapat menginput absensi untuk bulan berjalan dan bulan sebelumnya hanya sampai tanggal 2. Setelah itu, hubungi superadmin untuk periode lain.');
}

function apiSaveAttendanceBatch(token, data) {
  const s = gpsfmSession_(token), u = s.user;
  data = data || {};
  const activityId = gpsfmText_(data.activityId,80), date = gpsfmText_(data.attendanceDate,10);
  const entries = Array.isArray(data.entries) ? data.entries : [];
  if (u.role !== 'superadmin' && !gpsfmIsPJ_(u.userId,activityId)) throw new Error('Anda bukan PJ untuk kegiatan ini.');
  if (!gpsfmActivity_(activityId)) throw new Error('Kegiatan tidak aktif atau tidak ditemukan.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Tanggal tidak valid.');
  gpsfmValidatePJAttendanceMonth_(u, date);
  if (!entries.length || entries.length > 200) throw new Error('Pilih minimal satu guru dan maksimal 200 guru untuk disimpan.');

  const period = date.slice(0,7).split('-');
  const p = gpsfmRows_('Periods').find(function(x) { return Number(x.year) === Number(period[0]) && Number(x.month) === Number(period[1]); });
  if (p && String(p.status) === 'locked') throw new Error('Periode laporan sudah dikunci.');

  // Baca master sekali, lalu validasi seluruh entri di memori.
  const users = gpsfmRows_('Users');
  const activeUsers = {};
  users.forEach(function(x) {
    if (gpsfmBool_(x.active) && gpsfmIsParticipantRole_(x.role)) activeUsers[String(x.userId)] = true;
  });
  const activityTeachers = {};
  gpsfmRows_('ActivityTeachers').forEach(function(x) {
    if (String(x.activityId) === activityId && gpsfmBool_(x.active)) activityTeachers[String(x.teacherId)] = true;
  });

  const seen = {}, normalized = [];
  entries.forEach(function(e) {
    const teacherId = gpsfmText_(e.teacherId,80), status = gpsfmText_(e.status,20);
    if (!teacherId || seen[teacherId]) throw new Error('Daftar guru duplikat atau tidak valid. Muat ulang halaman dan coba lagi.');
    seen[teacherId] = true;
    if (GPSFM_STATUSES.indexOf(status) < 0) throw new Error('Status kehadiran tidak valid.');
    if (!activeUsers[teacherId]) throw new Error('Salah satu akun guru tidak aktif atau tidak ditemukan.');
    if (!activityTeachers[teacherId]) throw new Error('Ada guru yang tidak ditugaskan pada kegiatan ini. Muat ulang daftar guru.');
    normalized.push({teacherId:teacherId,status:status,note:gpsfmText_(e.note,300)});
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const attendanceSheet = gpsfmSheet_('Attendance');
    const attendanceValues = attendanceSheet.getDataRange().getValues();
    const attendanceHeaders = attendanceValues.length ? attendanceValues[0].map(String) : [];
    const attendanceCols = {};
    attendanceHeaders.forEach(function(h,i) { attendanceCols[h] = i; });
    ['attendanceId','activityId','teacherId','attendanceDate','status','note','recordedBy','createdAt','updatedAt'].forEach(function(h) {
      if (!Object.prototype.hasOwnProperty.call(attendanceCols,h)) throw new Error('Kolom Attendance tidak lengkap: ' + h);
    });

    const existingByTeacher = {};
    for (let i=1; i<attendanceValues.length; i++) {
      const row = attendanceValues[i];
      if (String(row[attendanceCols.activityId]) === activityId &&
          gpsfmDateKey_(row[attendanceCols.attendanceDate]) === date) {
        const teacherKey = String(row[attendanceCols.teacherId]);
        if (!Object.prototype.hasOwnProperty.call(existingByTeacher,teacherKey)) existingByTeacher[teacherKey] = {rowNumber:i+1,values:row};
      }
    }

    const now = gpsfmNow_(), updateRows = [], insertRows = [], auditRows = [];
    const auditSheet = gpsfmSheet_('AuditLogs');
    const auditHeaders = auditSheet.getRange(1,1,1,auditSheet.getLastColumn()).getDisplayValues()[0];
    const auditCols = {};
    auditHeaders.forEach(function(h,i) { auditCols[h] = i; });

    function makeAudit(action, entityId, status) {
      const obj = {logId:gpsfmId_(),actorId:u.userId,action:action,entityType:'Attendance',entityId:entityId,details:JSON.stringify({status:status,date:date,bulk:true}),createdAt:now};
      auditRows.push(auditHeaders.map(function(h) { return Object.prototype.hasOwnProperty.call(obj,h) ? obj[h] : ''; }));
    }

    normalized.forEach(function(e) {
      const old = existingByTeacher[e.teacherId];
      if (old) {
        const row = old.values.slice();
        row[attendanceCols.status] = e.status;
        row[attendanceCols.note] = e.note;
        row[attendanceCols.recordedBy] = u.userId;
        row[attendanceCols.updatedAt] = now;
        updateRows.push({rowNumber:old.rowNumber,values:row});
        makeAudit('UPDATE_ATTENDANCE',String(row[attendanceCols.attendanceId]),e.status);
      } else {
        const id = gpsfmId_();
        const obj = {attendanceId:id,activityId:activityId,teacherId:e.teacherId,attendanceDate:date,status:e.status,note:e.note,recordedBy:u.userId,createdAt:now,updatedAt:now};
        insertRows.push(attendanceHeaders.map(function(h) { return Object.prototype.hasOwnProperty.call(obj,h) ? obj[h] : ''; }));
        makeAudit('CREATE_ATTENDANCE',id,e.status);
      }
    });

    // Tulis baris update yang berurutan sebagai satu operasi per rentang.
    updateRows.sort(function(a,b) { return a.rowNumber-b.rowNumber; });
    let runStart = 0;
    while (runStart < updateRows.length) {
      let runEnd = runStart + 1;
      while (runEnd < updateRows.length && updateRows[runEnd].rowNumber === updateRows[runEnd-1].rowNumber + 1) runEnd++;
      const chunk = updateRows.slice(runStart,runEnd);
      attendanceSheet.getRange(chunk[0].rowNumber,1,chunk.length,attendanceHeaders.length)
        .setValues(chunk.map(function(x) { return x.values; }));
      runStart = runEnd;
    }

    // Semua data baru ditambahkan sekaligus.
    if (insertRows.length) {
      attendanceSheet.getRange(attendanceSheet.getLastRow()+1,1,insertRows.length,attendanceHeaders.length).setValues(insertRows);
    }
    // Semua log audit ditambahkan sekaligus tanpa menghilangkan jejak per guru.
    if (auditRows.length) {
      auditSheet.getRange(auditSheet.getLastRow()+1,1,auditRows.length,auditHeaders.length).setValues(auditRows);
    }
  } finally {
    lock.releaseLock();
  }
  return {ok:true,saved:normalized.length};
}

function apiSaveAttendance(token, data) {
  const s = gpsfmSession_(token), u = s.user;
  data = data || {};
  const activityId = gpsfmText_(data.activityId,80), teacherId = gpsfmText_(data.teacherId,80);
  const date = gpsfmText_(data.attendanceDate,10), status = gpsfmText_(data.status,20);
  if (u.role !== 'superadmin' && !gpsfmIsPJ_(u.userId,activityId)) throw new Error('Anda bukan PJ untuk kegiatan ini.');
  if (!gpsfmActivity_(activityId)) throw new Error('Kegiatan tidak aktif atau tidak ditemukan.');
  const teacher = gpsfmRows_('Users').find(function(x) { return String(x.userId) === teacherId && gpsfmBool_(x.active) && gpsfmIsParticipantRole_(x.role); });
  if (!teacher) throw new Error('Peserta kegiatan tidak ditemukan atau tidak aktif.');
  if (!gpsfmIsActivityTeacher_(teacherId,activityId)) throw new Error('Guru tersebut belum ditetapkan sebagai guru yang terlibat dalam kegiatan ini.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Tanggal tidak valid.');
  gpsfmValidatePJAttendanceMonth_(u, date);
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


/* Target absensi mingguan dan pengecualian pertemuan. Semua pemeriksaan akses di server. */
function gpsfmParseScheduleDays_(schedule) {
  const raw=String(schedule||'').toLowerCase().replace(/[–—]/g,'-'), days=['Senin','Selasa','Rabu','Kamis','Jumat','Sabtu','Minggu'];
  if(raw.indexOf('setiap hari')>=0)return days;
  if(raw.indexOf('senin-jumat')>=0)return days.slice(0,5);
  return days.filter(function(day){return new RegExp('(^|[^a-z])'+day.toLowerCase()+'([^a-z]|$)').test(raw);});
}
function gpsfmDateFromKey_(key){const p=String(key).split('-').map(Number);return new Date(Date.UTC(p[0],p[1]-1,p[2],12,0,0));}
function gpsfmDateKeyUTC_(d){return d.getUTCFullYear()+'-'+('0'+(d.getUTCMonth()+1)).slice(-2)+'-'+('0'+d.getUTCDate()).slice(-2);}
function gpsfmWeekdayName_(key){return ['Minggu','Senin','Selasa','Rabu','Kamis','Jumat','Sabtu'][gpsfmDateFromKey_(key).getUTCDay()];}
function gpsfmEnsureOccurrenceSheet_(){
 const ss=gpsfmDb_();let sh=ss.getSheetByName('ActivityOccurrences');
 if(!sh){
  sh=ss.insertSheet('ActivityOccurrences');
  sh.getRange(1,1,1,9).setValues([['occurrenceId','activityId','occurrenceDate','status','reason','teacherIdsJson','updatedBy','updatedAt','createdAt']]);
  sh.setFrozenRows(1);
 }
 return sh;
}
function gpsfmGetOccurrence_(activityId,date){
 gpsfmEnsureOccurrenceSheet_();
 return gpsfmRows_('ActivityOccurrences').find(function(o){return String(o.activityId)===String(activityId)&&gpsfmDateKey_(o.occurrenceDate)===String(date);});
}
function gpsfmSaveOccurrence_(activityId,date,status,reason,actorId,teacherIds){
 const lock=LockService.getScriptLock();lock.waitLock(10000);
 try{
  const old=gpsfmGetOccurrence_(activityId,date),now=gpsfmNow_(),values={activityId:String(activityId),occurrenceDate:String(date),status:status,reason:reason||'',teacherIdsJson:JSON.stringify(teacherIds||[]),updatedBy:actorId||'',updatedAt:now};
  if(old)gpsfmUpdateRow_('ActivityOccurrences',old._row,values);
  else gpsfmAppend_('ActivityOccurrences',{occurrenceId:gpsfmId_(),activityId:String(activityId),occurrenceDate:String(date),status:status,reason:reason||'',teacherIdsJson:JSON.stringify(teacherIds||[]),updatedBy:actorId||'',updatedAt:now,createdAt:now});
 } finally {lock.releaseLock();}
 return gpsfmGetOccurrence_(activityId,date);
}
function gpsfmActiveTeacherIds_(activityId){
 const active={};gpsfmRows_('ActivityTeachers').filter(function(a){return String(a.activityId)===String(activityId)&&gpsfmBool_(a.active);}).forEach(function(a){active[String(a.teacherId)]=true;});
 return gpsfmRows_('Users').filter(function(u){return active[String(u.userId)]&&gpsfmBool_(u.active)&&gpsfmIsParticipantRole_(u.role);}).map(function(u){return String(u.userId);});
}
function gpsfmSnapshotPastOccurrences_(activityId,teacherIds){
  const a=gpsfmRows_('Activities').find(function(x){return String(x.activityId)===String(activityId);});
  if(!a)return;
  const days=gpsfmParseScheduleDays_(a.schedule);
  if(!days.length)return;
  const today=Utilities.formatDate(new Date(),'Asia/Jakarta','yyyy-MM-dd');
  const startRaw=gpsfmDateKey_(a.createdAt)||today;
  const end=gpsfmDateFromKey_(today),start=gpsfmDateFromKey_(startRaw),earliest=new Date(end.getTime());
  earliest.setUTCDate(earliest.getUTCDate()-730);
  const cursor=start<earliest?earliest:start;
  cursor.setUTCHours(12,0,0,0);
  gpsfmEnsureOccurrenceSheet_();
  const lock=LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh=gpsfmSheet_('ActivityOccurrences');
    const lastRow=sh.getLastRow(),lastCol=sh.getLastColumn();
    const headers=sh.getRange(1,1,1,lastCol).getDisplayValues()[0].map(String),cols={};
    headers.forEach(function(h,i){cols[h]=i;});
    const existing={};
    if(lastRow>1){
      sh.getRange(2,1,lastRow-1,lastCol).getValues().forEach(function(row){
        existing[String(row[cols.activityId])+'|'+gpsfmDateKey_(row[cols.occurrenceDate])]=true;
      });
    }
    const now=gpsfmNow_(),pending=[];
    while(cursor<end){
      const date=gpsfmDateKeyUTC_(cursor);
      if(days.indexOf(gpsfmWeekdayName_(date))>=0){
        const key=String(activityId)+'|'+date;
        if(!existing[key]){
          const obj={occurrenceId:gpsfmId_(),activityId:String(activityId),occurrenceDate:date,status:'scheduled',reason:'',teacherIdsJson:JSON.stringify(teacherIds||[]),updatedBy:'',updatedAt:now,createdAt:now};
          pending.push(headers.map(function(h){return Object.prototype.hasOwnProperty.call(obj,h)?obj[h]:'';}));
          existing[key]=true;
        }
      }
      cursor.setUTCDate(cursor.getUTCDate()+1);
    }
    if(pending.length)sh.getRange(sh.getLastRow()+1,1,pending.length,headers.length).setValues(pending);
  } finally { lock.releaseLock(); }
}

function gpsfmOccurrenceTeacherIds_(o,id){
 if(o&&o.teacherIdsJson){try{const ids=JSON.parse(String(o.teacherIdsJson));if(Array.isArray(ids))return ids.map(String);}catch(e){}}
 return gpsfmActiveTeacherIds_(id);
}
function gpsfmScheduleStartMinutes_(schedule){
  const match=String(schedule||'').match(/(?:^|[^0-9])([0-9]{1,2})[.:]([0-9]{2})(?![0-9])/);
  if(!match)return null;
  const hour=Number(match[1]),minute=Number(match[2]);
  if(hour>23||minute>59)return null;
  return hour*60+minute;
}
function gpsfmScheduledDueKey_(schedule,date){
  const mins=gpsfmScheduleStartMinutes_(schedule);
  if(mins===null)return date+' 00:00';
  return date+' '+('0'+Math.floor(mins/60)).slice(-2)+':'+('0'+(mins%60)).slice(-2);
}
function apiGetAttendanceTargets(token){
  const s=gpsfmSession_(token),admin=String(s.user.role)==='superadmin',today=Utilities.formatDate(new Date(),'Asia/Jakarta','yyyy-MM-dd'),end=gpsfmDateFromKey_(today);
  // Baca daftar PJ sekali, lalu gunakan indeks agar tidak membaca seluruh sheet ActivityPJ
  // berulang kali untuk setiap kegiatan.
  const pjRows=admin?[]:gpsfmRows_('ActivityPJ'),pjActivityIds={};
  if(!admin)pjRows.forEach(function(row){if(gpsfmBool_(row.active)&&String(row.userId)===String(s.user.userId))pjActivityIds[String(row.activityId)]=true;});
  const activities=gpsfmRows_('Activities').filter(function(a){return gpsfmBool_(a.active)&&(admin||!!pjActivityIds[String(a.activityId)]);}),attendance=gpsfmRows_('Attendance');
  // Target ditentukan per kegiatan dan tanggal jadwal, bukan per guru.
  // Satu catatan absensi apa pun untuk kegiatan/tanggal tersebut menandakan target sudah diisi.
  const attendanceDates={};
  attendance.forEach(function(row){
    const key=String(row.activityId)+'|'+gpsfmDateKey_(row.attendanceDate);
    attendanceDates[key]=true;
  });
  // Indeks occurrence dibangun sekali; jangan baca seluruh ActivityOccurrences untuk tiap tanggal.
  // Jika ada duplikat lama, pertahankan perilaku lama: gunakan baris pertama yang cocok.
  gpsfmEnsureOccurrenceSheet_();
  const occurrenceByKey={};
  gpsfmRows_('ActivityOccurrences').forEach(function(row){
    const key=String(row.activityId)+'|'+gpsfmDateKey_(row.occurrenceDate);
    if(!Object.prototype.hasOwnProperty.call(occurrenceByKey,key))occurrenceByKey[key]=row;
  });
  const nowJakarta=Utilities.formatDate(new Date(),'Asia/Jakarta','yyyy-MM-dd HH:mm');
  const result=[];
  activities.forEach(function(a){
    const days=gpsfmParseScheduleDays_(a.schedule);if(!days.length)return;
    const monthStart=gpsfmDateFromKey_(today.slice(0,7)+'-01');
    // Target selalu menelusuri seluruh jadwal sejak tanggal 1 bulan berjalan,
    // tanpa membatasi rentang berdasarkan createdAt kegiatan.
    const cursor=monthStart;cursor.setUTCHours(12,0,0,0);const occurrences=[];
    while(cursor<=end){
      const date=gpsfmDateKeyUTC_(cursor);
      if(days.indexOf(gpsfmWeekdayName_(date))>=0){
        const due=gpsfmScheduledDueKey_(a.schedule,date),hasClock=gpsfmScheduleStartMinutes_(a.schedule)!==null;
        const isDue=hasClock?nowJakarta>=due:date<today;
        if(isDue){
          const occurrenceKey=String(a.activityId)+'|'+date;
          let o=occurrenceByKey[occurrenceKey];
          if(!o){
            o=gpsfmSaveOccurrence_(a.activityId,date,'scheduled','', '',[]);
            occurrenceByKey[occurrenceKey]=o;
          }
          if(['holiday','cancelled'].indexOf(String(o.status||'scheduled'))<0&&!attendanceDates[occurrenceKey])occurrences.push({activityId:String(a.activityId),activityName:String(a.name),schedule:String(a.schedule||''),date:date,weekday:gpsfmWeekdayName_(date),missingCount:1,status:String(o.status||'scheduled'),reason:String(o.reason||'')});
        }
      }
      cursor.setUTCDate(cursor.getUTCDate()+1);
    }
    if(occurrences.length)result.push({activityId:String(a.activityId),activityName:String(a.name),schedule:String(a.schedule||''),overdueCount:occurrences.length,missingTotal:occurrences.length,occurrences:occurrences.sort(function(x,y){return y.date.localeCompare(x.date);})});
  });
  return {asOf:today,totalOccurrences:result.reduce(function(n,a){return n+a.overdueCount;},0),totalMissing:result.reduce(function(n,a){return n+a.missingTotal;},0),activities:result};
}
function apiSetOccurrenceException(token,activityId,date,status,reason){
 const s=gpsfmSession_(token),id=gpsfmText_(activityId,80),key=gpsfmText_(date,10);status=String(status||'');reason=gpsfmText_(reason,300);
 if(s.user.role!=='superadmin'&&!gpsfmIsPJ_(s.user.userId,id))throw new Error('Anda bukan PJ untuk kegiatan ini.');
 const a=gpsfmActivity_(id);if(!a)throw new Error('Kegiatan tidak aktif atau tidak ditemukan.');
 if(!/^\d{4}-\d{2}-\d{2}$/.test(key))throw new Error('Tanggal tidak valid.');
 if(gpsfmParseScheduleDays_(a.schedule).indexOf(gpsfmWeekdayName_(key))<0)throw new Error('Tanggal tersebut bukan hari jadwal kegiatan ini.');
 if(['holiday','cancelled','scheduled'].indexOf(status)<0)throw new Error('Status pengecualian tidak valid.');
 if((status==='holiday'||status==='cancelled')&&!reason)throw new Error('Alasan libur atau pembatalan wajib diisi.');
 if(key>Utilities.formatDate(new Date(),'Asia/Jakarta','yyyy-MM-dd'))throw new Error('Pengecualian hanya dapat ditetapkan untuk pertemuan hari ini atau yang sudah lewat.');
 const old=gpsfmGetOccurrence_(id,key),ids=gpsfmOccurrenceTeacherIds_(old,id);
 gpsfmSaveOccurrence_(id,key,status,status==='scheduled'?'':reason,s.user.userId,ids);
 gpsfmAudit_(s.user.userId,'SET_OCCURRENCE_EXCEPTION','ActivityOccurrences',id,{date:key,status:status,reason:reason});return {ok:true};
}

function apiListMyAttendanceActivities(token) {
  const s = gpsfmSession_(token);
  const activities = gpsfmRows_('Activities').filter(function(a) { return gpsfmBool_(a.active); });
  // Evaluasi Diri selalu bersifat pribadi untuk seluruh role, termasuk Superadmin.
  const ids = {};
  gpsfmRows_('Attendance').forEach(function(row) {
    if (String(row.teacherId) === String(s.user.userId)) ids[String(row.activityId)] = true;
  });
  return activities.filter(function(a) { return !!ids[String(a.activityId)]; }).map(function(a) {
    return {activityId:String(a.activityId),name:String(a.name),schedule:String(a.schedule||'')};
  }).sort(function(a,b) { return a.name.localeCompare(b.name); });
}

function apiGetSupervisorDashboard(token) {
  const s = gpsfmSession_(token);
  if (String(s.user.role) !== 'supervisor' && String(s.user.role) !== 'admin_administrasi') throw new Error('Akses ditolak. Fitur ini hanya untuk Supervisor atau Admin Administrasi.');
  const today = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
  const prefix = today.slice(0, 7);
  const activities = gpsfmRows_('Activities').filter(function(a) { return gpsfmBool_(a.active); });
  const users = gpsfmRows_('Users').filter(function(u) { return gpsfmBool_(u.active) && gpsfmIsParticipantRole_(u.role); });
  const userMap = {};
  users.forEach(function(u) { userMap[String(u.userId)] = String(u.fullName); });
  const pjRows = gpsfmRows_('ActivityPJ').filter(function(p) { return gpsfmBool_(p.active); });
  const teacherRows = gpsfmRows_('ActivityTeachers');
  const attendance = gpsfmRows_('Attendance').filter(function(a) { return gpsfmDateKey_(a.attendanceDate).slice(0,7) === prefix; });
  const pjsByActivity = {}, teachersByActivity = {}, recordsByActivity = {}, pjUsers = {};
  pjRows.forEach(function(p) {
    const id=String(p.activityId);
    if(!pjsByActivity[id])pjsByActivity[id]=[];
    pjsByActivity[id].push(userMap[String(p.userId)] || 'Akun tidak ditemukan');
    pjUsers[String(p.userId)]=true;
  });
  teacherRows.forEach(function(t) {
    const id=String(t.activityId);
    if(!teachersByActivity[id])teachersByActivity[id]=[];
    teachersByActivity[id].push(String(t.teacherId));
  });
  attendance.forEach(function(row) {
    const id=String(row.activityId);
    if(!recordsByActivity[id])recordsByActivity[id]=[];
    recordsByActivity[id].push(row);
  });
  const result = activities.map(function(a) {
    const id=String(a.activityId),pjs=pjsByActivity[id]||[],teacherIds=teachersByActivity[id]||[],records=recordsByActivity[id]||[],present={};
    records.forEach(function(row){present[String(row.teacherId)]=true;});
    const presentTeacherCount=Object.keys(present).length,rosterCount=teacherIds.length;
    return {activityId:id,name:String(a.name),schedule:String(a.schedule||''),pjs:pjs,rosterCount:rosterCount,monthlyRecords:records.length,presentTeacherCount:presentTeacherCount,coveragePercent:rosterCount?Math.min(100,Math.round(presentTeacherCount/rosterCount*100)):0};
  });
  return {month:prefix,totalActivities:activities.length,totalPJ:Object.keys(pjUsers).length,totalPJAssignments:pjRows.length,totalMonthlyRecords:attendance.length,activities:result};
}

function apiGetTodayScheduledActivities(token) {
  const s = gpsfmSession_(token);
  const role = String(s.user.role);
  if (role !== 'superadmin' && role !== 'supervisor') throw new Error('Akses ditolak. Fitur ini hanya untuk Superadmin atau Supervisor.');

  const today = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
  const weekday = gpsfmWeekdayName_(today);
  const activities = gpsfmRows_('Activities').filter(function(a) {
    return gpsfmBool_(a.active) && gpsfmParseScheduleDays_(a.schedule).indexOf(weekday) >= 0;
  });
  const activityIds = {};
  activities.forEach(function(a) { activityIds[String(a.activityId)] = true; });

  const attendanceByActivity = {};
  gpsfmRows_('Attendance').forEach(function(row) {
    const id = String(row.activityId);
    if (activityIds[id] && gpsfmDateKey_(row.attendanceDate) === today) attendanceByActivity[id] = true;
  });

  gpsfmEnsureOccurrenceSheet_();
  const occurrenceByActivity = {};
  gpsfmRows_('ActivityOccurrences').forEach(function(row) {
    const id = String(row.activityId);
    if (activityIds[id] && gpsfmDateKey_(row.occurrenceDate) === today && !occurrenceByActivity[id]) {
      occurrenceByActivity[id] = row;
    }
  });

  return {
    date: today,
    weekday: weekday,
    activities: activities.map(function(a) {
      const id = String(a.activityId);
      const occurrence = occurrenceByActivity[id];
      const saved = !!attendanceByActivity[id];
      const exception = occurrence && String(occurrence.status || '') === 'holiday' ? 'Libur' :
        occurrence && String(occurrence.status || '') === 'cancelled' ? 'Dibatalkan' : '';
      return {
        activityId: id,
        name: String(a.name || ''),
        attendanceSaved: saved,
        status: saved ? 'Sudah' : (exception || 'Belum'),
        reason: exception ? String(occurrence.reason || '') : ''
      };
    }).sort(function(a,b) { return a.name.localeCompare(b.name); })
  };
}

function apiGetMonthlyReport(token, year, month) {
  const s = gpsfmSession_(token);
  year = Number(year); month = Number(month);
  if (!Number.isInteger(year) || year < 2020 || year > 2100 || !Number.isInteger(month) || month < 1 || month > 12) throw new Error('Periode laporan tidak valid.');
  const prefix = year + '-' + ('0'+month).slice(-2);
  const users = gpsfmRows_('Users'), allActivities = gpsfmRows_('Activities').filter(function(a) { return gpsfmBool_(a.active); });
  const userMap = {}, actMap = {};
  users.forEach(function(u) { userMap[String(u.userId)] = String(u.fullName); });
  allActivities.forEach(function(a) { actMap[String(a.activityId)] = String(a.name); });
  const attendance = gpsfmRows_('Attendance'), allowedActivityIds = {}, rows = [];
  attendance.forEach(function(a) {
    const activityId=String(a.activityId);
    if(String(a.teacherId)!==String(s.user.userId) || !Object.prototype.hasOwnProperty.call(actMap,activityId))return;
    allowedActivityIds[activityId]=true;
    const date=gpsfmDateKey_(a.attendanceDate);
    if(date.slice(0,7)===prefix)rows.push(a);
  });
  const activities = allActivities.filter(function(a) { return !!allowedActivityIds[String(a.activityId)]; }).map(function(a) {
    return {activityId:String(a.activityId), activityName:String(a.name)};
  }).sort(function(a,b) { return a.activityName.localeCompare(b.activityName); });
  const result = rows.map(function(a) {
    const date = gpsfmDateKey_(a.attendanceDate);
    return {date:date,teacherId:String(a.teacherId),teacherName:userMap[String(a.teacherId)] || 'Guru tidak ditemukan',activityId:String(a.activityId),activityName:actMap[String(a.activityId)] || 'Kegiatan tidak ditemukan',status:String(a.status),note:String(a.note || '')};
  }).sort(function(a,b) { return a.date.localeCompare(b.date) || a.activityName.localeCompare(b.activityName) || a.teacherName.localeCompare(b.teacherName); });
  const summary = {'Hadir':0,'A. Sakit':0,'B. Tugas Pondok':0,'C. Izin Pribadi':0,'D. Alfa':0};
  result.forEach(function(r) {
    const raw = String(r.status || '').trim().replace(/^[A-D][.]\s*/i, '').toLowerCase().replace(/\s+/g, ' ');
    let key = '';
    if (raw === 'hadir') key = 'Hadir';
    else if (raw === 'sakit') key = 'A. Sakit';
    else if (raw === 'tugas' || raw === 'tugas pondok') key = 'B. Tugas Pondok';
    else if (raw === 'izin' || raw === 'izin pribadi') key = 'C. Izin Pribadi';
    else if (raw === 'alpa' || raw === 'alfa') key = 'D. Alfa';
    if (key) summary[key]++;
  });
  return {year:year,month:month,summary:summary,rows:result,activities:activities};
}

function gpsfmExportStatusKey_(value) {
  const raw=String(value||'').trim().replace(/^[A-D][.]\s*/i,'').toLowerCase().replace(/\s+/g,' ');
  if(raw==='hadir')return 'Hadir';
  if(raw==='sakit')return 'A. Sakit';
  if(raw==='tugas'||raw==='tugas pondok')return 'B. Tugas Pondok';
  if(raw==='izin'||raw==='izin pribadi')return 'C. Izin Pribadi';
  if(raw==='alpa'||raw==='alfa')return 'D. Alfa';
  return '';
}

/**
 * Menyediakan data rekap untuk dibentuk menjadi XLSX di browser.
 * Struktur baris, urutan peserta, tanggal, dan ringkasan dipertahankan
 * dari mekanisme ekspor sebelumnya.
 */
function apiGetAttendanceExcelData(token,year,month) {
  const s=gpsfmSession_(token);
  year=Number(year);month=Number(month);
  if(!Number.isInteger(year)||year<2020||year>2100||!Number.isInteger(month)||month<1||month>12)throw new Error('Periode ekspor tidak valid.');
  const role=String(s.user.role),isGlobal=['superadmin','supervisor','admin_administrasi'].indexOf(role)>=0;
  if(!isGlobal&&role!=='gps')throw new Error('Akses ekspor tidak diizinkan.');
  const activities=gpsfmRows_('Activities').filter(function(a){return gpsfmBool_(a.active)&&(isGlobal||gpsfmIsPJ_(s.user.userId,a.activityId));});
  if(!activities.length)throw new Error(isGlobal?'Belum ada kegiatan aktif untuk diekspor.':'Anda belum ditugaskan sebagai PJ pada kegiatan aktif.');
  const users=gpsfmRows_('Users'),attendance=gpsfmRows_('Attendance'),prefix=year+'-'+('0'+month).slice(-2);
  const labels=['Hadir','A. Sakit','B. Tugas Pondok','C. Izin Pribadi','D. Alfa'];
  const teachers=users.filter(function(u){return gpsfmBool_(u.active)&&['gps','supervisor'].indexOf(String(u.role))>=0;});
  const teacherIds={};teachers.forEach(function(u){teacherIds[String(u.userId)]=true;});
  const attendanceByActivity={};
  attendance.forEach(function(a){
    const activityId=String(a.activityId);
    if(!teacherIds[String(a.teacherId)]||gpsfmDateKey_(a.attendanceDate).slice(0,7)!==prefix)return;
    if(!attendanceByActivity[activityId])attendanceByActivity[activityId]=[];
    attendanceByActivity[activityId].push(a);
  });
  const usedNames={};
  const sheets=activities.map(function(activity){
    const activityId=String(activity.activityId),relevant=attendanceByActivity[activityId]||[];
    const dates=Array.from(new Set(relevant.map(function(a){return gpsfmDateKey_(a.attendanceDate);}))).filter(Boolean).sort();
    const values={};
    relevant.forEach(function(a){values[String(a.teacherId)+'|'+gpsfmDateKey_(a.attendanceDate)]=String(a.status||'');});
    const header=['No','Nama Peserta'].concat(dates.map(function(d){return d.split('-').reverse().join('-');})).concat(labels).concat(['Total Catatan']);
    const rows=[['REKAP ABSENSI — '+String(activity.name||'Kegiatan')],['Periode: '+('0'+month).slice(-2)+'/'+year],[],header];
    const totals={};labels.forEach(function(k){totals[k]=0;});
    const dailyTotals={};dates.forEach(function(d){dailyTotals[d]=0;});
    dates.forEach(function(date){
      dailyTotals[date]=teachers.filter(function(u){
        return !!gpsfmExportStatusKey_(values[String(u.userId)+'|'+date]||'');
      }).length;
    });
    teachers.forEach(function(u,index){
      const counts={};labels.forEach(function(k){counts[k]=0;});
      const cells=dates.map(function(date){
        const raw=values[String(u.userId)+'|'+date]||'—',key=gpsfmExportStatusKey_(raw);
        if(key)counts[key]++;
        return raw;
      });
      labels.forEach(function(k){totals[k]+=counts[k];});
      const total=labels.reduce(function(sum,k){return sum+counts[k];},0);
      rows.push([index+1,String(u.fullName||'')].concat(cells).concat(labels.map(function(k){return counts[k];})).concat([total]));
    });
    rows.push(['','TOTAL'].concat(dates.map(function(date){return dailyTotals[date];})).concat(labels.map(function(k){return totals[k];})).concat([labels.reduce(function(sum,k){return sum+totals[k];},0)]));
    let base=String(activity.name||'Kegiatan').replace(/[\[\]:*?\/\\]/g,' ').replace(/^'+|'+$/g,'').trim();
    if(!base)base='Kegiatan';base=base.slice(0,31);
    let sheetName=base,n=2;
    while(usedNames[sheetName.toLowerCase()]){const suffix=' ('+n+')';n++;sheetName=base.slice(0,31-suffix.length)+suffix;}
    usedNames[sheetName.toLowerCase()]=true;
    return {name:sheetName,rows:rows,dateCount:dates.length,teacherCount:teachers.length};
  });
  gpsfmAudit_(s.user.userId,'EXPORT_ATTENDANCE_EXCEL','Attendance','',{year:year,month:month,activityCount:sheets.length});
  return {filename:'GPS-FM_Rekap-Absen_'+year+'-'+('0'+month).slice(-2)+'.xlsx',year:year,month:month,sheets:sheets,activityCount:sheets.length};
}

function apiGetActivityReport(token, activityId) {
  const s = gpsfmSession_(token);
  activityId = gpsfmText_(activityId,80);
  if (s.user.role !== 'superadmin' && s.user.role !== 'supervisor' && s.user.role !== 'admin_administrasi' && !gpsfmIsPJ_(s.user.userId, activityId)) throw new Error('Anda tidak memiliki akses laporan kegiatan ini.');
  const activity = gpsfmActivity_(activityId);
  if (!activity) throw new Error('Kegiatan tidak aktif atau tidak ditemukan.');
  const teachers = apiGetActivityTeachers(token, activityId);
  const teacherIds = teachers.map(function(t){return String(t.userId);});
  const rows = gpsfmRows_('Attendance').filter(function(a){return String(a.activityId) === activityId && teacherIds.indexOf(String(a.teacherId)) >= 0;});
  const dates = Array.from(new Set(rows.map(function(a){return gpsfmDateKey_(a.attendanceDate);}))).filter(Boolean).sort();
  const values = {};
  rows.forEach(function(a){values[String(a.teacherId)+'|'+gpsfmDateKey_(a.attendanceDate)] = String(a.status).replace(/^[A-D]\\. /,'');});
  return {activityName:String(activity.name),dates:dates,teachers:teachers.map(function(t){return {userId:String(t.userId),fullName:String(t.fullName),statuses:dates.map(function(d){return values[String(t.userId)+'|'+d] || '—';})};})};
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
  const s = gpsfmSession_(token, true);
  if (String(newPassword || '').length < 6) throw new Error('Password baru minimal 6 karakter.');
  if (gpsfmHashPassword_(String(oldPassword || ''),s.user.passwordSalt) !== String(s.user.passwordHash)) throw new Error('Password lama tidak sesuai.');
  const salt = gpsfmId_().replace(/-/g,'');
  gpsfmUpdateRow_('Users',s.user._row,{passwordSalt:salt,passwordHash:gpsfmHashPassword_(newPassword,salt),mustChangePassword:false,updatedAt:gpsfmNow_()});
  gpsfmAudit_(s.user.userId,'CHANGE_PASSWORD','Users',s.user.userId,{});
  return {ok:true};
}
