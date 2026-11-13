/**
 * GPS FM — Code.gs
 * Backend untuk Web App Apps Script. Otorisasi selalu dilakukan di server.
 */
const GPSFM_SESSION_TTL = 21600; // 6 jam untuk sesi biasa
const GPSFM_REMEMBER_SESSION_TTL = 2592000; // 30 hari jika Remember Me aktif
const GPSFM_SESSION_HEADERS = ['tokenHash','userId','passwordHash','issuedAt','expiresAt','active','updatedAt'];
// Cache objek Spreadsheet hanya selama satu eksekusi Apps Script.
var GPSFM_DB_INSTANCE_ = null;
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
  if (GPSFM_DB_INSTANCE_) return GPSFM_DB_INSTANCE_;
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
    GPSFM_DB_INSTANCE_ = active;
    return GPSFM_DB_INSTANCE_;
  }

  const id = PropertiesService.getScriptProperties().getProperty('GPSFM_SPREADSHEET_ID');
  if (!id) {
    throw new Error('Spreadsheet induk belum terhubung. Buka Apps Script melalui Ekstensi → Apps Script dari spreadsheet GPS FM, lalu jalankan setupGpsFmDatabase().');
  }
  GPSFM_DB_INSTANCE_ = SpreadsheetApp.openById(id);
  return GPSFM_DB_INSTANCE_;
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

/**
 * Cari satu baris berdasarkan beberapa kolom kunci.
 * Hanya membaca kolom kunci saat pencarian, lalu mengambil satu baris lengkap
 * untuk menjaga bentuk objek dan nomor baris yang dipakai fungsi lama.
 * modes: lower = perbandingan teks tanpa membedakan kapital; date = normalisasi tanggal.
 */
function gpsfmFindRowByCriteria_(sheetName, criteria, modes, predicate) {
  const sh = gpsfmSheet_(sheetName), lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return null;
  const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(String);
  const keys = Object.keys(criteria || {});
  if (!keys.length) throw new Error('Kriteria pencarian database kosong.');
  const columns = {};
  keys.forEach(function(key) {
    const index = headers.indexOf(key);
    if (index < 0) throw new Error('Kolom database tidak ditemukan: ' + key + ' (' + sheetName + ').');
    columns[key] = sh.getRange(2, index + 1, lastRow - 1, 1).getValues();
  });
  const compare = function(actual, expected, mode) {
    if (mode === 'date') return gpsfmDateKey_(actual) === String(expected);
    if (mode === 'lower') return String(actual == null ? '' : actual).trim().toLowerCase() === String(expected == null ? '' : expected).trim().toLowerCase();
    return String(actual == null ? '' : actual) === String(expected == null ? '' : expected);
  };
  for (let rowIndex = 0; rowIndex < lastRow - 1; rowIndex++) {
    let matched = true;
    for (let i = 0; i < keys.length; i++) {
      const key = keys[i];
      if (!compare(columns[key][rowIndex][0], criteria[key], modes && modes[key])) {
        matched = false;
        break;
      }
    }
    if (matched) {
      const rowNumber = rowIndex + 2;
      const values = sh.getRange(rowNumber, 1, 1, lastCol).getValues()[0];
      const obj = {_row:rowNumber};
      headers.forEach(function(header, index) { obj[header] = values[index]; });
      if (!predicate || predicate(obj)) return obj;
    }
  }
  return null;
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
  const cache = CacheService.getScriptCache();
  if (cache.get('gpsfm_schema_password_flag_v1') === '1') return;
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
  cache.put('gpsfm_schema_password_flag_v1', '1', 21600);
}
function gpsfmAudit_(actorId, action, entityType, entityId, details) {
  gpsfmAppend_('AuditLogs', {logId:gpsfmId_(), actorId:actorId || '', action:action, entityType:entityType || '', entityId:entityId || '', details:JSON.stringify(details || {}), createdAt:gpsfmNow_()});
}
function gpsfmTokenKey_(token) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(token), Utilities.Charset.UTF_8);
  return 'gpsfm_session_' + bytes.map(function(b) { return ('0' + ((b+256)%256).toString(16)).slice(-2); }).join('');
}
function gpsfmEnsureSessionSheet_() {
  const ss = gpsfmDb_();
  let sh = ss.getSheetByName('UserSessions');
  if (!sh) {
    sh = ss.insertSheet('UserSessions');
    sh.getRange(1,1,1,GPSFM_SESSION_HEADERS.length).setValues([GPSFM_SESSION_HEADERS]);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,GPSFM_SESSION_HEADERS.length).setFontWeight('bold');
  }
  return sh;
}
function gpsfmFindSessionRow_(tokenHash) {
  const sh = gpsfmEnsureSessionSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  const match = sh.getRange(2,1,lastRow-1,1).createTextFinder(String(tokenHash)).matchEntireCell(true).matchCase(true).findNext();
  if (!match) return null;
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(String);
  const values = sh.getRange(match.getRow(),1,1,headers.length).getValues()[0];
  const data = {_row:match.getRow()};
  headers.forEach(function(h,i){data[h]=values[i];});
  return data;
}
function gpsfmSaveSession_(tokenHash, session) {
  const sh = gpsfmEnsureSessionSheet_();
  const values = {
    tokenHash:tokenHash,userId:String(session.userId),passwordHash:String(session.passwordHash),
    issuedAt:Number(session.issuedAt),expiresAt:Number(session.expiresAt),active:true,updatedAt:gpsfmNow_()
  };
  sh.appendRow(GPSFM_SESSION_HEADERS.map(function(h){return Object.prototype.hasOwnProperty.call(values,h)?values[h]:'';}));
}
function gpsfmSession_(token, allowForcedPasswordChange) {
  if (!token || String(token).length < 30) throw new Error('Sesi tidak valid. Silakan login kembali.');
  const tokenHash = gpsfmTokenKey_(token);
  let session = gpsfmFindSessionRow_(tokenHash);

  // Migrasi sesi lama secara otomatis selama token lama masih tersedia di cache.
  // Setelah migrasi, validasi sesi tidak lagi bergantung pada CacheService.
  if (!session) {
    const raw = CacheService.getScriptCache().get(tokenHash);
    if (!raw) throw new Error('Sesi berakhir. Silakan login kembali.');
    try { session = JSON.parse(raw); } catch (e) { throw new Error('Sesi tidak valid. Silakan login kembali.'); }
    const issuedAt = Number(session.issuedAt || Date.now());
    const expiresAt = issuedAt + GPSFM_SESSION_TTL * 1000;
    if (expiresAt <= Date.now()) throw new Error('Sesi berakhir. Silakan login kembali.');
    session.expiresAt = expiresAt;
    gpsfmSaveSession_(tokenHash, session);
    session = gpsfmFindSessionRow_(tokenHash);
  }

  if (!session || !gpsfmBool_(session.active) || Number(session.expiresAt) <= Date.now()) {
    throw new Error('Sesi berakhir. Silakan login kembali.');
  }
  const user = gpsfmFindRowByValue_('Users', 'userId', session.userId);
  if (!user || !gpsfmBool_(user.active)) throw new Error('Akun tidak aktif atau tidak ditemukan.');
  if (String(session.passwordHash || '') !== String(user.passwordHash || '')) {
    gpsfmUpdateRow_('UserSessions',session._row,{active:false,updatedAt:gpsfmNow_()});
    throw new Error('Password akun berubah. Silakan login kembali.');
  }
  if(gpsfmBool_(user.mustChangePassword)&&!allowForcedPasswordChange)throw new Error('Anda wajib mengganti password sebelum menggunakan aplikasi.');
  return {user:user, token:token, session:session};
}
// Perbarui hanya sesi yang sedang mengganti password; sesi perangkat lain tetap tidak berubah.
function gpsfmRefreshCurrentSessionPasswordHash_(token, userId, passwordHash) {
  if (!token || String(token).length < 30) return false;
  const tokenHash = gpsfmTokenKey_(token);
  const session = gpsfmFindSessionRow_(tokenHash);
  if (!session || String(session.userId || '') !== String(userId) || !gpsfmBool_(session.active)) return false;
  gpsfmUpdateRow_('UserSessions',session._row,{passwordHash:String(passwordHash),updatedAt:gpsfmNow_()});
  CacheService.getScriptCache().put(tokenHash,JSON.stringify({
    userId:String(userId),issuedAt:Number(session.issuedAt),passwordHash:String(passwordHash)
  }),GPSFM_SESSION_TTL);
  return true;
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
  const assignment = gpsfmFindRowByCriteria_('ActivityTeachers', {activityId:String(activityId),teacherId:String(teacherId)}, null, function(row) {
    return gpsfmBool_(row.active);
  });
  return !!assignment;
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
    .createTextFinder(String(username)).matchEntireCell(true).matchCase(false).findNext();
  if (!matches) return null;
  const rowNumber = matches.getRow();
  const values = sh.getRange(rowNumber, 1, 1, lastCol).getValues()[0];
  const user = {_row:rowNumber};
  headers.forEach(function(header, index) { user[header] = values[index]; });
  return user;
}

function apiLogin(username, password, rememberMe) {
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
  const issuedAt = Date.now();
  const expiresIn = gpsfmBool_(rememberMe) ? GPSFM_REMEMBER_SESSION_TTL : GPSFM_SESSION_TTL;
  const session = {userId:String(user.userId),issuedAt:issuedAt,expiresAt:issuedAt+expiresIn*1000,passwordHash:String(user.passwordHash)};
  const tokenHash = gpsfmTokenKey_(token);
  gpsfmSaveSession_(tokenHash,session);
  // Cache hanya akselerator; sumber sesi yang tahan restart berada di sheet UserSessions.
  cache.put(tokenHash,JSON.stringify(session),GPSFM_SESSION_TTL);
  gpsfmAudit_(user.userId, 'LOGIN', 'Users', user.userId, {});
  return {token:token, user:gpsfmPublicUser_(user), expiresIn:expiresIn};
}

function apiLogout(token) {
  // Revoke token directly so logout also works during mandatory password change
  // or when the account has just been disabled. Never affect other tab tokens.
  if (!token || String(token).length < 30) return {ok:true};
  const cache = CacheService.getScriptCache();
  const key = gpsfmTokenKey_(token);
  let session = null;
  try {
    session = gpsfmFindSessionRow_(key);
    if (session) gpsfmUpdateRow_('UserSessions',session._row,{active:false,updatedAt:gpsfmNow_()});
  } catch (e) {}
  // Revoke both persistent and cached copies before audit logging.
  cache.remove(key);
  try {
    if (session && session.userId) {
      const user = gpsfmFindRowByValue_('Users', 'userId', session.userId);
      if (user) gpsfmAudit_(user.userId, 'LOGOUT', 'Users', user.userId, {});
    }
  } catch (e) {}
  return {ok:true};
}


/**
 * Jalankan satu kali dari editor Apps Script untuk memicu permintaan
 * otorisasi scope script.external_request.
 * Endpoint tetap dan tanpa kredensial; respons HTTP non-2xx tidak menggagalkan otorisasi.
 */
function authorizeExternalRequest() {
  const response = UrlFetchApp.fetch('https://api.onesignal.com', {
    method: 'get',
    muteHttpExceptions: true
  });
  return {
    authorized: true,
    httpStatus: response.getResponseCode()
  };
}

/* OneSignal Web Push — langganan perangkat dan pengiriman berdasarkan jadwal aktual. */
const GPSFM_ONESIGNAL_APP_ID = '28faec3d-9ee1-45a0-9234-4a5922ce4f1a';
const GPSFM_PUSH_SUBSCRIPTION_HEADERS = ['subscriptionId','userId','active','deviceLabel','createdAt','updatedAt','lastSeenAt'];
const GPSFM_PUSH_DELIVERY_HEADERS = ['dedupeKey','activityId','occurrenceDate','recipientUserId','notificationType','subscriptionIdsJson','status','response','createdAt'];

function gpsfmEnsurePushSheets_() {
  const ss = gpsfmDb_();
  [
    ['PushSubscriptions', GPSFM_PUSH_SUBSCRIPTION_HEADERS],
    ['PushDeliveryLog', GPSFM_PUSH_DELIVERY_HEADERS]
  ].forEach(function(def) {
    let sh = ss.getSheetByName(def[0]);
    if (!sh) {
      sh = ss.insertSheet(def[0]);
      sh.getRange(1,1,1,def[1].length).setValues([def[1]]);
      sh.setFrozenRows(1);
      sh.getRange(1,1,1,def[1].length).setFontWeight('bold');
      return;
    }
    const current = sh.getLastRow() ? sh.getRange(1,1,1,Math.max(sh.getLastColumn(),def[1].length)).getDisplayValues()[0] : [];
    const missing = def[1].filter(function(h) { return current.indexOf(h) < 0; });
    if (missing.length) throw new Error('Struktur sheet ' + def[0] + ' tidak sesuai. Kolom hilang: ' + missing.join(', '));
  });
}

function apiRegisterPushSubscription(token, subscriptionId, deviceLabel) {
  const session = gpsfmSession_(token);
  subscriptionId = gpsfmText_(subscriptionId,100);
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(subscriptionId)) {
    throw new Error('ID langganan OneSignal tidak valid.');
  }
  gpsfmEnsurePushSheets_();
  const rows = gpsfmRows_('PushSubscriptions');
  const old = rows.find(function(r) { return String(r.subscriptionId) === subscriptionId; });
  const now = gpsfmNow_();
  const values = {subscriptionId:subscriptionId,userId:String(session.user.userId),active:true,
    deviceLabel:gpsfmText_(deviceLabel,180),updatedAt:now,lastSeenAt:now};
  if (old) gpsfmUpdateRow_('PushSubscriptions',old._row,values);
  else {
    values.createdAt = now;
    gpsfmAppend_('PushSubscriptions',values);
  }
  return {ok:true,registered:true};
}

function apiUnregisterPushSubscription(token, subscriptionId) {
  const session = gpsfmSession_(token);
  subscriptionId = gpsfmText_(subscriptionId,100);
  gpsfmEnsurePushSheets_();
  const old = gpsfmRows_('PushSubscriptions').find(function(r) {
    return String(r.subscriptionId) === subscriptionId && String(r.userId) === String(session.user.userId) && gpsfmBool_(r.active);
  });
  if (old) gpsfmUpdateRow_('PushSubscriptions',old._row,{active:false,updatedAt:gpsfmNow_()});
  return {ok:true,unregistered:!!old};
}

/** Kirim notifikasi uji hanya ke perangkat pemanggil yang terdaftar pada akun aktif. */
function apiSendPushTest(token, subscriptionId) {
  const session = gpsfmSession_(token);
  subscriptionId = gpsfmText_(subscriptionId,100);
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(subscriptionId)) {
    throw new Error('ID perangkat tidak valid. Aktifkan notifikasi terlebih dahulu.');
  }
  gpsfmEnsurePushSheets_();
  const subscription = gpsfmRows_('PushSubscriptions').find(function(r) {
    return String(r.subscriptionId) === subscriptionId &&
      String(r.userId) === String(session.user.userId) && gpsfmBool_(r.active);
  });
  if (!subscription) throw new Error('Perangkat ini belum terdaftar aktif pada akun Anda. Aktifkan ulang notifikasi.');
  const response = gpsfmSendOneSignal_([subscriptionId],
    'Tes Notifikasi GPS FM',
    'Berhasil! Perangkat ini dapat menerima push notification dari GPS FM.',
    {type:'test',url:'https://appfathanmubina.github.io/GPS-FM/pwa/'});
  const now = gpsfmNow_();
  gpsfmAppend_('PushDeliveryLog',{
    dedupeKey:'test|' + Utilities.getUuid(),activityId:'TEST',
    occurrenceDate:Utilities.formatDate(new Date(),'Asia/Jakarta','yyyy-MM-dd'),
    recipientUserId:String(session.user.userId),notificationType:'test',
    subscriptionIdsJson:JSON.stringify([subscriptionId]),status:'sent',
    response:'accepted id=' + String(response.id),createdAt:now
  });
  return {ok:true,accepted:true,messageId:String(response.id)};
}

function gpsfmPushDeliveryKey_(activityId,date,userId,type,subscriptionId) {
  return [type,activityId,date,userId,subscriptionId].join('|');
}

function gpsfmSendOneSignal_(subscriptionIds, title, message, data) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ONESIGNAL_REST_API_KEY');
  if (!apiKey) throw new Error('Script Property ONESIGNAL_REST_API_KEY belum diatur.');
  if (!subscriptionIds.length) return {};
  const notificationPayload = {
    app_id:GPSFM_ONESIGNAL_APP_ID,include_subscription_ids:subscriptionIds,
    headings:{en:title,id:title},contents:{en:message,id:message},data:data || {}
  };
  // OneSignal hanya membuka halaman tujuan jika URL dikirim pada properti top-level,
  // bukan sekadar disisipkan ke custom data.
  if (data && data.url) notificationPayload.url = String(data.url);
  const response = UrlFetchApp.fetch('https://api.onesignal.com/notifications', {
    method:'post',contentType:'application/json',headers:{Authorization:'Key ' + apiKey},
    payload:JSON.stringify(notificationPayload),
    muteHttpExceptions:true
  });
  const status = response.getResponseCode(), body = response.getContentText();
  if (status < 200 || status >= 300) throw new Error('OneSignal HTTP ' + status + ': ' + body.slice(0,500));
  let parsed = {};
  try { parsed = JSON.parse(body); } catch (e) {}
  if (parsed.errors && Object.keys(parsed.errors).length) throw new Error('OneSignal menolak pengiriman: ' + JSON.stringify(parsed.errors).slice(0,500));
  // HTTP 200 saja belum cukup: OneSignal mengembalikan ID pesan hanya jika pesan berhasil dibuat.
  if (!parsed.id) throw new Error('OneSignal tidak membuat pesan. Tidak ada ID pesan; periksa subscription aktif dan konfigurasi Web Push. Respons: ' + body.slice(0,400));
  return parsed;
}

/** Trigger setiap menit: pengingat kegiatan sebelum jadwal dan pengingat absensi setelah mulai. */
function gpsfmRunScheduledPushNotifications() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return {ok:false,skipped:'another-run'};
  try {
    gpsfmEnsurePushSheets_();
    const now = new Date(), today = Utilities.formatDate(now,'Asia/Jakarta','yyyy-MM-dd');
    const nowMinute = Number(Utilities.formatDate(now,'Asia/Jakarta','HH')) * 60 + Number(Utilities.formatDate(now,'Asia/Jakarta','mm'));
    const weekday = gpsfmWeekdayName_(today);
    const activities = gpsfmRows_('Activities').filter(function(a) {
      return gpsfmBool_(a.active) && gpsfmParseScheduleDays_(a.schedule).indexOf(weekday) >= 0;
    });
    const activeUsers = {};
    gpsfmRows_('Users').forEach(function(u) { if (gpsfmBool_(u.active)) activeUsers[String(u.userId)] = u; });
    const teacherAssignments = gpsfmRows_('ActivityTeachers').filter(function(a) { return gpsfmBool_(a.active); });
    const pjAssignments = gpsfmRows_('ActivityPJ').filter(function(a) { return gpsfmBool_(a.active); });
    const attendanceByActivityDate = {};
    gpsfmRows_('Attendance').forEach(function(a) {
      const key = String(a.activityId)+'|'+gpsfmDateKey_(a.attendanceDate);
      if (!attendanceByActivityDate[key]) attendanceByActivityDate[key] = {};
      attendanceByActivityDate[key][String(a.teacherId)] = true;
    });
    // Ambil status sekaligus roster guru khusus tanggal ini. Roster occurrence
    // harus diutamakan agar perubahan penugasan tidak mengubah peserta tanggal lampau/hari ini.
    const occurrencesByActivity = {};
    gpsfmEnsureOccurrenceSheet_();
    gpsfmRows_('ActivityOccurrences').forEach(function(o) {
      if (gpsfmDateKey_(o.occurrenceDate) === today) occurrencesByActivity[String(o.activityId)] = o;
    });
    const subscriptionsByUser = {};
    gpsfmRows_('PushSubscriptions').forEach(function(p) {
      if (gpsfmBool_(p.active) && activeUsers[String(p.userId)]) {
        if (!subscriptionsByUser[String(p.userId)]) subscriptionsByUser[String(p.userId)] = [];
        subscriptionsByUser[String(p.userId)].push(String(p.subscriptionId));
      }
    });
    const sent = {};
    gpsfmRows_('PushDeliveryLog').forEach(function(r) { if (String(r.status) === 'sent') sent[String(r.dedupeKey)] = true; });
    const results = {activitiesChecked:activities.length,teacherNotifications:0,pjNotifications:0,failed:0,skippedNoTime:0};
    const pendingLogs = [];

    function deliver(activity,type,userIds,title,message) {
      const ids = [], keys = [];
      userIds.forEach(function(userId) {
        (subscriptionsByUser[String(userId)] || []).forEach(function(subscriptionId) {
          const key = gpsfmPushDeliveryKey_(activity.activityId,today,userId,type,subscriptionId);
          if (!sent[key]) { ids.push(subscriptionId); keys.push({key:key,userId:String(userId),subscriptionId:subscriptionId}); }
        });
      });
      if (!ids.length) return;
      let status = 'sent', detail = 'accepted';
      try {
        const oneSignalResult = gpsfmSendOneSignal_(ids,title,message,{type:type,activityId:String(activity.activityId),date:today,url:'https://appfathanmubina.github.io/GPS-FM/pwa/'});
        detail = 'accepted id=' + String(oneSignalResult.id) +
          (oneSignalResult.recipients !== undefined ? ' recipients=' + String(oneSignalResult.recipients) : '');
      } catch (error) {
        status = 'failed'; detail = String(error.message || error).slice(0,500); results.failed++;
      }
      keys.forEach(function(item) {
        if (status === 'sent') sent[item.key] = true;
        pendingLogs.push({dedupeKey:item.key,activityId:String(activity.activityId),occurrenceDate:today,
          recipientUserId:item.userId,notificationType:type,subscriptionIdsJson:JSON.stringify([item.subscriptionId]),
          status:status,response:detail,createdAt:gpsfmNow_()});
      });
      if (status === 'sent' && type === 'activity') results.teacherNotifications += keys.length;
      if (status === 'sent' && type === 'attendance') results.pjNotifications += keys.length;
    }

    activities.forEach(function(activity) {
      const activityId = String(activity.activityId), startMinute = gpsfmScheduleStartMinutes_(activity.schedule);
      if (startMinute === null) { results.skippedNoTime++; return; }
      const occurrenceStatus = String((occurrencesByActivity[activityId] || {}).status || 'scheduled');
      if (['holiday','cancelled'].indexOf(occurrenceStatus) >= 0) return;

      const teachers = {};
      const occurrence = occurrencesByActivity[activityId] || null;
      let candidateTeacherIds = teacherAssignments.filter(function(a) {
        return String(a.activityId) === activityId;
      }).map(function(a) { return String(a.teacherId); });
      if (occurrence && occurrence.teacherIdsJson) {
        try {
          const snapshotIds = JSON.parse(String(occurrence.teacherIdsJson));
          if (Array.isArray(snapshotIds)) candidateTeacherIds = snapshotIds.map(String);
        } catch (e) {}
      }
      candidateTeacherIds.forEach(function(id) {
        const user = activeUsers[id];
        if (user && gpsfmIsParticipantRole_(user.role)) teachers[id] = true;
      });
      const teacherIds = Object.keys(teachers);
      const pjs = {};
      pjAssignments.forEach(function(a) {
        const id = String(a.userId);
        if (String(a.activityId) === activityId && activeUsers[id]) pjs[id] = true;
      });
      // Pengingat kegiatan dikirim kepada seluruh guru roster dan PJ aktif.
      // Daftar roster tetap dipisahkan agar pemeriksaan kelengkapan absensi akurat.
      const activityRecipients = Object.assign({}, teachers, pjs);
      const activityRecipientIds = Object.keys(activityRecipients);
      // Lima menit sebelum jadwal. Jendela ini mengantisipasi trigger yang berjalan sedikit terlambat;
      // dedupe berdasarkan kegiatan/tanggal/penerima memastikan hanya satu notifikasi terkirim.
      if (nowMinute >= startMinute - 5 && nowMinute < startMinute && activityRecipientIds.length) {
        deliver(activity,'activity',activityRecipientIds,'Pengingat Kegiatan GPS FM',
          'Kegiatan ' + String(activity.name || 'GPS FM') + ' akan dimulai 5 menit lagi. Silakan bersiap dan buka GPS FM untuk melihat penugasan Anda.');
      }

      // Pertahankan pengingat absensi kepada PJ selama 30 menit pertama setelah kegiatan dimulai.
      const recordedTeachers = attendanceByActivityDate[activityId+'|'+today] || {};
      const hasMissingAttendance = teacherIds.some(function(teacherId) { return !recordedTeachers[teacherId]; });
      if (nowMinute >= startMinute && nowMinute < startMinute + 30 && teacherIds.length && hasMissingAttendance) {
        const pjIds = Object.keys(pjs);
        if (pjIds.length) deliver(activity,'attendance',pjIds,'Pengingat Input Absensi GPS FM',
          'Anda bertugas sebagai PJ kegiatan ' + String(activity.name || 'GPS FM') + '. Masih ada absensi guru yang belum lengkap. Mohon lengkapi input absensi pada GPS FM.');
      }
    });

    if (pendingLogs.length) {
      const sh = gpsfmSheet_('PushDeliveryLog'), headers = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
      const rows = pendingLogs.map(function(obj) { return headers.map(function(h) { return Object.prototype.hasOwnProperty.call(obj,h) ? obj[h] : ''; }); });
      sh.getRange(sh.getLastRow()+1,1,rows.length,headers.length).setValues(rows);
    }
    return results;
  } finally { lock.releaseLock(); }
}

/**
 * Pemeriksaan penerima push tanpa mengirim notifikasi.
 * Hanya dapat dijalankan oleh pemilik/editor project dari Apps Script editor.
 */
function previewGpsFmPushRecipients() {
  const activeEmail = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  const effectiveEmail = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  if (!activeEmail || !effectiveEmail || activeEmail !== effectiveEmail) {
    throw new Error('Preview ini hanya dapat dijalankan langsung oleh pemilik project dari Apps Script editor.');
  }

  gpsfmEnsurePushSheets_();
  const today = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
  const weekday = gpsfmWeekdayName_(today);
  const now = new Date();
  const nowMinute = Number(Utilities.formatDate(now, 'Asia/Jakarta', 'HH')) * 60 +
    Number(Utilities.formatDate(now, 'Asia/Jakarta', 'mm'));
  const users = {};
  gpsfmRows_('Users').forEach(function(u) {
    if (gpsfmBool_(u.active)) users[String(u.userId)] = u;
  });
  const subscriptions = {};
  gpsfmRows_('PushSubscriptions').forEach(function(p) {
    if (gpsfmBool_(p.active) && users[String(p.userId)]) {
      const id = String(p.userId);
      subscriptions[id] = (subscriptions[id] || 0) + 1;
    }
  });
  const attendance = {};
  gpsfmRows_('Attendance').forEach(function(a) {
    const key = String(a.activityId) + '|' + gpsfmDateKey_(a.attendanceDate);
    if (!attendance[key]) attendance[key] = {};
    attendance[key][String(a.teacherId)] = true;
  });
  // Gunakan snapshot roster per tanggal yang sama dengan proses pengiriman.
  const occurrencesByActivity = {};
  gpsfmEnsureOccurrenceSheet_();
  gpsfmRows_('ActivityOccurrences').forEach(function(o) {
    if (gpsfmDateKey_(o.occurrenceDate) === today) {
      occurrencesByActivity[String(o.activityId)] = o;
    }
  });
  const teacherAssignments = gpsfmRows_('ActivityTeachers').filter(function(a) { return gpsfmBool_(a.active); });
  const pjAssignments = gpsfmRows_('ActivityPJ').filter(function(a) { return gpsfmBool_(a.active); });
  const report = {
    date: today,
    weekday: weekday,
    generatedAt: Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss'),
    triggerHandlers: ScriptApp.getProjectTriggers().map(function(t) { return t.getHandlerFunction(); }),
    activities: []
  };

  gpsfmRows_('Activities').forEach(function(activity) {
    if (!gpsfmBool_(activity.active) ||
        gpsfmParseScheduleDays_(activity.schedule).indexOf(weekday) < 0) return;
    const activityId = String(activity.activityId);
    const startMinute = gpsfmScheduleStartMinutes_(activity.schedule);
    const occurrence = occurrencesByActivity[activityId] || null;
    const occurrenceStatus = String((occurrence || {}).status || 'scheduled');
    const cancelled = ['holiday', 'cancelled'].indexOf(occurrenceStatus) >= 0;
    const withinWindow = startMinute !== null && nowMinute >= startMinute && nowMinute < startMinute + 30;
    const teacherMap = {};
    let candidateTeacherIds = teacherAssignments.filter(function(a) {
      return String(a.activityId) === activityId;
    }).map(function(a) { return String(a.teacherId); });
    if (occurrence && occurrence.teacherIdsJson) {
      try {
        const snapshotIds = JSON.parse(String(occurrence.teacherIdsJson));
        if (Array.isArray(snapshotIds)) candidateTeacherIds = snapshotIds.map(String);
      } catch (e) {}
    }
    candidateTeacherIds.forEach(function(id) {
      const user = users[id];
      if (user && gpsfmIsParticipantRole_(user.role)) teacherMap[id] = user;
    });
    const teacherIds = Object.keys(teacherMap);
    const recorded = attendance[activityId + '|' + today] || {};
    const missingTeacherIds = teacherIds.filter(function(id) { return !recorded[id]; });
    const pjMap = {};
    if (teacherIds.length && missingTeacherIds.length && !cancelled) {
      pjAssignments.forEach(function(a) {
        const id = String(a.userId), user = users[id];
        if (String(a.activityId) === activityId && user) pjMap[id] = user;
      });
    }
    function recipientList(map) {
      return Object.keys(map).map(function(id) {
        return {
          fullName: String(map[id].fullName || map[id].username || id),
          role: String(map[id].role || ''),
          activeSubscriptions: subscriptions[id] || 0
        };
      }).sort(function(a, b) { return a.fullName.localeCompare(b.fullName); });
    }
    report.activities.push({
      activityName: String(activity.name || activityId),
      schedule: String(activity.schedule || ''),
      startTime: startMinute === null ? null : ('0' + Math.floor(startMinute / 60)).slice(-2) + ':' + ('0' + (startMinute % 60)).slice(-2),
      withinReminderWindow: withinWindow,
      occurrenceStatus: occurrenceStatus,
      skippedBecause: cancelled ? 'occurrence-' + occurrenceStatus : (startMinute === null ? 'schedule-time-not-found' : ''),
      assignedTeachers: recipientList(teacherMap),
      teachersMissingAttendance: missingTeacherIds.length,
      assignedPJsEligibleForReminder: recipientList(pjMap)
    });
  });
  Logger.log(JSON.stringify(report, null, 2));
  return report;
}

function setupGpsFmPushNotifications() {
  gpsfmEnsurePushSheets_();
  if (!PropertiesService.getScriptProperties().getProperty('ONESIGNAL_REST_API_KEY')) {
    throw new Error('Atur Script Property ONESIGNAL_REST_API_KEY terlebih dahulu. Gunakan REST API Key OneSignal, bukan App ID.');
  }
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'gpsfmRunScheduledPushNotifications') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('gpsfmRunScheduledPushNotifications').timeBased().everyMinutes(1).create();
  return {ok:true,intervalMinutes:1,activityReminderMinutesBefore:5,sheets:['PushSubscriptions','PushDeliveryLog']};
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
  const duplicate = gpsfmFindRowByCriteria_('Users', {username:username}, {username:'lower'}, function(u) {
    return String(u.userId) !== String(data.userId || '');
  });
  if (duplicate && String(duplicate.userId) !== String(data.userId || '')) throw new Error('Username sudah digunakan.');
  const now = gpsfmNow_(), existing = data.userId ? gpsfmFindRowByCriteria_('Users', {userId:String(data.userId)}) : null;
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
  const u = gpsfmFindRowByCriteria_('Users', {userId:String(userId)});
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
  const now = gpsfmNow_(), existing = data.activityId ? gpsfmFindRowByCriteria_('Activities', {activityId:String(data.activityId)}) : null;
  if (existing) gpsfmUpdateRow_('Activities',existing._row,{name:name,description:gpsfmText_(data.description,500),schedule:gpsfmText_(data.schedule,120),active:data.active !== false,updatedAt:now});
  else gpsfmAppend_('Activities',{activityId:gpsfmId_(),name:name,description:gpsfmText_(data.description,500),schedule:gpsfmText_(data.schedule,120),active:true,createdAt:now,updatedAt:now});
  gpsfmAudit_(s.user.userId,existing?'UPDATE_ACTIVITY':'CREATE_ACTIVITY','Activities',existing?existing.activityId:name,{name:name});
  return {ok:true};
}

function apiAssignPJ(token, activityId, userId, assigned) {
  const s = gpsfmSession_(token); gpsfmRequireAdmin_(s);
  const activity = gpsfmFindRowByCriteria_('Activities', {activityId:String(activityId)});
  const user = gpsfmFindRowByCriteria_('Users', {userId:String(userId)});
  if (!activity || !user || !gpsfmBool_(user.active) || !gpsfmIsParticipantRole_(user.role)) throw new Error('Kegiatan atau akun peserta tidak valid.');
  const old = gpsfmFindRowByCriteria_('ActivityPJ', {activityId:String(activityId),userId:String(userId)});
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
  const activity = gpsfmFindRowByCriteria_('Activities', {activityId:String(activityId)});
  const teacher = gpsfmFindRowByCriteria_('Users', {userId:String(teacherId)});
  if (!activity || !teacher || !gpsfmBool_(teacher.active) || !gpsfmIsParticipantRole_(teacher.role)) throw new Error('Kegiatan atau akun peserta tidak valid.');
  const old = gpsfmFindRowByCriteria_('ActivityTeachers', {activityId:String(activityId),teacherId:String(teacherId)});
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
  if (!gpsfmIsValidAttendanceDate_(date)) throw new Error('Tanggal tidak valid.');
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

function gpsfmIsValidAttendanceDate_(date) {
  const value = String(date || '');
  const match = value.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return false;
  const year = Number(match[1]), month = Number(match[2]), day = Number(match[3]);
  const check = new Date(Date.UTC(year, month - 1, day));
  return check.getUTCFullYear() === year && check.getUTCMonth() === month - 1 && check.getUTCDate() === day;
}

function gpsfmValidatePJAttendanceMonth_(user, date) {
  if (!gpsfmIsValidAttendanceDate_(date)) throw new Error('Tanggal tidak valid.');
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
  if (!gpsfmIsValidAttendanceDate_(date)) throw new Error('Tanggal tidak valid.');
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
  if (!gpsfmIsValidAttendanceDate_(date)) throw new Error('Tanggal tidak valid.');
  gpsfmValidatePJAttendanceMonth_(u, date);
  if (GPSFM_STATUSES.indexOf(status) < 0) throw new Error('Status kehadiran tidak valid.');
  const period = date.slice(0,7).split('-'), p = gpsfmRows_('Periods').find(function(x) { return Number(x.year) === Number(period[0]) && Number(x.month) === Number(period[1]); });
  if (p && String(p.status) === 'locked') throw new Error('Periode laporan sudah dikunci.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const old = gpsfmFindRowByCriteria_('Attendance', {activityId:activityId,teacherId:teacherId,attendanceDate:date}, {attendanceDate:'date'});
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
 return gpsfmFindRowByCriteria_('ActivityOccurrences', {activityId:String(activityId),occurrenceDate:String(date)}, {occurrenceDate:'date'});
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
  // Hanya hitung penugasan aktif yang terhubung ke akun peserta aktif.
  // Riwayat penugasan tetap disimpan, tetapi tidak boleh membesarkan roster dashboard.
  const pjRows = gpsfmRows_('ActivityPJ').filter(function(p) {
    return gpsfmBool_(p.active) && Object.prototype.hasOwnProperty.call(userMap, String(p.userId));
  });
  const teacherRows = gpsfmRows_('ActivityTeachers').filter(function(t) {
    return gpsfmBool_(t.active) && Object.prototype.hasOwnProperty.call(userMap, String(t.teacherId));
  });
  const attendance = gpsfmRows_('Attendance').filter(function(a) { return gpsfmDateKey_(a.attendanceDate).slice(0,7) === prefix; });
  const pjsByActivity = {}, teachersByActivity = {}, recordsByActivity = {}, pjUsers = {};
  pjRows.forEach(function(p) {
    const id=String(p.activityId), userId=String(p.userId);
    if(!pjsByActivity[id])pjsByActivity[id]=[];
    if(pjsByActivity[id].indexOf(userMap[userId])<0)pjsByActivity[id].push(userMap[userId]);
    pjUsers[userId]=true;
  });
  teacherRows.forEach(function(t) {
    const id=String(t.activityId), teacherId=String(t.teacherId);
    if(!teachersByActivity[id])teachersByActivity[id]=[];
    if(teachersByActivity[id].indexOf(teacherId)<0)teachersByActivity[id].push(teacherId);
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
  const userId = String(s.user.userId), role = String(s.user.role);
  const canSeeAllActivities = ['superadmin','supervisor','admin_administrasi'].indexOf(role) >= 0;
  const teacherActivityIds = {}, pjActivityIds = {}, assignedActivityIds = {};
  if (!canSeeAllActivities) {
    gpsfmRows_('ActivityTeachers').forEach(function(item) {
      if (String(item.teacherId) === userId && gpsfmBool_(item.active)) {
        teacherActivityIds[String(item.activityId)] = true;
        assignedActivityIds[String(item.activityId)] = true;
      }
    });
    gpsfmRows_('ActivityPJ').forEach(function(item) {
      if (String(item.userId) === userId && gpsfmBool_(item.active)) {
        pjActivityIds[String(item.activityId)] = true;
        assignedActivityIds[String(item.activityId)] = true;
      }
    });
  }

  const today = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
  const weekday = gpsfmWeekdayName_(today);
  const activities = gpsfmRows_('Activities').filter(function(a) {
    const id = String(a.activityId);
    return gpsfmBool_(a.active) &&
      gpsfmParseScheduleDays_(a.schedule).indexOf(weekday) >= 0 &&
      (canSeeAllActivities || !!assignedActivityIds[id]);
  });
  const activityIds = {};
  activities.forEach(function(a) { activityIds[String(a.activityId)] = true; });

  const attendanceByActivity = {};
  gpsfmRows_('Attendance').forEach(function(row) {
    const id = String(row.activityId);
    if (activityIds[id] && gpsfmDateKey_(row.attendanceDate) === today) {
      if (!attendanceByActivity[id]) attendanceByActivity[id] = {};
      attendanceByActivity[id][String(row.teacherId)] = true;
    }
  });

  gpsfmEnsureOccurrenceSheet_();
  const occurrenceByActivity = {};
  gpsfmRows_('ActivityOccurrences').forEach(function(row) {
    const id = String(row.activityId);
    if (activityIds[id] && gpsfmDateKey_(row.occurrenceDate) === today && !occurrenceByActivity[id]) {
      occurrenceByActivity[id] = row;
    }
  });
  const activeTeacherAssignments = gpsfmRows_('ActivityTeachers').filter(function(row) {
    return gpsfmBool_(row.active) && activityIds[String(row.activityId)];
  });

  return {
    date: today,
    weekday: weekday,
    activities: activities.map(function(a) {
      const id = String(a.activityId), occurrence = occurrenceByActivity[id];
      const exception = occurrence && String(occurrence.status || '') === 'holiday' ? 'Libur' :
        occurrence && String(occurrence.status || '') === 'cancelled' ? 'Dibatalkan' : '';
      let rosterIds = activeTeacherAssignments.filter(function(row) {
        return String(row.activityId) === id;
      }).map(function(row) { return String(row.teacherId); });
      if (occurrence && occurrence.teacherIdsJson) {
        try {
          const snapshotIds = JSON.parse(String(occurrence.teacherIdsJson));
          if (Array.isArray(snapshotIds)) rosterIds = snapshotIds.map(String);
        } catch (e) {}
      }
      rosterIds = rosterIds.filter(function(value,index,array) { return array.indexOf(value) === index; });
      const recorded = attendanceByActivity[id] || {};
      const ownTeacher = !!teacherActivityIds[id] || rosterIds.indexOf(userId) >= 0;
      let status = 'Belum';
      if (exception) status = exception;
      else if (ownTeacher) status = recorded[userId] ? 'Sudah' : 'Belum';
      else if (pjActivityIds[id] || canSeeAllActivities) {
        const complete = rosterIds.length > 0 && rosterIds.every(function(teacherId) { return !!recorded[teacherId]; });
        status = complete ? 'Lengkap' : 'Belum Lengkap';
      }
      return {
        activityId: id,
        name: String(a.name || ''),
        schedule: String(a.schedule || ''),
        status: status,
        attendanceCount: rosterIds.filter(function(teacherId) { return !!recorded[teacherId]; }).length,
        teacherCount: rosterIds.length,
        reason: exception ? String(occurrence.reason || '') : ''
      };
    }).sort(function(a,b) { return a.name.localeCompare(b.name); })
  };
}
/**
 * Ringkasan pencapaian kehadiran seluruh kegiatan yang diikuti akun aktif.
 * Data selalu dibatasi pada userId dari sesi; role lain tidak dapat meminta
 * data guru lain melalui parameter.
 */
function apiGetMyActivityAchievement(token) {
  const s = gpsfmSession_(token), userId = String(s.user.userId);
  const today = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
  const nowJakarta = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm');
  const activities = gpsfmRows_('Activities'), activityMap = {};
  activities.forEach(function(a) { activityMap[String(a.activityId)] = a; });
  const assignments = gpsfmRows_('ActivityTeachers').filter(function(row) { return String(row.teacherId) === userId; });
  const assignedByActivity = {};
  assignments.forEach(function(row) {
    const id = String(row.activityId);
    if (!assignedByActivity[id]) assignedByActivity[id] = [];
    assignedByActivity[id].push(row);
  });
  const occurrenceByKey = {};
  gpsfmEnsureOccurrenceSheet_();
  gpsfmRows_('ActivityOccurrences').forEach(function(row) {
    const key = String(row.activityId) + '|' + gpsfmDateKey_(row.occurrenceDate);
    if (!occurrenceByKey[key]) occurrenceByKey[key] = row;
  });
  const attendanceByKey = {};
  gpsfmRows_('Attendance').forEach(function(row) {
    if (String(row.teacherId) === userId) attendanceByKey[String(row.activityId) + '|' + gpsfmDateKey_(row.attendanceDate)] = String(row.status || '');
  });
  function dueByToday(date, schedule) {
    if (date < today) return true;
    if (date > today) return false;
    const minutes = gpsfmScheduleStartMinutes_(schedule);
    return minutes !== null && nowJakarta >= gpsfmScheduledDueKey_(schedule, date);
  }
  function addMeeting(date, id, set, schedule) {
    if (!date || date > today || !dueByToday(date, schedule)) return;
    const occurrence = occurrenceByKey[id + '|' + date];
    if (occurrence && ['holiday','cancelled'].indexOf(String(occurrence.status || '').toLowerCase()) >= 0) return;
    set[date] = true;
  }
  const result = [];
  Object.keys(assignedByActivity).forEach(function(id) {
    const activity = activityMap[id];
    if (!activity) return;
    const schedule = String(activity.schedule || ''), days = gpsfmParseScheduleDays_(schedule);
    if (!days.length) return;
    const meetingDates = {};
    // Snapshot roster menjaga riwayat guru meskipun penugasan kemudian diubah.
    Object.keys(occurrenceByKey).forEach(function(key) {
      if (key.indexOf(id + '|') !== 0) return;
      const occurrence = occurrenceByKey[key];
      let ids = [];
      try { ids = JSON.parse(String(occurrence.teacherIdsJson || '[]')); } catch (e) {}
      const date = key.slice(id.length + 1);
      if (Array.isArray(ids) && ids.map(String).indexOf(userId) >= 0 && days.indexOf(gpsfmWeekdayName_(date)) >= 0) {
        addMeeting(date, id, meetingDates, schedule);
      }
    });
    // Lengkapi jadwal sejak tanggal penugasan untuk penugasan yang masih aktif.
    assignedByActivity[id].filter(function(row) { return gpsfmBool_(row.active); }).forEach(function(row) {
      const assignedAt = gpsfmDateKey_(row.assignedAt) || gpsfmDateKey_(activity.createdAt) || today;
      const createdAt = gpsfmDateKey_(activity.createdAt) || assignedAt;
      const cursor = gpsfmDateFromKey_(assignedAt > createdAt ? assignedAt : createdAt);
      const end = gpsfmDateFromKey_(today);
      cursor.setUTCHours(12,0,0,0);
      while (cursor <= end) {
        const date = gpsfmDateKeyUTC_(cursor);
        if (days.indexOf(gpsfmWeekdayName_(date)) >= 0) addMeeting(date, id, meetingDates, schedule);
        cursor.setUTCDate(cursor.getUTCDate() + 1);
      }
    });
    const dates = Object.keys(meetingDates).sort();
    let attended = 0, sick = 0, duty = 0, leave = 0, absent = 0, unrecorded = 0;
    dates.forEach(function(date) {
      const raw = String(attendanceByKey[id + '|' + date] || '').trim().replace(/^[A-D][.]\\s*/i, '').toLowerCase().replace(/\\s+/g, ' ');
      if (!raw) { unrecorded++; return; }
      if (raw === 'hadir') attended++;
      else if (raw === 'sakit') sick++;
      else if (raw === 'tugas' || raw === 'tugas pondok') duty++;
      else if (raw === 'izin' || raw === 'izin pribadi') leave++;
      else if (raw === 'alpa' || raw === 'alfa') absent++;
      else unrecorded++;
    });
    const total = dates.length;
    result.push({activityId:id,activityName:String(activity.name || ''),totalMeetings:total,attended:attended,sick:sick,duty:duty,leave:leave,absent:absent,unrecorded:unrecorded,percentage:total ? Math.round(attended / total * 100) : 0});
  });
  result.sort(function(a,b) { return a.activityName.localeCompare(b.activityName); });
  const totals = result.reduce(function(sum, item) {
    sum.totalMeetings += item.totalMeetings; sum.attended += item.attended; return sum;
  }, {totalMeetings:0,attended:0});
  return {asOf:today,totalActivities:result.length,totalMeetings:totals.totalMeetings,attended:totals.attended,percentage:totals.totalMeetings ? Math.round(totals.attended / totals.totalMeetings * 100) : 0,activities:result};
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
  // Selaraskan roster ekspor dengan role peserta yang dapat ditugaskan ke kegiatan.
  const teachers=users.filter(function(u){return gpsfmBool_(u.active)&&gpsfmIsParticipantRole_(u.role);});
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
  rows.forEach(function(a){values[String(a.teacherId)+'|'+gpsfmDateKey_(a.attendanceDate)] = String(a.status).replace(/^[A-D]\. /,'');});
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
  const newPasswordHash = gpsfmHashPassword_(newPassword,salt);
  gpsfmUpdateRow_('Users',s.user._row,{
    passwordSalt:salt,
    passwordHash:newPasswordHash,
    mustChangePassword:false,
    updatedAt:gpsfmNow_()
  });

  // Sinkronkan hash pada token yang sedang dipakai agar pengguna tetap berada
  // di tab ini setelah ganti password. Token sesi lain tidak ikut diperbarui.
  const currentSessionUpdated = gpsfmRefreshCurrentSessionPasswordHash_(token,s.user.userId,newPasswordHash);
  gpsfmAudit_(s.user.userId,'CHANGE_PASSWORD','Users',s.user.userId,{});
  return {ok:true,sessionKept:currentSessionUpdated};
}
