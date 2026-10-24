/**
 * GPS FM — SetupDatabase.gs
 * Stage 1: database foundation. Safe to rerun; does not clear existing data.
 * Set Script Property GPS_FM_SPREADSHEET_ID to an existing spreadsheet ID,
 * or run gpsFmSetupDatabase() once to create a new database spreadsheet.
 */

const GPSFM_DB_VERSION = '1.0.0';
const GPSFM_SHEETS = {
  USERS: ['user_id','username','full_name','email','phone','password_salt','password_hash','role','active','must_change_password','created_at','updated_at','last_login_at'],
  ACTIVITIES: ['activity_id','name','description','activity_type','schedule_rule','pj_user_id','participant_mode','active','created_by','created_at','updated_at'],
  ACTIVITY_MEMBERS: ['membership_id','activity_id','user_id','active','start_date','end_date','created_by','created_at'],
  ATTENDANCE_SESSIONS: ['session_id','activity_id','session_date','title','status','notes','created_by','created_at','locked_at','locked_by'],
  ATTENDANCE_RECORDS: ['record_id','session_id','activity_id','user_id','attendance_status','notes','recorded_by','recorded_at','updated_at','updated_by'],
  MONTHLY_REPORTS: ['report_id','year_month','activity_id','generated_at','generated_by','record_count','summary_json'],
  AUDIT_LOGS: ['audit_id','event_at','actor_user_id','action','entity_type','entity_id','details_json'],
  APP_SETTINGS: ['setting_key','setting_value','updated_at','updated_by']
};

function gpsFmSetupDatabase() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('GPS_FM_SPREADSHEET_ID');
  let ss = null;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) {
      throw new Error('GPS_FM_SPREADSHEET_ID tidak dapat dibuka. Periksa Script Properties dan izin spreadsheet.');
    }
  } else {
    ss = SpreadsheetApp.create('GPS FM — Database');
    id = ss.getId();
    props.setProperty('GPS_FM_SPREADSHEET_ID', id);
  }

  Object.keys(GPSFM_SHEETS).forEach(function(key) {
    const headers = GPSFM_SHEETS[key];
    let sh = ss.getSheetByName(key);
    if (!sh) sh = ss.insertSheet(key);
    const current = sh.getLastRow() ? sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), headers.length)).getValues()[0] : [];
    const matches = headers.every(function(h, i) { return current[i] === h; });
    if (!matches) {
      if (sh.getLastRow() === 0) {
        sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      } else {
        // Never overwrite existing rows. Stop so schema drift can be reviewed.
        throw new Error('Header sheet ' + key + ' tidak sesuai. Data tidak diubah; audit header sebelum melanjutkan.');
      }
    }
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sh.autoResizeColumns(1, headers.length);
  });

  const settings = ss.getSheetByName('APP_SETTINGS');
  const rows = settings.getLastRow() > 1 ? settings.getRange(2,1,settings.getLastRow()-1,2).getValues() : [];
  const existing = {};
  rows.forEach(r => existing[String(r[0])] = true);
  const defaults = [
    ['APP_NAME','GPS FM'],
    ['DB_VERSION',GPSFM_DB_VERSION],
    ['TIMEZONE','Asia/Jakarta'],
    ['REPORT_MONTH_CLOSE_DAY','0']
  ];
  defaults.forEach(function(r) {
    if (!existing[r[0]]) settings.appendRow([r[0], r[1], new Date(), 'system']);
  });

  return { success: true, databaseId: id, databaseUrl: ss.getUrl(), version: GPSFM_DB_VERSION, sheets: Object.keys(GPSFM_SHEETS) };
}

/** Run once after setup to create the first Superadmin. The password is never stored in plaintext. */
function gpsFmCreateInitialSuperadmin(username, fullName, password) {
  username = String(username || '').trim().toLowerCase();
  fullName = String(fullName || '').trim();
  password = String(password || '');
  if (!/^[a-z0-9._-]{4,40}$/.test(username)) throw new Error('Username 4–40 karakter: huruf kecil, angka, titik, garis bawah atau strip.');
  if (fullName.length < 2) throw new Error('Nama lengkap wajib diisi.');
  if (password.length < 12) throw new Error('Kata sandi awal minimal 12 karakter.');
  const db = gpsFmDb_();
  const users = gpsFmRead_('USERS');
  if (users.some(u => u.username === username)) throw new Error('Username sudah digunakan.');
  if (users.some(u => u.role === 'superadmin' && u.active === true)) throw new Error('Superadmin aktif sudah ada.');
  const salt = Utilities.getUuid() + Utilities.getUuid();
  const hash = gpsFmHashPassword_(password, salt);
  gpsFmAppend_('USERS', {
    user_id: Utilities.getUuid(), username: username, full_name: fullName, email: '', phone: '',
    password_salt: salt, password_hash: hash, role: 'superadmin', active: true,
    must_change_password: true, created_at: new Date(), updated_at: new Date(), last_login_at: ''
  });
  gpsFmAudit_('system', 'CREATE_INITIAL_SUPERADMIN', 'USERS', username, { role: 'superadmin' });
  return { success: true, message: 'Superadmin awal dibuat. Login dan ganti kata sandi pada penggunaan pertama.' };
}

/** Internal helpers shared by Code.gs. */
function gpsFmDb_() {
  const id = PropertiesService.getScriptProperties().getProperty('GPS_FM_SPREADSHEET_ID');
  if (!id) throw new Error('Database belum disiapkan. Jalankan gpsFmSetupDatabase() terlebih dahulu.');
  return SpreadsheetApp.openById(id);
}
function gpsFmRead_(sheetName) {
  const sh = gpsFmDb_().getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet tidak ditemukan: ' + sheetName);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  return values.slice(1).filter(r => r.some(v => v !== '' && v !== null)).map(function(row) {
    const o = {};
    headers.forEach((h,i) => o[h] = gpsFmNormalizeCell_(row[i]));
    return o;
  });
}
function gpsFmNormalizeCell_(v) {
  if (v instanceof Date) return v.toISOString();
  if (v === 'TRUE') return true;
  if (v === 'FALSE') return false;
  return v;
}
function gpsFmAppend_(sheetName, obj) {
  const sh = gpsFmDb_().getSheetByName(sheetName);
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
  sh.appendRow(headers.map(h => Object.prototype.hasOwnProperty.call(obj,h) ? obj[h] : ''));
}
function gpsFmUpdateById_(sheetName, idColumn, id, changes) {
  const sh = gpsFmDb_().getSheetByName(sheetName);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) throw new Error('Data tidak ditemukan.');
  const headers = values[0].map(String), idIndex = headers.indexOf(idColumn);
  if (idIndex < 0) throw new Error('Kolom ID tidak ditemukan.');
  for (let r=1; r<values.length; r++) {
    if (String(values[r][idIndex]) === String(id)) {
      Object.keys(changes).forEach(function(k) {
        const c = headers.indexOf(k);
        if (c >= 0) sh.getRange(r+1,c+1).setValue(changes[k]);
      });
      return true;
    }
  }
  throw new Error('Data tidak ditemukan.');
}
function gpsFmHashPassword_(password, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + password, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(bytes);
}
function gpsFmAudit_(actorId, action, entityType, entityId, details) {
  try {
    gpsFmAppend_('AUDIT_LOGS', {
      audit_id: Utilities.getUuid(), event_at: new Date(), actor_user_id: actorId || 'system',
      action: action, entity_type: entityType, entity_id: String(entityId || ''),
      details_json: JSON.stringify(details || {})
    });
  } catch (e) { console.error('Audit log failed: ' + e.message); }
}
