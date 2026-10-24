/**
 * GPS FM — SetupDatabase.gs
 * Jalankan setupGpsFmDatabase() satu kali dari editor Apps Script.
 * Opsional: isi Script Properties GPSFM_INITIAL_ADMIN_PASSWORD sebelum setup.
 * Spreadsheet ID disimpan di Script Properties, bukan ditanam di source code.
 */
const GPSFM_SCHEMA = {
  Users: ['userId','username','fullName','role','passwordSalt','passwordHash','active','createdAt','updatedAt'],
  Activities: ['activityId','name','description','schedule','active','createdAt','updatedAt'],
  ActivityPJ: ['assignmentId','activityId','userId','active','assignedAt','assignedBy'],
  Attendance: ['attendanceId','activityId','teacherId','attendanceDate','status','note','recordedBy','createdAt','updatedAt'],
  Periods: ['periodId','year','month','status','lockedAt','lockedBy'],
  AuditLogs: ['logId','actorId','action','entityType','entityId','details','createdAt'],
  AppSettings: ['key','value','updatedAt']
};

function setupGpsFmDatabase() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('GPSFM_SPREADSHEET_ID');
  let ss;
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create('GPS FM — Database');
    id = ss.getId();
    props.setProperty('GPSFM_SPREADSHEET_ID', id);
  }

  Object.keys(GPSFM_SCHEMA).forEach(function(name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const headers = GPSFM_SCHEMA[name];
    const current = sh.getLastRow() ? sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), headers.length)).getDisplayValues()[0] : [];
    if (!current.length || current.every(function(v) { return !v; })) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    } else {
      const missing = headers.filter(function(h) { return current.indexOf(h) < 0; });
      if (missing.length) throw new Error('Struktur sheet ' + name + ' tidak sesuai. Sheet tidak diubah agar data lama aman. Kolom yang belum ada: ' + missing.join(', '));
    }
  });

  const users = ss.getSheetByName('Users');
  const rows = users.getDataRange().getValues();
  const hasAdmin = rows.slice(1).some(function(r) { return String(r[3]) === 'superadmin' && String(r[6]) === 'TRUE'; });
  if (!hasAdmin) {
    const initialPassword = props.getProperty('GPSFM_INITIAL_ADMIN_PASSWORD');
    if (!initialPassword || initialPassword.length < 12) {
      throw new Error('Database siap, tetapi superadmin belum dibuat. Atur Script Property GPSFM_INITIAL_ADMIN_PASSWORD dengan password awal minimal 12 karakter, lalu jalankan setupGpsFmDatabase() lagi. Setelah login pertama, segera ganti password.');
    }
    const salt = Utilities.getUuid().replace(/-/g, '');
    const hash = gpsfmHashPassword_(initialPassword, salt);
    users.appendRow([Utilities.getUuid(), 'superadmin', 'Superadmin GPS FM', 'superadmin', salt, hash, true, new Date().toISOString(), new Date().toISOString()]);
    props.deleteProperty('GPSFM_INITIAL_ADMIN_PASSWORD');
  }

  const settings = ss.getSheetByName('AppSettings');
  const settingRows = settings.getDataRange().getValues();
  const known = settingRows.slice(1).map(function(r) { return String(r[0]); });
  [['APP_NAME','GPS FM'],['TIMEZONE','Asia/Jakarta'],['ATTENDANCE_STATUSES','Hadir,Izin,Sakit,Alpa']].forEach(function(pair) {
    if (known.indexOf(pair[0]) < 0) settings.appendRow([pair[0], pair[1], new Date().toISOString()]);
  });

  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1 && defaultSheet.getLastRow() === 0) ss.deleteSheet(defaultSheet);
  Logger.log('GPS FM database siap: ' + ss.getUrl());
  return {spreadsheetId:id, spreadsheetUrl:ss.getUrl(), sheets:Object.keys(GPSFM_SCHEMA)};
}
