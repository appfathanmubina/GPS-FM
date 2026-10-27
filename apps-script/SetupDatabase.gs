/**
 * GPS FM — SetupDatabase.gs
 * Jalankan setupGpsFmDatabase() dari editor Apps Script yang dibuka melalui
 * Ekstensi → Apps Script pada spreadsheet GPS FM.
 * Setup memakai spreadsheet induk yang sudah ada; tidak membuat spreadsheet baru.
 * Sebelum setup awal, isi Script Property GPSFM_INITIAL_ADMIN_PASSWORD.
 * Username awal default: abanadheera (dapat diganti melalui GPSFM_INITIAL_ADMIN_USERNAME).
 * Password awal boleh 6 karakter; setelah login, segera ganti dengan password yang lebih kuat.
 * ID spreadsheet induk disimpan untuk akses Web App ketika konteks aktif tidak tersedia.
 */
const GPSFM_SCHEMA = {
  Users: ['userId','username','fullName','role','passwordSalt','passwordHash','active','createdAt','updatedAt'],
  Activities: ['activityId','name','description','schedule','active','createdAt','updatedAt'],
  ActivityPJ: ['assignmentId','activityId','userId','active','assignedAt','assignedBy'],
  ActivityTeachers: ['assignmentId','activityId','teacherId','active','assignedAt','assignedBy','updatedAt'],
  Attendance: ['attendanceId','activityId','teacherId','attendanceDate','status','note','recordedBy','createdAt','updatedAt'],
  ActivityOccurrences: ['occurrenceId','activityId','occurrenceDate','status','reason','teacherIdsJson','updatedBy','updatedAt','createdAt'],
  Periods: ['periodId','year','month','status','lockedAt','lockedBy'],
  AuditLogs: ['logId','actorId','action','entityType','entityId','details','createdAt'],
  AppSettings: ['key','value','updatedAt']
};

function setupGpsFmDatabase() {
  const props = PropertiesService.getScriptProperties();
  // Wajib container-bound: sumber ID adalah spreadsheet tempat ekstensi Apps Script dibuka.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Spreadsheet induk tidak ditemukan. Buka project ini melalui Ekstensi → Apps Script dari spreadsheet GPS FM, lalu jalankan setup kembali.');
  }
  const id = ss.getId();
  // Simpan ID induk agar backend Web App tetap dapat membuka spreadsheet yang benar.
  props.setProperty('GPSFM_SPREADSHEET_ID', id);

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
  const hasAdmin = rows.slice(1).some(function(r) { return String(r[3]) === 'superadmin' && (r[6] === true || String(r[6]).toLowerCase() === 'true'); });
  if (!hasAdmin) {
    const initialPassword = props.getProperty('GPSFM_INITIAL_ADMIN_PASSWORD');
    if (!initialPassword || initialPassword.length < 6) {
      throw new Error('Database siap, tetapi superadmin belum dibuat. Atur Script Property GPSFM_INITIAL_ADMIN_PASSWORD dengan password awal minimal 6 karakter, lalu jalankan setupGpsFmDatabase() lagi. Password awal hanya untuk bootstrap; setelah login, segera ganti dengan password yang kuat.');
    }
    const initialUsername = String(props.getProperty('GPSFM_INITIAL_ADMIN_USERNAME') || 'abanadheera').trim().toLowerCase();
    if (!/^[a-z0-9._-]{3,40}$/.test(initialUsername)) {
      throw new Error('GPSFM_INITIAL_ADMIN_USERNAME hanya boleh berisi huruf kecil, angka, titik, garis bawah, atau tanda hubung (3–40 karakter).');
    }
    const duplicateUsername = rows.slice(1).some(function(r) { return String(r[1]).trim().toLowerCase() === initialUsername; });
    if (duplicateUsername) throw new Error('Username awal ' + initialUsername + ' sudah digunakan. Periksa sheet Users atau atur GPSFM_INITIAL_ADMIN_USERNAME.');
    const salt = Utilities.getUuid().replace(/-/g, '');
    const hash = gpsfmHashPassword_(initialPassword, salt);
    users.appendRow([Utilities.getUuid(), initialUsername, 'Superadmin GPS FM', 'superadmin', salt, hash, true, new Date().toISOString(), new Date().toISOString()]);
    props.deleteProperty('GPSFM_INITIAL_ADMIN_PASSWORD');
    props.deleteProperty('GPSFM_INITIAL_ADMIN_USERNAME');
  }

  const settings = ss.getSheetByName('AppSettings');
  const settingRows = settings.getDataRange().getValues();
  const known = settingRows.slice(1).map(function(r) { return String(r[0]); });
  [['APP_NAME','GPS FM'],['TIMEZONE','Asia/Jakarta'],['ATTENDANCE_STATUSES','Hadir,Izin,Sakit,Alpa,Tugas']].forEach(function(pair) {
    if (known.indexOf(pair[0]) < 0) settings.appendRow([pair[0], pair[1], new Date().toISOString()]);
  });

  // Tidak menghapus sheet lain dari spreadsheet induk; sheet non-GPS FM tetap dipertahankan.
  Logger.log('GPS FM database siap: ' + ss.getUrl());
  return {spreadsheetId:id, spreadsheetUrl:ss.getUrl(), sheets:Object.keys(GPSFM_SCHEMA)};
}
