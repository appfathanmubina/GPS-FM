/**
 * GPS FM — Code.gs
 * Stage 1–2 foundation: app shell, authentication, role checks, activity/member CRUD,
 * attendance sessions/records, and monthly CSV-ready report data.
 * For production, deploy as Web App and validate the GitHub PWA bridge against the
 * actual deployment URL. Do not expose spreadsheet IDs or password hashes to clients.
 */

const GPSFM_SESSION_TTL_SECONDS = 21600;
const GPSFM_ROLES = ['superadmin','pj','guru','administrasi'];

function doGet(e) {
  const pwa = e && e.parameter && e.parameter.view === 'api';
  if (pwa) return gpsFmJson_({success:true, app:'GPS FM', version:'1.0.0', api:'ready'});
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('GPS FM — Guru Pembina Santri')
    .addMetaTag('viewport','width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

/** JSON endpoint for health checks only. Full PWA API bridge requires deployment/CORS validation. */
function doPost(e) {
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (payload.action === 'health') return gpsFmJson_({success:true, app:'GPS FM', version:'1.0.0'});
    return gpsFmJson_({success:false, error:'PWA_API_BRIDGE_NOT_ENABLED', message:'Gunakan Apps Script UI untuk tahap awal; aktifkan bridge setelah pengujian deployment.'});
  } catch (err) {
    return gpsFmJson_({success:false,error:'BAD_REQUEST',message:'Payload tidak valid.'});
  }
}
function gpsFmJson_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

function apiLogin(username, password) {
  username = String(username || '').trim().toLowerCase();
  password = String(password || '');
  if (!username || !password) throw new Error('Username dan kata sandi wajib diisi.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const users = gpsFmRead_('USERS');
    const user = users.find(u => String(u.username).toLowerCase() === username && u.active === true);
    // Generic failure message avoids revealing whether a username exists.
    if (!user || gpsFmHashPassword_(password, String(user.password_salt)) !== String(user.password_hash)) {
      gpsFmAudit_('unknown', 'LOGIN_FAILED', 'USERS', username, {});
      throw new Error('Username atau kata sandi tidak sesuai, atau akun tidak aktif.');
    }
    const token = Utilities.getUuid() + Utilities.getUuid();
    const session = {
      userId: user.user_id, username: user.username, role: user.role,
      issuedAt: Date.now(), expiresAt: Date.now() + GPSFM_SESSION_TTL_SECONDS * 1000
    };
    CacheService.getScriptCache().put('GPSFM_SESSION_' + token, JSON.stringify(session), GPSFM_SESSION_TTL_SECONDS);
    PropertiesService.getScriptProperties().setProperty('GPSFM_SESSION_' + token, JSON.stringify(session));
    gpsFmUpdateById_('USERS','user_id',user.user_id,{last_login_at:new Date(),updated_at:new Date()});
    gpsFmAudit_(user.user_id, 'LOGIN_SUCCESS', 'USERS', user.user_id, {});
    return {success:true, token:token, user:gpsFmPublicUser_(user), mustChangePassword:user.must_change_password === true};
  } finally { lock.releaseLock(); }
}
function apiLogout(token) {
  const session = gpsFmSession_(token, false);
  if (token) {
    CacheService.getScriptCache().remove('GPSFM_SESSION_' + token);
    PropertiesService.getScriptProperties().deleteProperty('GPSFM_SESSION_' + token);
  }
  if (session) gpsFmAudit_(session.userId,'LOGOUT','USERS',session.userId,{});
  return {success:true};
}
function apiGetBootstrap(token) {
  const s = gpsFmSession_(token);
  const user = gpsFmUserById_(s.userId);
  return {success:true,user:gpsFmPublicUser_(user),activities:gpsFmVisibleActivities_(s),dashboard:gpsFmDashboard_(s)};
}
function apiChangePassword(token, currentPassword, newPassword) {
  const s = gpsFmSession_(token), user = gpsFmUserById_(s.userId);
  if (String(newPassword || '').length < 12) throw new Error('Kata sandi baru minimal 12 karakter.');
  if (gpsFmHashPassword_(String(currentPassword || ''), String(user.password_salt)) !== String(user.password_hash)) throw new Error('Kata sandi saat ini tidak sesuai.');
  const salt = Utilities.getUuid() + Utilities.getUuid();
  gpsFmUpdateById_('USERS','user_id',user.user_id,{password_salt:salt,password_hash:gpsFmHashPassword_(newPassword,salt),must_change_password:false,updated_at:new Date()});
  gpsFmAudit_(user.user_id,'CHANGE_PASSWORD','USERS',user.user_id,{});
  return {success:true};
}

/** Superadmin user management. New users must change their initial password. */
function apiCreateUser(token, input) {
  const s = gpsFmRequire_(token,['superadmin']);
  input = input || {};
  const username = String(input.username || '').trim().toLowerCase();
  const fullName = String(input.full_name || '').trim();
  const role = String(input.role || '');
  const password = String(input.password || '');
  if (!/^[a-z0-9._-]{4,40}$/.test(username)) throw new Error('Format username tidak valid.');
  if (fullName.length < 2) throw new Error('Nama lengkap wajib diisi.');
  if (GPSFM_ROLES.indexOf(role) < 0) throw new Error('Peran tidak valid.');
  if (password.length < 12) throw new Error('Kata sandi awal minimal 12 karakter.');
  if (gpsFmRead_('USERS').some(u => u.username === username)) throw new Error('Username sudah digunakan.');
  const salt = Utilities.getUuid() + Utilities.getUuid();
  const id = Utilities.getUuid();
  gpsFmAppend_('USERS',{user_id:id,username:username,full_name:fullName,email:String(input.email||''),phone:String(input.phone||''),password_salt:salt,password_hash:gpsFmHashPassword_(password,salt),role:role,active:true,must_change_password:true,created_at:new Date(),updated_at:new Date(),last_login_at:''});
  gpsFmAudit_(s.userId,'CREATE_USER','USERS',id,{role:role});
  return {success:true,userId:id};
}
function apiListUsers(token) {
  gpsFmRequire_(token,['superadmin','administrasi']);
  return {success:true,users:gpsFmRead_('USERS').map(gpsFmPublicUser_)};
}

/** Activity CRUD and PJ assignment. activity_type = rutin | insidental. */
function apiSaveActivity(token, input) {
  const s = gpsFmRequire_(token,['superadmin']);
  input = input || {};
  const name = String(input.name || '').trim();
  const type = String(input.activity_type || '');
  const pj = String(input.pj_user_id || '');
  const mode = String(input.participant_mode || 'selected');
  if (!name) throw new Error('Nama kegiatan wajib diisi.');
  if (['rutin','insidental'].indexOf(type) < 0) throw new Error('Jenis kegiatan tidak valid.');
  if (['selected','all_gps'].indexOf(mode) < 0) throw new Error('Mode peserta tidak valid.');
  const users = gpsFmRead_('USERS');
  if (!users.some(u => u.user_id === pj && u.active === true && ['pj','superadmin'].indexOf(u.role) >= 0)) throw new Error('PJ harus berupa akun aktif dengan peran PJ atau Superadmin.');
  const now = new Date();
  const id = String(input.activity_id || Utilities.getUuid());
  const exists = gpsFmRead_('ACTIVITIES').some(a => a.activity_id === id);
  const row = {activity_id:id,name:name,description:String(input.description||''),activity_type:type,schedule_rule:String(input.schedule_rule||''),pj_user_id:pj,participant_mode:mode,active:input.active !== false,created_by:s.userId,created_at:input.created_at || now,updated_at:now};
  if (exists) {
    const old = gpsFmRead_('ACTIVITIES').find(a => a.activity_id === id);
    if (!old) throw new Error('Kegiatan tidak ditemukan.');
    gpsFmUpdateById_('ACTIVITIES','activity_id',id,row);
  } else gpsFmAppend_('ACTIVITIES',row);
  gpsFmAudit_(s.userId,exists?'UPDATE_ACTIVITY':'CREATE_ACTIVITY','ACTIVITIES',id,{name:name,type:type,pj_user_id:pj});
  return {success:true,activityId:id};
}
function apiSetActivityMembers(token, activityId, userIds) {
  const s = gpsFmRequire_(token,['superadmin']);
  const activity = gpsFmRead_('ACTIVITIES').find(a => a.activity_id === activityId);
  if (!activity) throw new Error('Kegiatan tidak ditemukan.');
  if (!Array.isArray(userIds)) throw new Error('Daftar peserta tidak valid.');
  const validUsers = gpsFmRead_('USERS').filter(u => u.active === true && u.role !== 'administrasi');
  const validIds = new Set(validUsers.map(u => String(u.user_id)));
  const wanted = Array.from(new Set(userIds.map(String)));
  if (wanted.some(id => !validIds.has(id))) throw new Error('Daftar peserta berisi akun tidak aktif atau tidak valid.');
  const existing = gpsFmRead_('ACTIVITY_MEMBERS').filter(m => m.activity_id === activityId);
  existing.forEach(m => gpsFmUpdateById_('ACTIVITY_MEMBERS','membership_id',m.membership_id,{active:wanted.indexOf(String(m.user_id)) >= 0}));
  wanted.forEach(function(uid) {
    if (!existing.some(m => String(m.user_id) === uid)) gpsFmAppend_('ACTIVITY_MEMBERS',{membership_id:Utilities.getUuid(),activity_id:activityId,user_id:uid,active:true,start_date:new Date(),end_date:'',created_by:s.userId,created_at:new Date()});
  });
  gpsFmAudit_(s.userId,'SET_ACTIVITY_MEMBERS','ACTIVITIES',activityId,{participantCount:wanted.length});
  return {success:true,count:wanted.length};
}
function apiListActivities(token) {
  const s = gpsFmSession_(token);
  return {success:true,activities:gpsFmVisibleActivities_(s)};
}

/** Create a session and seed one pending record per configured participant. */
function apiCreateAttendanceSession(token, input) {
  const s = gpsFmSession_(token), activityId = String((input||{}).activity_id||'');
  const activity = gpsFmRead_('ACTIVITIES').find(a => a.activity_id === activityId && a.active === true);
  if (!activity) throw new Error('Kegiatan tidak ditemukan atau tidak aktif.');
  if (!gpsFmCanManageActivity_(s,activity)) throw new Error('Anda tidak berwenang menginput kegiatan ini.');
  const date = String(input.session_date || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Tanggal kegiatan tidak valid.');
  const existing = gpsFmRead_('ATTENDANCE_SESSIONS').find(x => x.activity_id === activityId && String(x.session_date).slice(0,10) === date && x.status !== 'cancelled');
  if (existing) throw new Error('Sesi kegiatan pada tanggal ini sudah ada.');
  const id = Utilities.getUuid(), now = new Date();
  gpsFmAppend_('ATTENDANCE_SESSIONS',{session_id:id,activity_id:activityId,session_date:new Date(date+'T00:00:00+07:00'),title:String(input.title||activity.name),status:'draft',notes:String(input.notes||''),created_by:s.userId,created_at:now,locked_at:'',locked_by:''});
  const members = gpsFmRead_('ACTIVITY_MEMBERS').filter(m => m.activity_id === activityId && m.active === true).map(m => String(m.user_id));
  const participantIds = activity.participant_mode === 'all_gps'
    ? gpsFmRead_('USERS').filter(u => u.active === true && ['guru','pj','superadmin'].indexOf(u.role)>=0).map(u => String(u.user_id))
    : members;
  participantIds.forEach(uid => gpsFmAppend_('ATTENDANCE_RECORDS',{record_id:Utilities.getUuid(),session_id:id,activity_id:activityId,user_id:uid,attendance_status:'belum_diisi',notes:'',recorded_by:s.userId,recorded_at:now,updated_at:now,updated_by:s.userId}));
  gpsFmAudit_(s.userId,'CREATE_ATTENDANCE_SESSION','ATTENDANCE_SESSIONS',id,{activityId:activityId,date:date,participantCount:participantIds.length});
  return {success:true,sessionId:id,participantCount:participantIds.length};
}
function apiGetSession(token, sessionId) {
  const s = gpsFmSession_(token);
  const session = gpsFmRead_('ATTENDANCE_SESSIONS').find(x => x.session_id === sessionId);
  if (!session) throw new Error('Sesi tidak ditemukan.');
  const activity = gpsFmRead_('ACTIVITIES').find(a => a.activity_id === session.activity_id);
  if (!activity || (!gpsFmCanManageActivity_(s,activity) && !['superadmin','administrasi'].includes(s.role) && s.role !== 'guru')) throw new Error('Tidak memiliki akses ke sesi ini.');
  const records = gpsFmRead_('ATTENDANCE_RECORDS').filter(r => r.session_id === sessionId).map(function(r) {
    const u = gpsFmUserById_(r.user_id);
    return {record_id:r.record_id,user_id:r.user_id,full_name:u.full_name,attendance_status:r.attendance_status,notes:r.notes};
  });
  return {success:true,session:session,activity:activity,records:records};
}
function apiSaveAttendance(token, sessionId, records) {
  const s = gpsFmSession_(token);
  if (!Array.isArray(records)) throw new Error('Data kehadiran tidak valid.');
  const session = gpsFmRead_('ATTENDANCE_SESSIONS').find(x => x.session_id === sessionId);
  if (!session) throw new Error('Sesi tidak ditemukan.');
  if (session.status === 'locked') throw new Error('Sesi sudah dikunci.');
  const activity = gpsFmRead_('ACTIVITIES').find(a => a.activity_id === session.activity_id);
  if (!activity || !gpsFmCanManageActivity_(s,activity)) throw new Error('Anda tidak berwenang mengubah kehadiran kegiatan ini.');
  const allowed = ['hadir','izin','sakit','alpa','belum_diisi'];
  const existing = gpsFmRead_('ATTENDANCE_RECORDS').filter(r => r.session_id === sessionId);
  const existingIds = new Set(existing.map(r => String(r.user_id)));
  const now = new Date();
  records.forEach(function(input) {
    const uid = String(input.user_id || ''), status = String(input.attendance_status || '');
    if (!existingIds.has(uid)) throw new Error('Peserta tidak terdaftar pada sesi ini.');
    if (allowed.indexOf(status) < 0) throw new Error('Status kehadiran tidak valid.');
    const rec = existing.find(r => String(r.user_id) === uid);
    gpsFmUpdateById_('ATTENDANCE_RECORDS','record_id',rec.record_id,{attendance_status:status,notes:String(input.notes||''),updated_at:now,updated_by:s.userId,recorded_by:s.userId,recorded_at:now});
  });
  gpsFmUpdateById_('ATTENDANCE_SESSIONS','session_id',sessionId,{status:'submitted'});
  gpsFmAudit_(s.userId,'SAVE_ATTENDANCE','ATTENDANCE_SESSIONS',sessionId,{recordCount:records.length});
  return {success:true,updated:records.length};
}
function apiLockAttendanceSession(token, sessionId, locked) {
  const s = gpsFmRequire_(token,['superadmin','administrasi']);
  const session = gpsFmRead_('ATTENDANCE_SESSIONS').find(x => x.session_id === sessionId);
  if (!session) throw new Error('Sesi tidak ditemukan.');
  gpsFmUpdateById_('ATTENDANCE_SESSIONS','session_id',sessionId,{status:locked?'locked':'submitted',locked_at:locked?new Date():'',locked_by:locked?s.userId:''});
  gpsFmAudit_(s.userId,locked?'LOCK_ATTENDANCE':'UNLOCK_ATTENDANCE','ATTENDANCE_SESSIONS',sessionId,{});
  return {success:true};
}

/** Monthly report data; can be exported by the UI to CSV. */
function apiGetMonthlyReport(token, yearMonth, activityId) {
  const s = gpsFmSession_(token);
  if (!/^\d{4}-\d{2}$/.test(String(yearMonth||''))) throw new Error('Format bulan harus YYYY-MM.');
  const sessions = gpsFmRead_('ATTENDANCE_SESSIONS').filter(x => String(x.session_date).slice(0,7) === yearMonth && x.status !== 'cancelled' && (!activityId || x.activity_id === activityId));
  const activities = gpsFmRead_('ACTIVITIES');
  const allRecords = gpsFmRead_('ATTENDANCE_RECORDS');
  const visibleSessions = sessions.filter(function(session) {
    const a = activities.find(x => x.activity_id === session.activity_id);
    if (!a) return false;
    return ['superadmin','administrasi'].includes(s.role) || s.role === 'guru' || gpsFmCanManageActivity_(s,a);
  });
  const rows = [];
  visibleSessions.forEach(function(session) {
    const a = activities.find(x => x.activity_id === session.activity_id);
    allRecords.filter(r => r.session_id === session.session_id).forEach(function(r) {
      if (s.role === 'guru' && r.user_id !== s.userId) return;
      const u = gpsFmUserById_(r.user_id);
      rows.push({tanggal:String(session.session_date).slice(0,10),kegiatan:a.name,pj:(gpsFmUserById_(a.pj_user_id)||{}).full_name||'',nama_guru:u.full_name,status:r.attendance_status,catatan:r.notes||''});
    });
  });
  gpsFmAudit_(s.userId,'VIEW_MONTHLY_REPORT','MONTHLY_REPORTS',yearMonth,{activityId:activityId||'',rowCount:rows.length});
  return {success:true,yearMonth:yearMonth,rows:rows,summary:{total:rows.length,hadir:rows.filter(r=>r.status==='hadir').length,izin:rows.filter(r=>r.status==='izin').length,sakit:rows.filter(r=>r.status==='sakit').length,alpa:rows.filter(r=>r.status==='alpa').length,belumDiisi:rows.filter(r=>r.status==='belum_diisi').length}};
}
function apiGetAdminReport(token, yearMonth) {
  gpsFmRequire_(token,['superadmin','administrasi']);
  const report = apiGetMonthlyReport(token,yearMonth,'');
  const byActivity = {};
  report.rows.forEach(r => {
    if (!byActivity[r.kegiatan]) byActivity[r.kegiatan] = {kegiatan:r.kegiatan,total:0,hadir:0,izin:0,sakit:0,alpa:0,belum_diisi:0};
    const x=byActivity[r.kegiatan]; x.total++; if (x[r.status] !== undefined) x[r.status]++;
  });
  return {success:true,yearMonth:yearMonth,summary:report.summary,activities:Object.keys(byActivity).map(k=>byActivity[k]),rows:report.rows};
}

function gpsFmSession_(token, required) {
  token = String(token || '');
  if (!token) throw new Error('Sesi tidak valid. Silakan login kembali.');
  let raw = CacheService.getScriptCache().get('GPSFM_SESSION_' + token);
  if (!raw) raw = PropertiesService.getScriptProperties().getProperty('GPSFM_SESSION_' + token);
  if (!raw) throw new Error('Sesi berakhir. Silakan login kembali.');
  const s = JSON.parse(raw);
  if (Date.now() > Number(s.expiresAt)) {
    CacheService.getScriptCache().remove('GPSFM_SESSION_' + token);
    PropertiesService.getScriptProperties().deleteProperty('GPSFM_SESSION_' + token);
    throw new Error('Sesi berakhir. Silakan login kembali.');
  }
  const user = gpsFmUserById_(s.userId);
  if (!user || user.active !== true) throw new Error('Akun tidak aktif.');
  if (required && required.indexOf(s.role) < 0) throw new Error('Akses ditolak.');
  return s;
}
function gpsFmRequire_(token, roles) { return gpsFmSession_(token,roles); }
function gpsFmUserById_(id) { return gpsFmRead_('USERS').find(u => String(u.user_id) === String(id)); }
function gpsFmPublicUser_(u) {
  if (!u) return null;
  return {user_id:u.user_id,username:u.username,full_name:u.full_name,email:u.email,phone:u.phone,role:u.role,active:u.active,must_change_password:u.must_change_password};
}
function gpsFmCanManageActivity_(s, activity) { return s.role === 'superadmin' || (s.role === 'pj' && String(activity.pj_user_id) === String(s.userId)); }
function gpsFmVisibleActivities_(s) {
  const activities = gpsFmRead_('ACTIVITIES').filter(a => a.active === true);
  if (s.role === 'superadmin' || s.role === 'administrasi') return activities;
  if (s.role === 'pj') return activities.filter(a => String(a.pj_user_id) === String(s.userId));
  return activities.filter(a => gpsFmRead_('ACTIVITY_MEMBERS').some(m => m.activity_id === a.activity_id && String(m.user_id) === String(s.userId) && m.active === true));
}
function gpsFmDashboard_(s) {
  const acts = gpsFmVisibleActivities_(s);
  const sessions = gpsFmRead_('ATTENDANCE_SESSIONS');
  const userRecords = gpsFmRead_('ATTENDANCE_RECORDS').filter(r => s.role !== 'guru' || String(r.user_id) === String(s.userId));
  return {activityCount:acts.length,sessionCount:sessions.filter(x => acts.some(a => a.activity_id === x.activity_id)).length,attendanceRecordCount:userRecords.length};
}
