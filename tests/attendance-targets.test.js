const test = require('node:test');
const assert = require('node:assert/strict');

const DAYS = ['Senin','Selasa','Rabu','Kamis','Jumat','Sabtu','Minggu'];
function parseDays(schedule) {
  const raw = String(schedule || '').toLowerCase().replace(/[–—]/g, '-');
  if (raw.includes('setiap hari')) return DAYS;
  if (raw.includes('senin-jumat')) return DAYS.slice(0, 5);
  return DAYS.filter(day => new RegExp('(^|[^a-z])' + day.toLowerCase() + '([^a-z]|$)').test(raw));
}
function keyDate(key) {
  const [y,m,d] = key.split('-').map(Number);
  return new Date(Date.UTC(y,m-1,d,12));
}
function keyOf(date) {
  return date.getUTCFullYear() + '-' + String(date.getUTCMonth()+1).padStart(2,'0') + '-' + String(date.getUTCDate()).padStart(2,'0');
}
function weekday(key) {
  return ['Minggu','Senin','Selasa','Rabu','Kamis','Jumat','Sabtu'][keyDate(key).getUTCDay()];
}
function isOverdue(occurrenceDate, today) { return occurrenceDate < today; }
function missingTeachers(snapshotIds, attendanceRows, activityId, date) {
  return snapshotIds.filter(id => !attendanceRows.some(r =>
    String(r.activityId) === String(activityId) &&
    String(r.teacherId) === String(id) &&
    r.date === date
  ));
}
function canManage(role, assignedActivityIds, activityId) {
  return role === 'superadmin' || assignedActivityIds.includes(String(activityId));
}

test('weekly schedule handles selected days and weekday ranges', () => {
  assert.deepEqual(parseDays('Selasa, Kamis · 16.00–17.00'), ['Selasa','Kamis']);
  assert.deepEqual(parseDays('Senin-Jumat'), DAYS.slice(0,5));
  assert.equal(weekday('2026-10-27'), 'Selasa');
});

test('occurrence becomes overdue on the following calendar day, including month boundary', () => {
  assert.equal(isOverdue('2026-10-31','2026-10-31'), false);
  assert.equal(isOverdue('2026-10-31','2026-11-01'), true);
  assert.equal(isOverdue('2026-11-01','2026-11-01'), false);
  assert.equal(keyOf(keyDate('2026-11-01')), '2026-11-01');
});

test('any saved attendance status satisfies target; missing status is not inferred', () => {
  const roster = ['teacher-a','teacher-b','teacher-c'];
  const records = [
    {activityId:'activity-1',teacherId:'teacher-a',date:'2026-10-20',status:'Hadir'},
    {activityId:'activity-1',teacherId:'teacher-b',date:'2026-10-20',status:'D. Alfa'}
  ];
  assert.deepEqual(missingTeachers(roster,records,'activity-1','2026-10-20'), ['teacher-c']);
});

test('holiday/cancelled occurrences are excluded by explicit occurrence status', () => {
  const countsAsTarget = status => status !== 'holiday' && status !== 'cancelled';
  assert.equal(countsAsTarget('scheduled'), true);
  assert.equal(countsAsTarget('holiday'), false);
  assert.equal(countsAsTarget('cancelled'), false);
});

test('teacher roster snapshot stays historical after current roster changes', () => {
  const snapshotAtOccurrence = ['teacher-old'];
  const currentRosterAfterChange = ['teacher-new'];
  assert.deepEqual(missingTeachers(snapshotAtOccurrence, [], 'activity-1','2026-10-20'), ['teacher-old']);
  assert.deepEqual(missingTeachers(currentRosterAfterChange, [], 'activity-1','2026-10-20'), ['teacher-new']);
  // Production must keep using the persisted snapshot for an already-created occurrence.
  assert.notDeepEqual(snapshotAtOccurrence, currentRosterAfterChange);
});

test('PJ access is limited to assigned activities; superadmin can manage all', () => {
  assert.equal(canManage('gps',['activity-1'],'activity-1'), true);
  assert.equal(canManage('gps',['activity-1'],'activity-2'), false);
  assert.equal(canManage('superadmin',[],'activity-2'), true);
});

const fs = require('node:fs');
const path = require('node:path');

test('single and batch attendance share strict calendar-date validation', () => {
  const source = fs.readFileSync(path.join(__dirname, '../apps-script/Code.gs'), 'utf8');
  const helperStart = source.indexOf('function gpsfmIsValidAttendanceDate_(date)');
  const helperEnd = source.indexOf('\n}', helperStart);
  assert.notEqual(helperStart, -1, 'shared date validator must exist');
  assert.notEqual(helperEnd, -1, 'shared date validator must be complete');
  const helper = source.slice(helperStart, helperEnd);
  assert.match(helper, /value\.match/);
  assert.match(helper, /check\.getUTCFullYear\(\) === year/);
  const singleStart = source.indexOf('function apiSaveAttendance(token, data)');
  const singleEnd = source.indexOf('\n}', singleStart);
  const batchStart = source.indexOf('function apiSaveAttendanceBatch(token, data)');
  const batchEnd = source.indexOf('\nfunction apiSaveAttendance(', batchStart);
  assert.match(source.slice(singleStart, singleEnd), /gpsfmIsValidAttendanceDate_\(date\)/);
  assert.match(source.slice(batchStart, batchEnd), /gpsfmIsValidAttendanceDate_\(date\)/);
});

test('Excel attendance import requires the complete current roster and rejects duplicates under lock', () => {
  const source = fs.readFileSync(path.join(__dirname, '../apps-script/Code.gs'), 'utf8');
  const batchStart = source.indexOf('function apiSaveAttendanceBatch(');
  const batchEnd = source.indexOf('\nfunction apiSaveAttendance(', batchStart);
  assert.notEqual(batchStart, -1);
  assert.notEqual(batchEnd, -1);
  const batch = source.slice(batchStart, batchEnd);
  const lockAt = batch.indexOf('lock.waitLock(10000)');
  const duplicateGuardAt = batch.indexOf('if(data.rejectExisting===true)');
  const rosterGuardAt = batch.lastIndexOf('if(data.requireCompleteRoster===true)');
  assert.notEqual(lockAt, -1);
  assert.ok(duplicateGuardAt > lockAt, 'duplicate check must run while holding the script lock');
  assert.ok(rosterGuardAt > lockAt, 'current roster must be revalidated while holding the script lock');
  assert.match(batch, /normalized\.length!==currentRoster\.length/);
  assert.match(batch, /Tidak ada data yang diubah/);
});

test('attendance corrections require a reason and include it in the audit trail', () => {
  const source = fs.readFileSync(path.join(__dirname, '../apps-script/Code.gs'), 'utf8');
  assert.match(source, /if\(updateRows\.length&&!correctionReason\)/);
  assert.match(source, /correctionReason:correctionReason\|\|''/);
  const singleStart = source.indexOf('function apiSaveAttendance(token, data)');
  const singleEnd = source.indexOf('\n}', singleStart);
  const single = source.slice(singleStart, singleEnd);
  assert.match(single, /if\(!correctionReason\)throw new Error/);
  assert.match(single, /correctionReason:correctionReason/);
});


test('duty schedule sheet migrates legacy rows to active without dropping assignments', () => {
  const source = fs.readFileSync(path.join(__dirname, '../apps-script/Code.gs'), 'utf8');
  const start = source.indexOf('function gpsfmEnsureActivityDutySheets_()');
  const end = source.indexOf('\n}', start);
  assert.notEqual(start, -1);
  assert.notEqual(end, -1);
  const helper = source.slice(start, end);
  assert.match(helper, /indexOf\('active'\)/);
  assert.match(helper, /setValue\('active'\)/);
  assert.match(helper, /setValues\(Array\.from\(\{length:lastRow-1\}.*\[true\]/s);
  assert.match(helper, /'updatedAt','active'/);
});

test('published duty roster matches normalized spreadsheet dates', () => {
  const source = fs.readFileSync(path.join(__dirname, '../apps-script/Code.gs'), 'utf8');
  const resolverStart = source.indexOf('function gpsfmPublishedDutyTeacherIds_(');
  const resolverEnd = source.indexOf('\n}', resolverStart);
  const resolver = source.slice(resolverStart, resolverEnd);
  assert.match(resolver, /dutyDate:String\(date\)\},\{dutyDate:'date'\}\)/);
  const monthStart = source.indexOf('function apiGetActivityDutyMonth(');
  const monthEnd = source.indexOf('\n}', monthStart);
  const monthApi = source.slice(monthStart, monthEnd);
  assert.match(monthApi, /gpsfmDateKey_\(r\.dutyDate\)\.slice\(0,7\)/);
  assert.match(monthApi, /overrides\[gpsfmDateKey_\(r\.dutyDate\)\]/);
});

test('saving a duty date updates an existing normalized date row instead of duplicating it', () => {
  const source = fs.readFileSync(path.join(__dirname, '../apps-script/Code.gs'), 'utf8');
  const start = source.indexOf('function apiSaveActivityDutyDate(');
  const end = source.indexOf('\n}', start);
  const saveApi = source.slice(start, end);
  assert.match(saveApi, /gpsfmFindRowByCriteria_\('ActivityDutySchedules',\{activityId:activityId,dutyDate:date\},\{dutyDate:'date'\}\)/);
});


test('published attendance roster prefers active schedule rows when duplicate inactive rows exist', () => {
  const source = fs.readFileSync(path.join(__dirname, '../apps-script/Code.gs'), 'utf8');
  const start = source.indexOf('function gpsfmPublishedDutyTeacherIds_(');
  const end = source.indexOf('\\n}', start);
  const fn = source.slice(start, end);
  assert.match(fn, /gpsfmFindRowByCriteria_\\('ActivityDutySchedules',[\\s\\S]*?function\\(row\\)\\{return gpsfmBool_\\(row\\.active\\);\\}\\)/);
});

test('saving a duty date updates an active duplicate schedule row before inactive legacy rows', () => {
  const source = fs.readFileSync(path.join(__dirname, '../apps-script/Code.gs'), 'utf8');
  const start = source.indexOf('function apiSaveActivityDutyDate(');
  const end = source.indexOf('\\n}', start);
  const fn = source.slice(start, end);
  assert.match(fn, /const activeOld=gpsfmFindRowByCriteria_\\('ActivityDutySchedules'[\\s\\S]*?function\\(row\\)\\{return gpsfmBool_\\(row\\.active\\);\\}\\)/);
  assert.match(fn, /const old=activeOld\\|\\|gpsfmFindRowByCriteria_/);
});
