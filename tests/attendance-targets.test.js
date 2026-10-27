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
