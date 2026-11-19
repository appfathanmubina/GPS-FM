const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '../apps-script/Code.gs'), 'utf8');
const start = source.indexOf('function gpsfmRunScheduledPushNotifications()');
const end = source.indexOf('\nfunction previewGpsFmPushRecipients()', start);
assert.notEqual(start, -1, 'scheduled push handler must exist');
assert.notEqual(end, -1, 'scheduled push handler must have a clear end');
const handler = source.slice(start, end);

test('Apps Script backend source remains syntactically valid JavaScript', () => {
  assert.doesNotThrow(() => new vm.Script(source));
});

test('scheduled push processing has a soft execution deadline and defers remaining activities', () => {
  assert.match(handler, /const deadline = Date\.now\(\) \+ 270000/);
  assert.match(handler, /if \(Date\.now\(\) >= deadline\)/);
  assert.match(handler, /results\.deferredActivities = activities\.length - activityIndex/);
  assert.match(handler, /for \(let activityIndex = 0;/);
});

test('each OneSignal delivery is logged immediately instead of waiting for the entire run', () => {
  const sendAt = handler.indexOf('gpsfmSendOneSignal_(');
  const persistAt = handler.indexOf('deliverySheet.getRange(nextDeliveryRow');
  assert.ok(sendAt >= 0 && persistAt > sendAt, 'delivery result must be persisted after the OneSignal request');
  assert.match(handler, /deliverySheet\.getRange\(nextDeliveryRow,1,logRows\.length,deliveryHeaders\.length\)\.setValues\(logRows\)/);
  assert.doesNotMatch(handler, /pendingLogs/);
});

test('activity assignments are indexed once rather than rescanned for every activity', () => {
  assert.match(handler, /teacherAssignmentsByActivity/);
  assert.match(handler, /pjAssignmentsByActivity/);
  assert.doesNotMatch(handler, /teacherAssignments\.filter/);
  assert.doesNotMatch(handler, /pjAssignments\.forEach/);
});

test('skipped activities do not terminate the entire scheduled push run', () => {
  assert.match(handler, /if \(startMinute === null\) \{ results\.skippedNoTime\+\+; continue; \}/);
  assert.match(handler, /if \(\['holiday','cancelled'\]\.indexOf\(occurrenceStatus\) >= 0\) continue;/);
});
