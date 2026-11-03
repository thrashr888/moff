import { test } from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';

test('local clock resolves next weekdays and DST-aware midnight boundaries', () => {
  const moduleURL = new URL('../dist/durable/context.js', import.meta.url).href;
  const output = execFileSync(process.execPath, ['--input-type=module', '-e', `
    import { currentTime } from ${JSON.stringify(moduleURL)};
    console.log(JSON.stringify([
      currentTime(new Date('2026-10-17T19:30:00-07:00')),
      currentTime(new Date('2026-10-19T12:00:00-07:00')),
      currentTime(new Date('2026-03-07T12:00:00-08:00')),
      currentTime(new Date('2026-10-31T12:00:00-07:00')),
    ]));
  `], { encoding: 'utf8', env: { ...process.env, TZ: 'America/Los_Angeles' } });
  const [saturday, monday, spring, fall] = JSON.parse(output);
  assert.equal(saturday.timeZone, 'America/Los_Angeles');
  assert.equal(saturday.utcOffsetMinutes, -420);
  assert.equal(saturday.utc, '2026-10-18T02:30:00.000Z');
  assert.deepEqual(saturday.upcomingWeekdays.Monday, {
    start: '2026-10-19T00:00:00-07:00', endExclusive: '2026-10-20T00:00:00-07:00',
  });
  assert.equal(monday.upcomingWeekdays.Monday.start, '2026-10-26T00:00:00-07:00');
  assert.deepEqual(spring.upcomingWeekdays.Sunday, {
    start: '2026-03-08T00:00:00-08:00', endExclusive: '2026-03-09T00:00:00-07:00',
  });
  assert.deepEqual(fall.upcomingWeekdays.Sunday, {
    start: '2026-11-01T00:00:00-07:00', endExclusive: '2026-11-02T00:00:00-08:00',
  });
});
