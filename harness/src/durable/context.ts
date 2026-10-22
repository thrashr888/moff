import { Type } from '@earendil-works/pi-ai';
import { defineExtension, defineTool, section } from '@earendil-works/pi-durable';

export function currentTime(now = new Date()) {
  const localISO = (date: Date) => {
    const offset = -date.getTimezoneOffset();
    const local = new Date(date.getTime() + offset * 60_000).toISOString().slice(0, 19);
    const sign = offset >= 0 ? '+' : '-';
    return `${local}${sign}${String(Math.floor(Math.abs(offset) / 60)).padStart(2, '0')}:${String(Math.abs(offset) % 60).padStart(2, '0')}`;
  };
  const upcomingWeekdays = Object.fromEntries(
    ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'].map((weekday, index) => {
      const start = new Date(now);
      start.setDate(start.getDate() + ((index - now.getDay() + 7) % 7 || 7));
      start.setHours(0, 0, 0, 0);
      const end = new Date(start);
      end.setDate(end.getDate() + 1);
      return [weekday, { start: localISO(start), endExclusive: localISO(end) }];
    }),
  );
  return {
    utc: now.toISOString(),
    local: new Intl.DateTimeFormat('en-CA', {
      dateStyle: 'full', timeStyle: 'long',
    }).format(now),
    timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone,
    utcOffsetMinutes: -now.getTimezoneOffset(),
    upcomingWeekdays,
  };
}

export const LocalContext = defineExtension({
  name: 'local-context',
  sections: [section('local-context', () =>
    'For relative dates such as today or Monday, call current_time first to get the actual date and timezone. ' +
    'Never guess a year from training data. current_time.upcomingWeekdays gives exact start/endExclusive timestamps: ' +
    'for Monday copy upcomingWeekdays.Monday.start and .endExclusive into the calendar range, not a week-long range. ' +
    'For an upcoming weekday use its next occurrence in the local timezone, ' +
    'and state the resolved date in your answer; ask if the intended date is ambiguous. ' +
    'Calendar workflow, in order: ' +
    '(1) current_time; (2) moff {"args":["outlook","calendar","folders"]}; ' +
    '(3) moff {"args":["outlook","calendar","view","--help"]}; ' +
    '(4) moff calendar view with a FolderId.Id copied from step 2 and the weekday timestamps from step 1. ' +
    'Never invent folder IDs such as primary. If one part of a request is blocked, answer the parts you completed, ' +
    'then explain the remaining question or error. A rejected ID is recoverable: discover folders and use the returned ID, ' +
    'rather than stopping or asking the user to provide an ID. Do not retry an unchanged invalid tool call. ' +
    'Before your final answer, check every part of the user request. Include the exact Subject of each requested email ' +
    'and each returned calendar event; do not replace source subjects with placeholders or paraphrases. ' +
    'Keep email results even if calendar lookup fails.')],
  tools: [defineTool({
    name: 'current_time',
    description: 'Get the current date, local timezone, and UTC offset. Use before resolving relative dates or scheduling reminders.',
    parameters: Type.Object({}, { additionalProperties: false }),
    replay: 'safe',
    execute: async () => ({ output: [{ type: 'text', text: JSON.stringify(currentTime()) }] }),
  })],
});
