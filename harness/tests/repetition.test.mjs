import { test } from 'node:test';
import assert from 'node:assert/strict';
import { hasRepetitionLoop } from '../dist/ollama/repetition.js';

test('repeated numbered entries and paragraphs are loops', () => {
  const line = 'A substantial repeated message from the same sender with identical details';
  assert.equal(hasRepetitionLoop(Array.from({ length: 6 }, (_, i) => `${i + 1}. ${line}`).join('\n')), true);
  assert.equal(hasRepetitionLoop((line + ' ').repeat(6)), true);
  assert.equal(hasRepetitionLoop((line + ' ').repeat(5)), false);
});

test('long varied answers and similar entries with distinct dates are allowed', () => {
  const answer = Array.from({ length: 400 }, (_, i) =>
    `${i + 1}. Report for record ${i}: this email has independent details and timestamp 2026-10-13T10:${String(i).padStart(3, '0')}:00Z`,
  ).join('\n');
  assert.ok(answer.length > 16_384);
  assert.equal(hasRepetitionLoop(answer), false);
  assert.equal(hasRepetitionLoop('x'.repeat(20_000)), false);
});

test('numbered George entries with invented changing timestamps trigger the guard', () => {
  const times = Array.from({ length: 8 }, (_, i) => `2026-10-12T22:${String(i).padStart(2, '0')}:00Z`);
  const output = times.map((time, i) =>
    `${i + 1}. "George Simmons" from George Simmons, sent on ${time}\n`,
  ).join('');
  const source = JSON.stringify(times.slice(0, 2));
  assert.equal(hasRepetitionLoop(output, source), true);
});

test('similar records with timestamps present in tools are not treated as an invented loop', () => {
  const times = Array.from({ length: 12 }, (_, i) => `2026-10-12T22:${String(i).padStart(2, '0')}:00Z`);
  const output = times.map((time, i) =>
    `${i + 1}. "George Simmons" from George Simmons, sent on ${time}\n`,
  ).join('');
  assert.equal(hasRepetitionLoop(output, JSON.stringify(times)), false);
});
