import assert from 'node:assert/strict';
import test from 'node:test';
import { exportEventsToIcs, parseIcsToEvents } from './icsCalendarAdapter';
import type { CalendarScheduleEvent } from './deterministicAnalyzer';

test('exportEventsToIcs and parseIcsToEvents round-trip Korean schedules, escaping, and statuses', () => {
  const sampleEvents: CalendarScheduleEvent[] = [
    {
      id: 'ev-1',
      title: 'H2710 도면 출도 회의; 최종 검토, 2차',
      date: '2026-11-25',
      timeLabel: '14:30',
      category: 'meeting',
      projectCode: 'H2710',
      notes: '참석: 설계1팀\n안건: 파이프 스풀 확정',
      completed: false,
    },
    {
      id: 'ev-2',
      title: '선급 승인 기한 마감',
      date: '2026-11-30',
      timeLabel: '09:00',
      category: 'deadline',
      completed: true,
    },
  ];

  const ics = exportEventsToIcs(sampleEvents);
  assert.ok(ics.includes('BEGIN:VCALENDAR'));
  assert.ok(ics.includes('END:VCALENDAR'));
  assert.ok(ics.includes('DTSTART:20261125T143000'));
  assert.ok(ics.includes('STATUS:COMPLETED'));

  const imported = parseIcsToEvents(ics);
  assert.equal(imported.length, 2);
  assert.equal(imported[0].title, 'H2710 도면 출도 회의; 최종 검토, 2차');
  assert.equal(imported[0].date, '2026-11-25');
  assert.equal(imported[0].timeLabel, '14:30');
  assert.equal(imported[0].category, 'meeting');
  assert.equal(imported[0].projectCode, 'H2710');
  assert.equal(imported[0].completed, false);
  assert.ok(imported[0].notes?.includes('파이프 스풀 확정'));

  assert.equal(imported[1].title, '선급 승인 기한 마감');
  assert.equal(imported[1].category, 'deadline');
  assert.equal(imported[1].completed, true);
});

test('parseIcsToEvents handles RFC 5545 folded lines and recovers cleanly from malformed blocks', () => {
  const foldedIcs = [
    'BEGIN:VCALENDAR',
    'BEGIN:VEVENT',
    'DTSTART;VALUE=DATE:20261201',
    'SUMMARY:아주 긴 제목이 여러 줄로 접혀서',
    ' 들어온 경우 테스트',
    'END:VEVENT',
    'BEGIN:VEVENT',
    'SUMMARY:DTSTART 누락 이벤트',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  const events = parseIcsToEvents(foldedIcs);
  assert.equal(events.length, 1);
  assert.equal(events[0].title, '아주 긴 제목이 여러 줄로 접혀서들어온 경우 테스트');
  assert.equal(events[0].date, '2026-12-01');
  assert.equal(events[0].timeLabel, '09:00');
  assert.deepEqual(parseIcsToEvents('not-an-ics-file'), []);
});
