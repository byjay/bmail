import type { CalendarScheduleEvent } from './deterministicAnalyzer';

function escapeIcsText(raw: string): string {
  return String(raw || '')
    .replace(/\\/g, '\\\\')
    .replace(/\r\n|\r|\n/g, '\\n')
    .replace(/;/g, '\\;')
    .replace(/,/g, '\\,');
}

function unescapeIcsText(raw: string): string {
  return String(raw || '')
    .replace(/\\n/gi, '\n')
    .replace(/\\,/g, ',')
    .replace(/\\;/g, ';')
    .replace(/\\\\/g, '\\')
    .trim();
}

function formatIcsDateTime(dateStr: string, timeLabel?: string): { dtStart: string; dtEnd: string; isAllDay: boolean } {
  const cleanDate = (dateStr || '').replace(/[^0-9]/g, '');
  const ymd = cleanDate.length >= 8 ? cleanDate.slice(0, 8) : new Date().toISOString().slice(0, 10).replace(/-/g, '');
  const m = /^(\d{1,2}):(\d{2})$/.exec((timeLabel || '').trim());
  if (!m) {
    return {
      dtStart: `;VALUE=DATE:${ymd}`,
      dtEnd: `;VALUE=DATE:${ymd}`,
      isAllDay: true,
    };
  }
  const hh = String(Math.min(23, Math.max(0, Number(m[1])))).padStart(2, '0');
  const mm = String(Math.min(59, Math.max(0, Number(m[2])))).padStart(2, '0');
  const endHourNum = Math.min(23, Number(hh) + 1);
  const endHh = String(endHourNum).padStart(2, '0');
  return {
    dtStart: `:${ymd}T${hh}${mm}00`,
    dtEnd: `:${ymd}T${endHh}${mm}00`,
    isAllDay: false,
  };
}

function foldIcsLine(line: string): string {
  if (line.length <= 74) return line;
  const parts: string[] = [];
  let cursor = 0;
  while (cursor < line.length) {
    const chunkLen = cursor === 0 ? 74 : 73;
    const chunk = line.slice(cursor, cursor + chunkLen);
    parts.push(cursor === 0 ? chunk : ` ${chunk}`);
    cursor += chunkLen;
  }
  return parts.join('\r\n');
}

export function exportEventsToIcs(
  events: CalendarScheduleEvent[],
  calendarName = 'Bmail Outlook AI 일정'
): string {
  const dtStamp = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d+Z$/, 'Z');
  const lines: string[] = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//Bmail Outlook AI//Calendar Adapter v1.2.2//KO',
    'CALSCALE:GREGORIAN',
    'METHOD:PUBLISH',
    `X-WR-CALNAME:${escapeIcsText(calendarName)}`,
  ];

  events.forEach((ev, idx) => {
    const { dtStart, dtEnd } = formatIcsDateTime(ev.date, ev.timeLabel);
    const uid = `${(ev.id || `bmail-ev-${idx}`).replace(/[^a-zA-Z0-9._-]/g, '_')}@bmail.seastar.work`;
    const categories = [ev.category, ev.projectCode].filter(Boolean).join(',');
    const descriptionParts = [
      ev.notes || '',
      ev.sourceEmailSubject ? `원본 메일: ${ev.sourceEmailSubject}` : '',
      ev.projectCode ? `프로젝트: ${ev.projectCode}` : '',
    ].filter(Boolean);

    lines.push('BEGIN:VEVENT');
    lines.push(`UID:${uid}`);
    lines.push(`DTSTAMP:${dtStamp}`);
    lines.push(`DTSTART${dtStart}`);
    lines.push(`DTEND${dtEnd}`);
    lines.push(`SUMMARY:${escapeIcsText(ev.title || '일정')}`);
    if (descriptionParts.length > 0) {
      lines.push(`DESCRIPTION:${escapeIcsText(descriptionParts.join('\n'))}`);
    }
    if (categories) {
      lines.push(`CATEGORIES:${escapeIcsText(categories)}`);
    }
    lines.push(`STATUS:${ev.completed ? 'COMPLETED' : 'CONFIRMED'}`);
    lines.push('END:VEVENT');
  });

  lines.push('END:VCALENDAR');
  return `${lines.map(foldIcsLine).join('\r\n')}\r\n`;
}

function parseIcsDateProperty(rawValue: string): { date: string; timeLabel: string } {
  const clean = rawValue.trim();
  const match = /(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2}))?/.exec(clean);
  if (!match) {
    return {
      date: new Date().toISOString().slice(0, 10),
      timeLabel: '09:00',
    };
  }
  const date = `${match[1]}-${match[2]}-${match[3]}`;
  const timeLabel = match[4] && match[5] ? `${match[4]}:${match[5]}` : '09:00';
  return { date, timeLabel };
}

export function parseIcsToEvents(icsText: string): Array<Omit<CalendarScheduleEvent, 'id'>> {
  if (!icsText || typeof icsText !== 'string') return [];

  // RFC 5545 line unfolding (CRLF/LF followed by single space or tab)
  const unfolded = icsText.replace(/\r?\n[ \t]/g, '');
  const rawLines = unfolded.split(/\r?\n/);

  const parsed: Array<Omit<CalendarScheduleEvent, 'id'>> = [];
  let inEvent = false;
  let current: Partial<{
    summary: string;
    dtstart: string;
    description: string;
    categories: string;
    status: string;
  }> = {};

  for (const rawLine of rawLines) {
    const line = rawLine.trim();
    if (!line) continue;
    const upper = line.toUpperCase();
    if (upper === 'BEGIN:VEVENT') {
      inEvent = true;
      current = {};
      continue;
    }
    if (upper === 'END:VEVENT') {
      if (inEvent && current.summary && current.dtstart) {
        const { date, timeLabel } = parseIcsDateProperty(current.dtstart);
        const cats = (current.categories || '')
          .split(',')
          .map((c) => unescapeIcsText(c))
          .filter(Boolean);

        let category: CalendarScheduleEvent['category'] = 'meeting';
        let projectCode: string | undefined;
        for (const cat of cats) {
          if (cat === 'deadline' || cat === 'action-item' || cat === 'email-extracted' || cat === 'meeting') {
            category = cat;
          } else if (!projectCode && /^[A-Z0-9_-]{2,20}$/i.test(cat)) {
            projectCode = cat;
          }
        }
        const desc = current.description ? unescapeIcsText(current.description) : undefined;
        parsed.push({
          title: unescapeIcsText(current.summary),
          date,
          timeLabel,
          category,
          projectCode,
          notes: desc,
          completed: (current.status || '').toUpperCase() === 'COMPLETED',
        });
      }
      inEvent = false;
      current = {};
      continue;
    }

    if (!inEvent) continue;
    const colonIdx = line.indexOf(':');
    if (colonIdx <= 0) continue;
    const propHead = line.slice(0, colonIdx).toUpperCase();
    const propName = propHead.split(';')[0];
    const propValue = line.slice(colonIdx + 1);

    if (propName === 'SUMMARY') current.summary = propValue;
    else if (propName === 'DTSTART') current.dtstart = propValue;
    else if (propName === 'DESCRIPTION') current.description = propValue;
    else if (propName === 'CATEGORIES') current.categories = propValue;
    else if (propName === 'STATUS') current.status = propValue;
  }

  return parsed;
}
