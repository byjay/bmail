import React, { useMemo, useRef, useState } from 'react';
import {
  Calendar as CalendarIcon,
  CheckCircle2,
  ChevronLeft,
  ChevronRight,
  Circle,
  Download,
  Mic,
  Plus,
  Trash2,
  Upload,
} from 'lucide-react';
import type { CalendarScheduleEvent, ProjectDefinition } from '../utils/deterministicAnalyzer';
import { exportEventsToIcs, parseIcsToEvents } from '../utils/icsCalendarAdapter';

interface OutlookCalendarViewProps {
  events: CalendarScheduleEvent[];
  projects: ProjectDefinition[];
  onAddManualEvent: (event: Omit<CalendarScheduleEvent, 'id'>) => void;
  onImportIcsEvents?: (events: Array<Omit<CalendarScheduleEvent, 'id'>>) => number;
  onToggleCompleteEvent: (id: string) => void;
  onDeleteManualEvent: (id: string) => void;
  onOpenMeetingAnalyzer: (title: string, date: string) => void;
}

const WEEKDAYS = ['일', '월', '화', '수', '목', '금', '토'];

export const OutlookCalendarView: React.FC<OutlookCalendarViewProps> = ({
  events,
  projects,
  onAddManualEvent,
  onImportIcsEvents,
  onToggleCompleteEvent,
  onDeleteManualEvent,
  onOpenMeetingAnalyzer,
}) => {
  const now = new Date();
  const icsFileInputRef = useRef<HTMLInputElement | null>(null);
  const [icsNotice, setIcsNotice] = useState<string | null>(null);
  const [viewYear, setViewYear] = useState(now.getFullYear());
  const [viewMonth, setViewMonth] = useState(now.getMonth()); // 0-indexed
  const [selectedDate, setSelectedDate] = useState(now.toISOString().slice(0, 10));

  const [newTitle, setNewTitle] = useState('');
  const [newTime, setNewTime] = useState('10:00');
  const [newCategory, setNewCategory] = useState<CalendarScheduleEvent['category']>('meeting');
  const [newProjectCode, setNewProjectCode] = useState('');
  const [newNotes, setNewNotes] = useState('');

  const eventsByDate = useMemo(() => {
    const map = new Map<string, CalendarScheduleEvent[]>();
    for (const ev of events) {
      const arr = map.get(ev.date) ?? [];
      arr.push(ev);
      map.set(ev.date, arr);
    }
    return map;
  }, [events]);

  const calendarCells = useMemo(() => {
    const firstDay = new Date(viewYear, viewMonth, 1);
    const startWeekday = firstDay.getDay();
    const daysInMonth = new Date(viewYear, viewMonth + 1, 0).getDate();
    const cells: Array<{ dateStr: string; dayNum: number; inMonth: boolean }> = [];

    // Leading days from previous month
    const prevMonthDays = new Date(viewYear, viewMonth, 0).getDate();
    for (let i = startWeekday - 1; i >= 0; i -= 1) {
      const d = prevMonthDays - i;
      const dt = new Date(viewYear, viewMonth - 1, d);
      cells.push({
        dateStr: `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`,
        dayNum: d,
        inMonth: false,
      });
    }

    for (let d = 1; d <= daysInMonth; d += 1) {
      cells.push({
        dateStr: `${viewYear}-${String(viewMonth + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`,
        dayNum: d,
        inMonth: true,
      });
    }

    while (cells.length % 7 !== 0) {
      const idx = cells.length - (startWeekday + daysInMonth) + 1;
      const dt = new Date(viewYear, viewMonth + 1, idx);
      cells.push({
        dateStr: `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}-${String(dt.getDate()).padStart(2, '0')}`,
        dayNum: dt.getDate(),
        inMonth: false,
      });
    }

    return cells;
  }, [viewYear, viewMonth]);

  const selectedDayEvents = eventsByDate.get(selectedDate) ?? [];
  const todayStr = new Date().toISOString().slice(0, 10);

  const handleShiftMonth = (delta: number) => {
    const next = new Date(viewYear, viewMonth + delta, 1);
    setViewYear(next.getFullYear());
    setViewMonth(next.getMonth());
  };

  const handleAddSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newTitle.trim()) return;
    onAddManualEvent({
      title: newTitle.trim(),
      date: selectedDate,
      timeLabel: newTime || '09:00',
      category: newCategory,
      projectCode: newProjectCode || undefined,
      notes: newNotes.trim() || undefined,
      completed: false,
    });
    setNewTitle('');
    setNewNotes('');
  };

  const handleExportIcs = () => {
    const icsText = exportEventsToIcs(events);
    const blob = new Blob([icsText], { type: 'text/calendar;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `bmail-calendar-${new Date().toISOString().slice(0, 10)}.ics`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    setIcsNotice(`Windows/Outlook 호환 .ics 파일로 일정 ${events.length}건을 내보냈습니다.`);
  };

  const handleImportIcsFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;
    try {
      const text = await file.text();
      const imported = parseIcsToEvents(text);
      if (imported.length === 0) {
        setIcsNotice('선택한 .ics 파일에서 유효한 일정(VEVENT)을 찾지 못했습니다.');
        return;
      }
      const addedCount = onImportIcsEvents ? onImportIcsEvents(imported) : 0;
      setIcsNotice(
        `${file.name}에서 일정 ${imported.length}건 중 중복 제외 ${addedCount}건을 가져왔습니다.`
      );
    } catch {
      setIcsNotice('.ics 일정 파일을 읽는 중 오류가 발생했습니다.');
    }
  };

  const categoryColor = (cat: CalendarScheduleEvent['category']) => {
    switch (cat) {
      case 'deadline':
        return 'bg-red-100 text-red-800 border-red-200';
      case 'action-item':
        return 'bg-emerald-50 text-emerald-800 border-emerald-200';
      case 'email-extracted':
        return 'bg-[#EBF3FC] text-[#0F6CBD] border-[#B4D6FA]';
      default:
        return 'bg-slate-100 text-[#242424] border-[#E0E0E0]';
    }
  };

  return (
    <div className="flex-1 flex flex-col min-h-0 bg-[#F5F5F5] overflow-hidden">
      {/* Outlook Calendar Toolbar */}
      <div className="bg-white border-b border-[#E0E0E0] px-6 py-3 flex flex-wrap items-center justify-between gap-3 shrink-0">
        <div className="flex items-center gap-3">
          <CalendarIcon className="w-5 h-5 text-[#0F6CBD]" />
          <h1 className="text-base font-bold text-[#242424]">
            {viewYear}년 {viewMonth + 1}월 일정 · 스케줄 관리
          </h1>
          <div className="flex items-center gap-1 ml-2">
            <button
              type="button"
              onClick={() => handleShiftMonth(-1)}
              aria-label="이전 달"
              className="p-1.5 rounded hover:bg-[#F5F5F5] text-[#424242]"
            >
              <ChevronLeft className="w-4 h-4" />
            </button>
            <button
              type="button"
              onClick={() => {
                const t = new Date();
                setViewYear(t.getFullYear());
                setViewMonth(t.getMonth());
                setSelectedDate(todayStr);
              }}
              className="px-2.5 py-1 text-xs font-medium bg-white border border-[#D1D1D1] hover:bg-[#F5F5F5] rounded-[5px]"
            >
              오늘
            </button>
            <button
              type="button"
              onClick={() => handleShiftMonth(1)}
              aria-label="다음 달"
              className="p-1.5 rounded hover:bg-[#F5F5F5] text-[#424242]"
            >
              <ChevronRight className="w-4 h-4" />
            </button>
          </div>
        </div>
        <div className="flex flex-wrap items-center gap-2">
          <span className="text-xs text-[#616161]">
            메일 자동 추출 + Action Item + 직접 등록 ({events.length}건)
          </span>
          <button
            type="button"
            onClick={handleExportIcs}
            className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-medium text-[#242424] bg-white border border-[#D1D1D1] hover:bg-[#FAF9F8] rounded-[4px]"
          >
            <Download className="w-3.5 h-3.5 text-[#0F6CBD]" />
            .ics 내보내기
          </button>
          {onImportIcsEvents && (
            <>
              <button
                type="button"
                onClick={() => icsFileInputRef.current?.click()}
                className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-medium text-[#0F6CBD] bg-[#EBF3FC] hover:bg-[#DCEBFA] border border-[#B4D6FA] rounded-[4px]"
              >
                <Upload className="w-3.5 h-3.5" />
                .ics 가져오기
              </button>
              <input
                ref={icsFileInputRef}
                type="file"
                accept=".ics,text/calendar"
                onChange={(e) => void handleImportIcsFile(e)}
                className="hidden"
              />
            </>
          )}
        </div>
      </div>

      {icsNotice && (
        <div className="mx-3 mt-2 px-3 py-1.5 text-xs bg-[#EBF3FC] border border-[#B4D6FA] text-[#0F6CBD] rounded flex items-center justify-between gap-2 shrink-0">
          <span>{icsNotice}</span>
          <button
            type="button"
            onClick={() => setIcsNotice(null)}
            className="text-[11px] font-semibold hover:underline"
          >
            닫기
          </button>
        </div>
      )}

      {/* Main Split: Calendar Month Grid + Selected Day Agenda & Add Form */}
      <div className="flex-1 min-h-0 p-3 grid grid-cols-1 lg:grid-cols-12 gap-3 overflow-hidden">
        {/* Month Grid */}
        <section className="lg:col-span-8 bg-white border border-[#E0E0E0] rounded-lg flex flex-col overflow-hidden">
          <div className="grid grid-cols-7 border-b border-[#EDEBE9] bg-[#FAF9F8]">
            {WEEKDAYS.map((day, i) => (
              <div
                key={day}
                className={`py-2 text-center text-xs font-semibold ${
                  i === 0 ? 'text-red-600' : i === 6 ? 'text-[#0F6CBD]' : 'text-[#424242]'
                }`}
              >
                {day}
              </div>
            ))}
          </div>

          <div className="flex-1 grid grid-cols-7 auto-rows-fr divide-x divide-y divide-[#EDEBE9] overflow-y-auto">
            {calendarCells.map((cell) => {
              const dayEvents = eventsByDate.get(cell.dateStr) ?? [];
              const isSelected = cell.dateStr === selectedDate;
              const isToday = cell.dateStr === todayStr;

              return (
                <button
                  key={cell.dateStr}
                  type="button"
                  onClick={() => setSelectedDate(cell.dateStr)}
                  className={`p-1.5 text-left flex flex-col gap-1 min-h-[84px] transition-colors ${
                    isSelected
                      ? 'bg-[#EBF3FC] ring-1 ring-inset ring-[#0F6CBD]'
                      : cell.inMonth
                        ? 'bg-white hover:bg-[#FAF9F8]'
                        : 'bg-[#FAFAFA] text-[#9E9E9E]'
                  }`}
                >
                  <span
                    className={`w-6 h-6 text-xs font-mono tabular-nums flex items-center justify-center rounded-full ${
                      isToday
                        ? 'bg-[#0F6CBD] text-white font-bold'
                        : isSelected
                          ? 'text-[#0F6CBD] font-bold'
                          : 'text-[#242424]'
                    }`}
                  >
                    {cell.dayNum}
                  </span>
                  <div className="space-y-0.5 w-full overflow-hidden">
                    {dayEvents.slice(0, 2).map((ev) => (
                      <div
                        key={ev.id}
                        className={`px-1.5 py-0.5 text-[10px] rounded border truncate ${categoryColor(
                          ev.category
                        )}`}
                      >
                        {ev.timeLabel} {ev.title}
                      </div>
                    ))}
                    {dayEvents.length > 2 && (
                      <div className="text-[10px] text-[#0F6CBD] font-semibold pl-1">
                        +{dayEvents.length - 2}건 더 보기
                      </div>
                    )}
                  </div>
                </button>
              );
            })}
          </div>
        </section>

        {/* Right Agenda & Quick Schedule Creator */}
        <aside className="lg:col-span-4 bg-white border border-[#E0E0E0] rounded-lg flex flex-col overflow-hidden">
          <div className="px-4 py-3 border-b border-[#EDEBE9] flex items-center justify-between">
            <div>
              <h2 className="text-sm font-bold text-[#242424] font-mono">{selectedDate}</h2>
              <p className="text-[11px] text-[#616161]">선택한 날짜 일정 {selectedDayEvents.length}건</p>
            </div>
            <button
              type="button"
              onClick={() => onOpenMeetingAnalyzer(`${selectedDate} 회의`, selectedDate)}
              className="inline-flex items-center gap-1 px-2.5 py-1.5 text-[11px] font-semibold text-[#0F6CBD] bg-[#EBF3FC] hover:bg-[#CFE4FA] rounded-[5px]"
            >
              <Mic className="w-3.5 h-3.5" />
              음성·회의록 연동
            </button>
          </div>

          <div className="flex-1 overflow-y-auto p-4 space-y-2.5">
            {selectedDayEvents.length === 0 ? (
              <div className="p-5 rounded-lg bg-[#FAF9F8] border border-[#EDEBE9] text-center text-xs text-[#616161]">
                이 날짜에 등록되거나 메일에서 감지된 일정이 없습니다.
              </div>
            ) : (
              selectedDayEvents.map((ev) => (
                <div
                  key={ev.id}
                  className={`p-3 rounded-lg border space-y-1.5 ${categoryColor(ev.category)}`}
                >
                  <div className="flex items-start justify-between gap-2">
                    <button
                      type="button"
                      onClick={() => onToggleCompleteEvent(ev.id)}
                      className="flex items-start gap-2 text-left min-w-0"
                    >
                      {ev.completed ? (
                        <CheckCircle2 className="w-4 h-4 shrink-0 mt-0.5 text-emerald-700" />
                      ) : (
                        <Circle className="w-4 h-4 shrink-0 mt-0.5 opacity-60" />
                      )}
                      <div className="min-w-0">
                        <div
                          className={`text-xs font-bold ${ev.completed ? 'line-through opacity-60' : ''}`}
                        >
                          {ev.title}
                        </div>
                        <div className="text-[11px] font-mono opacity-80">
                          {ev.timeLabel}
                          {ev.projectCode ? ` · [${ev.projectCode}]` : ''}
                        </div>
                      </div>
                    </button>
                    <div className="flex items-center gap-1 shrink-0">
                      <button
                        type="button"
                        onClick={() => onOpenMeetingAnalyzer(ev.title, ev.date)}
                        title="이 일정으로 음성·회의록 분석 열기"
                        className="p-1 rounded hover:bg-black/5"
                      >
                        <Mic className="w-3.5 h-3.5" />
                      </button>
                      {ev.category === 'manual' && (
                        <button
                          type="button"
                          onClick={() => onDeleteManualEvent(ev.id)}
                          aria-label="일정 삭제"
                          className="p-1 rounded hover:bg-black/5"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      )}
                    </div>
                  </div>
                  {ev.notes && <p className="text-[11px] opacity-85 leading-snug">{ev.notes}</p>}
                </div>
              ))
            )}
          </div>

          {/* Add New Schedule Event Form */}
          <form onSubmit={handleAddSubmit} className="p-4 border-t border-[#EDEBE9] bg-[#FAF9F8] space-y-2.5">
            <div className="text-xs font-bold text-[#242424]">새 일정 · 회의 추가 ({selectedDate})</div>
            <input
              type="text"
              value={newTitle}
              onChange={(e) => setNewTitle(e.target.value)}
              placeholder="일정 제목 (예: 고객사 견적 검토 미팅)"
              className="w-full px-2.5 py-1.5 text-xs bg-white border border-[#D1D1D1] rounded-[5px]"
            />
            <div className="grid grid-cols-3 gap-2">
              <input
                type="time"
                value={newTime}
                onChange={(e) => setNewTime(e.target.value)}
                className="px-2 py-1.5 text-xs bg-white border border-[#D1D1D1] rounded-[5px]"
              />
              <select
                value={newCategory}
                onChange={(e) => setNewCategory(e.target.value as CalendarScheduleEvent['category'])}
                className="px-2 py-1.5 text-xs bg-white border border-[#D1D1D1] rounded-[5px]"
              >
                <option value="meeting">회의</option>
                <option value="deadline">마감</option>
                <option value="manual">일반</option>
              </select>
              <select
                value={newProjectCode}
                onChange={(e) => setNewProjectCode(e.target.value)}
                className="px-2 py-1.5 text-xs bg-white border border-[#D1D1D1] rounded-[5px]"
              >
                <option value="">프로젝트 없음</option>
                {projects.map((p) => (
                  <option key={p.id} value={p.code}>
                    {p.code}
                  </option>
                ))}
              </select>
            </div>
            <input
              type="text"
              value={newNotes}
              onChange={(e) => setNewNotes(e.target.value)}
              placeholder="메모 (선택)"
              className="w-full px-2.5 py-1.5 text-xs bg-white border border-[#D1D1D1] rounded-[5px]"
            />
            <button
              type="submit"
              disabled={!newTitle.trim()}
              className="w-full inline-flex items-center justify-center gap-1.5 px-3 py-2 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] disabled:bg-slate-300 rounded-[5px]"
            >
              <Plus className="w-3.5 h-3.5" />
              일정 등록
            </button>
          </form>
        </aside>
      </div>
    </div>
  );
};
