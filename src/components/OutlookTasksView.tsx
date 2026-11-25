import React, { useMemo, useState } from 'react';
import {
  Calendar,
  CheckCircle2,
  Circle,
  Flag,
  Mail,
  Plus,
  Trash2,
} from 'lucide-react';
import type { MailItem } from '../types/mail';
import type { OutlookTaskItem } from '../utils/outlookWorkspaceModel';

interface OutlookTasksViewProps {
  tasks: OutlookTaskItem[];
  emails: MailItem[];
  onChangeTasks: (next: OutlookTaskItem[]) => void;
  onOpenMail: (mail: MailItem) => void;
  onToggleMailFlag: (mailId: string) => void;
}

export const OutlookTasksView: React.FC<OutlookTasksViewProps> = ({
  tasks,
  emails,
  onChangeTasks,
  onOpenMail,
  onToggleMailFlag,
}) => {
  const [title, setTitle] = useState('');
  const [dueDate, setDueDate] = useState('');
  const [priority, setPriority] = useState<'low' | 'normal' | 'high'>('normal');
  const [filter, setFilter] = useState<'all' | 'active' | 'completed'>('active');

  const flaggedEmails = useMemo(
    () => emails.filter((m) => m.isFlagged).slice(0, 100),
    [emails]
  );

  const visibleTasks = useMemo(() => {
    return tasks.filter((t) => {
      if (filter === 'active') return !t.completed;
      if (filter === 'completed') return t.completed;
      return true;
    });
  }, [tasks, filter]);

  const handleAddTask = (e: React.FormEvent) => {
    e.preventDefault();
    const cleanTitle = title.trim();
    if (!cleanTitle) return;
    const item: OutlookTaskItem = {
      id: `task-${Date.now()}`,
      title: cleanTitle,
      dueDate: dueDate || undefined,
      priority,
      completed: false,
      createdAt: new Date().toISOString(),
    };
    onChangeTasks([item, ...tasks]);
    setTitle('');
    setDueDate('');
    setPriority('normal');
  };

  return (
    <div className="flex-1 flex flex-col min-h-0 bg-[#FAF9F8] text-[#242424] overflow-y-auto p-4 sm:p-6">
      <div className="max-w-5xl w-full mx-auto space-y-6">
        {/* Header & Quick Add */}
        <div className="bg-white rounded-[8px] border border-[#E0E0E0] p-4 space-y-3 shadow-2xs">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div className="flex items-center gap-2">
              <CheckCircle2 className="w-4 h-4 text-[#0F6CBD]" />
              <h2 className="text-sm font-bold text-[#242424]">할 일 · To-Do & 플래그된 메일</h2>
            </div>
            <div className="inline-flex rounded-[4px] border border-[#D1D1D1] overflow-hidden">
              {(
                [
                  { id: 'active', label: `진행 중 (${tasks.filter((t) => !t.completed).length})` },
                  { id: 'completed', label: `완료 (${tasks.filter((t) => t.completed).length})` },
                  { id: 'all', label: `전체 (${tasks.length})` },
                ] as const
              ).map((tab) => (
                <button
                  key={tab.id}
                  type="button"
                  onClick={() => setFilter(tab.id)}
                  className={`px-3 py-1 text-xs ${
                    filter === tab.id
                      ? 'bg-[#0F6CBD] text-white font-semibold'
                      : 'bg-white text-[#424242] hover:bg-[#F5F5F5]'
                  }`}
                >
                  {tab.label}
                </button>
              ))}
            </div>
          </div>

          <form onSubmit={handleAddTask} className="flex flex-wrap gap-2">
            <input
              type="text"
              value={title}
              onChange={(e) => setTitle(e.target.value)}
              placeholder="새 작업 추가 (예: DWG-2201 개정 검토 회신)..."
              aria-label="새 작업 제목"
              className="flex-1 min-w-[200px] px-3 py-2 text-xs bg-white border border-[#D1D1D1] rounded-[4px] focus:outline-none focus:border-[#0F6CBD]"
            />
            <input
              type="date"
              value={dueDate}
              onChange={(e) => setDueDate(e.target.value)}
              aria-label="작업 기한"
              className="px-2.5 py-2 text-xs bg-white border border-[#D1D1D1] rounded-[4px] font-mono tabular-nums"
            />
            <select
              value={priority}
              onChange={(e) => setPriority(e.target.value as 'low' | 'normal' | 'high')}
              aria-label="작업 중요도"
              className="px-2.5 py-2 text-xs bg-white border border-[#D1D1D1] rounded-[4px]"
            >
              <option value="high">중요도: 높음</option>
              <option value="normal">중요도: 보통</option>
              <option value="low">중요도: 낮음</option>
            </select>
            <button
              type="submit"
              className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-[4px]"
            >
              <Plus className="w-3.5 h-3.5" />
              추가
            </button>
          </form>
        </div>

        {/* Two Column Tasks & Flagged Emails */}
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          {/* Custom Tasks */}
          <div className="bg-white rounded-[8px] border border-[#E0E0E0] p-4 space-y-3 shadow-2xs">
            <div className="text-xs font-bold text-[#242424] flex items-center justify-between">
              <span>내 작업 ({visibleTasks.length})</span>
            </div>
            {visibleTasks.length === 0 ? (
              <div className="p-6 text-center text-xs text-[#616161] border border-dashed border-[#D1D1D1] rounded-[6px]">
                표시할 작업이 없습니다. 위에서 할 일을 등록하세요.
              </div>
            ) : (
              <div className="space-y-1.5">
                {visibleTasks.map((task) => (
                  <div
                    key={task.id}
                    className="flex items-center justify-between gap-2.5 px-3 py-2 rounded-[6px] border border-[#EDEBE9] hover:bg-[#FAF9F8]"
                  >
                    <button
                      type="button"
                      onClick={() =>
                        onChangeTasks(
                          tasks.map((t) =>
                            t.id === task.id ? { ...t, completed: !t.completed } : t
                          )
                        )
                      }
                      className="flex items-center gap-2.5 min-w-0 flex-1 text-left"
                    >
                      {task.completed ? (
                        <CheckCircle2 className="w-4 h-4 text-[#107C10] shrink-0" />
                      ) : (
                        <Circle className="w-4 h-4 text-[#616161] shrink-0" />
                      )}
                      <div className="min-w-0">
                        <div
                          className={`text-xs font-medium truncate ${
                            task.completed ? 'line-through text-[#616161]' : 'text-[#242424]'
                          }`}
                        >
                          {task.title}
                        </div>
                        <div className="flex items-center gap-2 text-[10px] text-[#616161] mt-0.5">
                          {task.priority === 'high' && (
                            <span className="px-1.5 py-0.2 rounded bg-red-50 text-red-700 font-semibold">
                              긴급
                            </span>
                          )}
                          {task.dueDate && (
                            <span className="inline-flex items-center gap-1 font-mono">
                              <Calendar className="w-3 h-3" />
                              {task.dueDate}
                            </span>
                          )}
                        </div>
                      </div>
                    </button>
                    <button
                      type="button"
                      onClick={() => onChangeTasks(tasks.filter((t) => t.id !== task.id))}
                      aria-label="작업 삭제"
                      className="p-1 text-[#616161] hover:text-red-700 rounded"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                    </button>
                  </div>
                ))}
              </div>
            )}
          </div>

          {/* Flagged Emails */}
          <div className="bg-white rounded-[8px] border border-[#E0E0E0] p-4 space-y-3 shadow-2xs">
            <div className="text-xs font-bold text-[#242424] flex items-center gap-1.5">
              <Flag className="w-3.5 h-3.5 text-[#D13438]" />
              <span>플래그가 지정된 메일 ({flaggedEmails.length})</span>
            </div>
            {flaggedEmails.length === 0 ? (
              <div className="p-6 text-center text-xs text-[#616161] border border-dashed border-[#D1D1D1] rounded-[6px]">
                플래그가 지정된 메일이 없습니다. 메일 목록에서 플래그를 누르면 여기에 자동 연동됩니다.
              </div>
            ) : (
              <div className="space-y-1.5 max-h-[520px] overflow-y-auto">
                {flaggedEmails.map((m) => (
                  <div
                    key={m.id}
                    className="flex items-center justify-between gap-2.5 px-3 py-2 rounded-[6px] border border-[#EDEBE9] hover:bg-[#FAF9F8]"
                  >
                    <button
                      type="button"
                      onClick={() => onOpenMail(m)}
                      className="min-w-0 flex-1 text-left"
                    >
                      <div className="text-xs font-semibold text-[#242424] truncate">{m.subject}</div>
                      <div className="text-[11px] text-[#616161] truncate">
                        {m.fromName || m.fromEmail} · {m.date}
                      </div>
                    </button>
                    <div className="flex items-center gap-1 shrink-0">
                      <button
                        type="button"
                        onClick={() => onOpenMail(m)}
                        title="메일 본문 열기"
                        className="p-1.5 text-[#0F6CBD] hover:bg-[#EBF3FC] rounded"
                      >
                        <Mail className="w-3.5 h-3.5" />
                      </button>
                      <button
                        type="button"
                        onClick={() => onToggleMailFlag(m.id)}
                        title="플래그 해제"
                        className="p-1.5 text-[#D13438] hover:bg-red-50 rounded"
                      >
                        <Flag className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};
export default OutlookTasksView;
