import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
  FileText,
  Pin,
  Plus,
  Search,
  StickyNote,
  Trash2,
  X,
} from 'lucide-react';
import type { MailItem } from '../types/mail';

export type StickyNoteColor = 'yellow' | 'blue' | 'green' | 'pink';

export interface StickyNoteRecord {
  id: string;
  title: string;
  content: string;
  color: StickyNoteColor;
  pinned: boolean;
  linkedEmailId?: string;
  linkedSubject?: string;
  updatedAt: string;
}

interface StickyNotesWidgetProps {
  open: boolean;
  onClose: () => void;
  activeEmail?: MailItem | null;
  onOpenEmailById?: (emailId: string) => void;
}

const COLOR_STYLES: Record<
  StickyNoteColor,
  { bg: string; border: string; header: string; swatch: string; label: string }
> = {
  yellow: {
    bg: 'bg-[#FFF9C4]',
    border: 'border-[#FBC02D]/60',
    header: 'bg-[#FFF59D]',
    swatch: 'bg-[#FDE047]',
    label: '노랑',
  },
  blue: {
    bg: 'bg-[#E3F2FD]',
    border: 'border-[#64B5F6]/70',
    header: 'bg-[#BBDEFB]',
    swatch: 'bg-[#60A5FA]',
    label: '파랑',
  },
  green: {
    bg: 'bg-[#E8F5E9]',
    border: 'border-[#81C784]/70',
    header: 'bg-[#C8E6C9]',
    swatch: 'bg-[#4ADE80]',
    label: '초록',
  },
  pink: {
    bg: 'bg-[#FCE4EC]',
    border: 'border-[#F06292]/60',
    header: 'bg-[#F8BBD0]',
    swatch: 'bg-[#F472B6]',
    label: '분홍',
  },
};

const STORAGE_KEY = 'bmail.stickyNotes.v1';

export const StickyNotesWidget: React.FC<StickyNotesWidgetProps> = ({
  open,
  onClose,
  activeEmail,
  onOpenEmailById,
}) => {
  const [notes, setNotes] = useState<StickyNoteRecord[]>(() => {
    try {
      const raw = window.localStorage.getItem(STORAGE_KEY);
      const parsed = raw ? (JSON.parse(raw) as StickyNoteRecord[]) : [];
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  });
  const [searchQuery, setSearchQuery] = useState('');
  const [loadedFromServer, setLoadedFromServer] = useState(false);
  const persistTimerRef = useRef<number | null>(null);

  useEffect(() => {
    let active = true;
    fetch('/api/local/bootstrap?limit=1')
      .then((res) => (res.ok ? res.json() : null))
      .then((data: { kvState?: Record<string, unknown> } | null) => {
        if (!active || !data?.kvState) return;
        const serverNotes = data.kvState.sticky_notes;
        if (Array.isArray(serverNotes) && serverNotes.length > 0) {
          setNotes(serverNotes as StickyNoteRecord[]);
          try {
            window.localStorage.setItem(STORAGE_KEY, JSON.stringify(serverNotes));
          } catch {
            // ignore
          }
        }
      })
      .catch(() => {
        // offline fallback to localStorage
      })
      .finally(() => {
        if (active) setLoadedFromServer(true);
      });
    return () => {
      active = false;
    };
  }, []);

  useEffect(() => {
    if (!loadedFromServer) return;
    try {
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify(notes));
    } catch {
      // ignore
    }
    if (persistTimerRef.current !== null) {
      window.clearTimeout(persistTimerRef.current);
    }
    persistTimerRef.current = window.setTimeout(() => {
      fetch('/api/local/state', {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ key: 'sticky_notes', value: notes }),
      }).catch(() => {
        // non-fatal
      });
    }, 350);

    return () => {
      if (persistTimerRef.current !== null) {
        window.clearTimeout(persistTimerRef.current);
      }
    };
  }, [notes, loadedFromServer]);

  const filteredNotes = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    const sorted = [...notes].sort((a, b) => {
      if (a.pinned !== b.pinned) return a.pinned ? -1 : 1;
      return b.updatedAt.localeCompare(a.updatedAt);
    });
    if (!q) return sorted;
    return sorted.filter(
      (n) =>
        n.title.toLowerCase().includes(q) ||
        n.content.toLowerCase().includes(q) ||
        (n.linkedSubject || '').toLowerCase().includes(q)
    );
  }, [notes, searchQuery]);

  if (!open) return null;

  const handleAddNote = (fromEmail?: MailItem | null) => {
    const next: StickyNoteRecord = {
      id: `note-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
      title: fromEmail ? `[메일 메모] ${fromEmail.subject.slice(0, 32)}` : '새 업무 메모',
      content: fromEmail
        ? `발신: ${fromEmail.fromName} <${fromEmail.fromEmail}>\n일자: ${fromEmail.date}\n핵심 메모: `
        : '',
      color: fromEmail ? 'blue' : 'yellow',
      pinned: false,
      linkedEmailId: fromEmail?.id,
      linkedSubject: fromEmail?.subject,
      updatedAt: new Date().toISOString(),
    };
    setNotes((prev) => [next, ...prev]);
  };

  const handlePatchNote = (id: string, patch: Partial<StickyNoteRecord>) => {
    setNotes((prev) =>
      prev.map((n) =>
        n.id === id ? { ...n, ...patch, updatedAt: new Date().toISOString() } : n
      )
    );
  };

  const handleDeleteNote = (id: string) => {
    setNotes((prev) => prev.filter((n) => n.id !== id));
  };

  return (
    <aside
      aria-label="스티커 메모 플로팅 위젯"
      className="fixed bottom-16 right-4 z-40 w-[350px] max-w-[calc(100vw-2rem)] bg-white border border-[#D1D1D1] rounded-lg shadow-2xl flex flex-col overflow-hidden"
      style={{ maxHeight: '72vh' }}
    >
      {/* Header */}
      <div className="px-3.5 py-2.5 bg-[#0F6CBD] text-white flex items-center justify-between gap-2">
        <div className="flex items-center gap-2 min-w-0">
          <StickyNote className="w-4 h-4 shrink-0" />
          <span className="text-xs font-bold truncate">
            스티커 메모 (Sticky Notes · {notes.length}건)
          </span>
        </div>
        <div className="flex items-center gap-1 shrink-0">
          <button
            type="button"
            onClick={() => handleAddNote(null)}
            title="빈 메모 추가"
            className="inline-flex items-center gap-1 px-2 py-1 text-[11px] font-semibold bg-white/15 hover:bg-white/25 rounded"
          >
            <Plus className="w-3.5 h-3.5" />
            새 메모
          </button>
          <button
            type="button"
            onClick={onClose}
            aria-label="스티커 메모 닫기"
            className="p-1 hover:bg-white/20 rounded"
          >
            <X className="w-4 h-4" />
          </button>
        </div>
      </div>

      {/* Quick Action Bar */}
      <div className="px-3 py-2 bg-[#FAF9F8] border-b border-[#EDEBE9] flex flex-col gap-1.5">
        <div className="flex items-center gap-1.5">
          <div className="relative flex-1">
            <Search className="w-3.5 h-3.5 text-[#707070] absolute left-2 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="메모 제목 · 본문 검색..."
              className="w-full pl-7 pr-2 py-1 text-[11px] bg-white border border-[#D1D1D1] rounded focus:outline-none focus:border-[#0F6CBD]"
            />
          </div>
          {activeEmail && (
            <button
              type="button"
              onClick={() => handleAddNote(activeEmail)}
              title={`"${activeEmail.subject}" 기준 메모 생성`}
              className="inline-flex items-center gap-1 px-2 py-1 text-[11px] font-semibold text-[#0F6CBD] bg-[#EBF3FC] hover:bg-[#DCEBFA] border border-[#B4D6FA] rounded whitespace-nowrap"
            >
              <FileText className="w-3 h-3" />
              메일 메모
            </button>
          )}
        </div>
      </div>

      {/* Notes Stack */}
      <div className="flex-1 overflow-y-auto p-3 space-y-2.5 bg-[#F5F5F5]">
        {filteredNotes.length === 0 ? (
          <div className="py-8 px-4 text-center text-xs text-[#616161] space-y-2">
            <StickyNote className="w-7 h-7 text-[#9E9E9E] mx-auto" />
            <p>저장된 스티커 메모가 없습니다.</p>
            <button
              type="button"
              onClick={() => handleAddNote(activeEmail)}
              className="inline-flex items-center gap-1 px-3 py-1.5 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded"
            >
              <Plus className="w-3.5 h-3.5" />
              첫 메모 작성하기
            </button>
          </div>
        ) : (
          filteredNotes.map((note) => {
            const palette = COLOR_STYLES[note.color] || COLOR_STYLES.yellow;
            return (
              <div
                key={note.id}
                className={`rounded-md border shadow-sm overflow-hidden ${palette.bg} ${palette.border}`}
              >
                <div className={`px-2.5 py-1.5 flex items-center justify-between gap-1.5 ${palette.header}`}>
                  <input
                    type="text"
                    value={note.title}
                    onChange={(e) => handlePatchNote(note.id, { title: e.target.value })}
                    className="flex-1 min-w-0 bg-transparent text-xs font-bold text-[#242424] focus:outline-none"
                    placeholder="메모 제목"
                  />
                  <div className="flex items-center gap-1 shrink-0">
                    {(Object.keys(COLOR_STYLES) as StickyNoteColor[]).map((c) => (
                      <button
                        key={c}
                        type="button"
                        onClick={() => handlePatchNote(note.id, { color: c })}
                        aria-label={`색상 ${COLOR_STYLES[c].label}`}
                        className={`w-3 h-3 rounded-full border ${COLOR_STYLES[c].swatch} ${
                          note.color === c ? 'border-[#242424] scale-110' : 'border-black/20 opacity-70'
                        }`}
                      />
                    ))}
                    <button
                      type="button"
                      onClick={() => handlePatchNote(note.id, { pinned: !note.pinned })}
                      title={note.pinned ? '상단 고정 해제' : '상단 고정'}
                      className={`p-0.5 rounded ${
                        note.pinned ? 'text-[#0F6CBD] bg-white/60' : 'text-[#616161] hover:bg-white/40'
                      }`}
                    >
                      <Pin className="w-3 h-3" />
                    </button>
                    <button
                      type="button"
                      onClick={() => handleDeleteNote(note.id)}
                      title="메모 삭제"
                      className="p-0.5 rounded text-[#616161] hover:text-red-700 hover:bg-white/40"
                    >
                      <Trash2 className="w-3 h-3" />
                    </button>
                  </div>
                </div>
                <div className="p-2.5 space-y-1.5">
                  <textarea
                    rows={3}
                    value={note.content}
                    onChange={(e) => handlePatchNote(note.id, { content: e.target.value })}
                    placeholder="업무 메모, 체크사항, 도면 번호 등을 자유롭게 기록하세요..."
                    className="w-full bg-transparent text-[11px] leading-relaxed text-[#242424] resize-y focus:outline-none"
                  />
                  <div className="flex items-center justify-between gap-2 text-[10px] text-[#616161]">
                    <span className="font-mono">{note.updatedAt.slice(0, 16).replace('T', ' ')}</span>
                    {note.linkedEmailId && onOpenEmailById && (
                      <button
                        type="button"
                        onClick={() => onOpenEmailById(note.linkedEmailId!)}
                        className="text-[#0F6CBD] hover:underline truncate max-w-[160px]"
                      >
                        원본 메일 이동 →
                      </button>
                    )}
                  </div>
                </div>
              </div>
            );
          })
        )}
      </div>
    </aside>
  );
};
