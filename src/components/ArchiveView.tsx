import React, { useMemo, useState } from 'react';
import { FileText, Search, Trash2 } from 'lucide-react';
import type { SavedArchiveEntry } from '../types/mail';
import { AnalysisResultPanel } from './AnalysisResultPanel';

interface ArchiveViewProps {
  entries: SavedArchiveEntry[];
  onDeleteEntry: (id: string) => void;
  onGoToWorkbench: () => void;
}

export const ArchiveView: React.FC<ArchiveViewProps> = ({
  entries,
  onDeleteEntry,
  onGoToWorkbench,
}) => {
  const [query, setQuery] = useState('');
  const [selectedId, setSelectedId] = useState<string>(entries[0]?.id ?? '');

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return entries;
    return entries.filter(
      (entry) =>
        entry.result.workLog.title.toLowerCase().includes(q) ||
        entry.result.blogPost.title.toLowerCase().includes(q) ||
        entry.result.periodLabel.toLowerCase().includes(q)
    );
  }, [entries, query]);

  const activeEntry =
    filtered.find((e) => e.id === selectedId) ?? filtered[0] ?? null;

  if (entries.length === 0) {
    return (
      <div className="flex-1 flex flex-col items-center justify-center text-center px-6 py-16 bg-slate-50">
        <FileText className="w-10 h-10 text-slate-300 mb-3" />
        <h2 className="text-base font-bold text-slate-900">저장된 업무 일지 및 블로그 기록이 없습니다</h2>
        <p className="mt-1.5 text-sm text-slate-600 max-w-md">
          Outlook 워크벤치에서 메일 분석을 실행하면 생성된 업무 일지와 블로그 초안이 자동으로 이곳에 보관됩니다.
        </p>
        <button
          type="button"
          onClick={onGoToWorkbench}
          className="mt-5 px-4 py-2 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-lg transition-colors"
        >
          Outlook 워크벤치로 이동
        </button>
      </div>
    );
  }

  return (
    <div className="flex-1 flex flex-col lg:flex-row min-h-0 bg-white overflow-hidden">
      <aside className="w-full lg:w-[320px] shrink-0 border-b lg:border-b-0 lg:border-r border-slate-200 flex flex-col bg-slate-50/50">
        <div className="p-3.5 border-b border-slate-200 space-y-2">
          <div className="flex items-center justify-between">
            <h2 className="text-xs font-bold text-slate-900">생성 기록 보관함</h2>
            <span className="text-xs font-mono tabular-nums text-slate-500">{entries.length}건</span>
          </div>
          <div className="relative">
            <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
            <label htmlFor="archive-search" className="sr-only">
              기록 검색
            </label>
            <input
              id="archive-search"
              type="search"
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder="일지·블로그 제목, 기간 검색"
              className="w-full pl-8 pr-3 py-1.5 text-xs bg-white border border-slate-200 rounded-md focus:outline-none focus:border-[#0F6CBD]"
            />
          </div>
        </div>

        <div className="flex-1 overflow-y-auto divide-y divide-slate-200">
          {filtered.map((entry) => {
            const isActive = activeEntry?.id === entry.id;
            return (
              <div
                key={entry.id}
                className={`p-3.5 flex items-start justify-between gap-2 transition-colors ${
                  isActive ? 'bg-white' : 'hover:bg-white/70'
                }`}
              >
                <button
                  type="button"
                  onClick={() => setSelectedId(entry.id)}
                  className="flex-1 min-w-0 text-left space-y-1"
                >
                  <div className="text-xs font-semibold text-slate-900 truncate">
                    {entry.result.workLog.title}
                  </div>
                  <div className="text-[11px] text-slate-600 truncate">
                    {entry.result.blogPost.title}
                  </div>
                  <div className="text-[11px] font-mono tabular-nums text-slate-400">
                    {entry.result.periodLabel} · {entry.result.analyzedCount}건
                  </div>
                </button>
                <button
                  type="button"
                  onClick={() => onDeleteEntry(entry.id)}
                  aria-label={`'${entry.result.workLog.title}' 삭제`}
                  className="p-1 text-slate-400 hover:text-red-600 rounded shrink-0"
                >
                  <Trash2 className="w-3.5 h-3.5" />
                </button>
              </div>
            );
          })}
        </div>
      </aside>

      <div className="flex-1 min-w-0 flex flex-col overflow-hidden">
        <AnalysisResultPanel result={activeEntry?.result ?? null} isAnalyzing={false} />
      </div>
    </div>
  );
};
