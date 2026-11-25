import React, { useEffect, useRef, useState } from 'react';
import {
  AlertCircle,
  Paperclip,
  RotateCcw,
  SlidersHorizontal,
  X,
} from 'lucide-react';
import {
  countActiveGlobalSearchFilters,
  DEFAULT_GLOBAL_SEARCH_FILTERS,
  resolveSearchDateRange,
  validateGlobalSearchFilters,
  type GlobalSearchFilters,
  type SearchDateRangePreset,
} from '../utils/mailFilter';

interface GlobalSearchFilterDropdownProps {
  filters: GlobalSearchFilters;
  onChange: (next: GlobalSearchFilters) => void;
  onApply?: () => void;
  referenceDate: string;
  integrated?: boolean;
}

const DATE_PRESET_OPTIONS: Array<{ id: SearchDateRangePreset; label: string }> = [
  { id: 'any', label: '전체 기간' },
  { id: 'today', label: '오늘' },
  { id: '7d', label: '최근 7일' },
  { id: '30d', label: '최근 30일' },
  { id: 'custom', label: '직접 지정' },
];

const FLUENT_INPUT_CLASS =
  'w-full px-2.5 py-1.5 text-xs text-[#242424] bg-white border border-[#D1D1D1] border-b-[#616161] rounded-[4px] hover:border-[#B3B3B3] focus:outline-none focus:border-[#D1D1D1] focus:border-b-2 focus:border-b-[#0F6CBD] placeholder:text-[#707070] transition-colors duration-150';

export const GlobalSearchFilterDropdown: React.FC<GlobalSearchFilterDropdownProps> = ({
  filters,
  onChange,
  onApply,
  referenceDate,
  integrated = false,
}) => {
  const [isOpen, setIsOpen] = useState(false);
  const [draft, setDraft] = useState<GlobalSearchFilters>(filters);
  const containerRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    setDraft(filters);
  }, [filters]);

  useEffect(() => {
    if (!isOpen) return;
    const handleOutside = (e: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) {
        setIsOpen(false);
      }
    };
    const handleEsc = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setIsOpen(false);
    };
    window.addEventListener('mousedown', handleOutside);
    window.addEventListener('keydown', handleEsc);
    return () => {
      window.removeEventListener('mousedown', handleOutside);
      window.removeEventListener('keydown', handleEsc);
    };
  }, [isOpen]);

  const activeCount = countActiveGlobalSearchFilters(filters);
  const validation = validateGlobalSearchFilters(draft);

  const handleSelectDatePreset = (preset: SearchDateRangePreset) => {
    if (preset === 'custom') {
      setDraft((prev) => ({ ...prev, datePreset: 'custom' }));
      return;
    }
    const range = resolveSearchDateRange(preset, referenceDate);
    setDraft((prev) => ({
      ...prev,
      datePreset: preset,
      dateFrom: range.dateFrom,
      dateTo: range.dateTo,
    }));
  };

  const handleApply = () => {
    if (!validation.valid) return;
    onChange(draft);
    onApply?.();
    setIsOpen(false);
  };

  const handleReset = () => {
    setDraft(DEFAULT_GLOBAL_SEARCH_FILTERS);
    onChange(DEFAULT_GLOBAL_SEARCH_FILTERS);
    onApply?.();
  };

  return (
    <div ref={containerRef} className="relative shrink-0">
      <button
        type="button"
        onClick={() => setIsOpen((v) => !v)}
        aria-label="고급 검색 필터"
        aria-expanded={isOpen}
        aria-haspopup="dialog"
        title="검색 필터 (보낸 사람 · 받는 사람 · 날짜 범위 · 첨부파일)"
        className={
          integrated
            ? `relative h-6 px-1.5 inline-flex items-center justify-center gap-1 rounded-[4px] text-xs transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]/60 ${
                activeCount > 0 || isOpen
                  ? 'bg-[#EBF3FC] text-[#0F6CBD] font-semibold'
                  : 'text-[#616161] hover:bg-black/[0.06] hover:text-[#242424] active:bg-black/[0.1]'
              }`
            : `relative h-8 px-2 inline-flex items-center justify-center gap-1 rounded-[4px] text-xs transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-white/80 ${
                activeCount > 0
                  ? 'bg-white text-[#0F6CBD] font-semibold shadow-xs'
                  : 'bg-white/95 hover:bg-white text-[#424242] hover:text-[#0F6CBD] shadow-xs'
              }`
        }
      >
        <SlidersHorizontal className="w-3.5 h-3.5" strokeWidth={1.75} />
        <span className="hidden md:inline text-[11px] font-medium whitespace-nowrap">필터</span>
        {activeCount > 0 && (
          <span className="min-w-[16px] h-[16px] px-1 rounded-full bg-[#0F6CBD] text-white text-[10px] font-mono tabular-nums font-bold inline-flex items-center justify-center leading-none">
            {activeCount}
          </span>
        )}
      </button>

      {isOpen && (
        <div
          role="dialog"
          aria-label="Bmail 고급 검색 필터"
          className="absolute right-0 top-[calc(100%+8px)] w-[356px] sm:w-[420px] max-w-[calc(100vw-24px)] bg-white text-[#242424] rounded-[8px] border border-[#E0E0E0] shadow-[0_8px_16px_rgba(0,0,0,0.14),0_0_2px_rgba(0,0,0,0.12)] z-50 overflow-hidden"
        >
          {/* Fluent 2 Popover Header */}
          <div className="px-4 py-2.5 bg-[#FAF9F8] border-b border-[#EDEBE9] flex items-center justify-between">
            <div className="flex items-center gap-2">
              <SlidersHorizontal className="w-4 h-4 text-[#0F6CBD]" strokeWidth={1.75} />
              <span className="text-[13px] font-semibold text-[#242424]">검색 필터</span>
              {activeCount > 0 && (
                <span className="text-[11px] font-medium text-[#0F6CBD] font-mono tabular-nums">
                  {activeCount}개 적용됨
                </span>
              )}
            </div>
            <button
              type="button"
              onClick={() => setIsOpen(false)}
              aria-label="검색 필터 닫기"
              className="w-7 h-7 flex items-center justify-center rounded-[4px] text-[#616161] hover:bg-black/[0.05] hover:text-[#242424] active:bg-black/[0.08] transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]/60"
            >
              <X className="w-4 h-4" strokeWidth={1.75} />
            </button>
          </div>

          {/* Fluent 2 Filter Form Body */}
          <form
            onSubmit={(e) => {
              e.preventDefault();
              handleApply();
            }}
            className="p-4 space-y-3 text-xs"
          >
            <div className="grid grid-cols-[76px_1fr] items-center gap-2.5">
              <label htmlFor="sf-from" className="text-[12px] text-[#424242] font-medium">
                보낸 사람
              </label>
              <input
                id="sf-from"
                type="text"
                value={draft.from}
                onChange={(e) => setDraft((prev) => ({ ...prev, from: e.target.value }))}
                placeholder="이름 또는 이메일 (예: 김철수, @vendor.com)"
                className={FLUENT_INPUT_CLASS}
              />
            </div>

            <div className="grid grid-cols-[76px_1fr] items-center gap-2.5">
              <label htmlFor="sf-to" className="text-[12px] text-[#424242] font-medium">
                받는 사람
              </label>
              <input
                id="sf-to"
                type="text"
                value={draft.to}
                onChange={(e) => setDraft((prev) => ({ ...prev, to: e.target.value }))}
                placeholder="수신자 이메일 또는 이름 (To · Cc)"
                className={FLUENT_INPUT_CLASS}
              />
            </div>

            {/* Date Presets + Range */}
            <div className="pt-2.5 border-t border-[#EDEBE9] space-y-2">
              <div className="flex items-center justify-between">
                <span className="text-[12px] text-[#424242] font-medium">날짜 범위</span>
                <span className="text-[11px] font-mono tabular-nums text-[#616161]">
                  {draft.dateFrom || draft.dateTo
                    ? `${draft.dateFrom || '...'} ~ ${draft.dateTo || '...'}`
                    : '제한 없음'}
                </span>
              </div>
              <div className="flex flex-wrap gap-1">
                {DATE_PRESET_OPTIONS.map((preset) => {
                  const selected = draft.datePreset === preset.id;
                  return (
                    <button
                      key={preset.id}
                      type="button"
                      onClick={() => handleSelectDatePreset(preset.id)}
                      className={`px-2 py-1 text-[11px] rounded-[4px] border transition-colors duration-150 whitespace-nowrap focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]/60 ${
                        selected
                          ? 'bg-[#EBF3FC] border-[#0F6CBD] text-[#0F6CBD] font-semibold'
                          : 'bg-white border-[#D1D1D1] text-[#424242] hover:bg-[#F5F5F5] active:bg-[#EBEBEB]'
                      }`}
                    >
                      {preset.label}
                    </button>
                  );
                })}
              </div>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label htmlFor="sf-date-from" className="block text-[11px] text-[#616161] mb-1">
                    시작일
                  </label>
                  <input
                    id="sf-date-from"
                    type="date"
                    value={draft.dateFrom}
                    onChange={(e) =>
                      setDraft((prev) => ({
                        ...prev,
                        datePreset: 'custom',
                        dateFrom: e.target.value,
                      }))
                    }
                    className={`${FLUENT_INPUT_CLASS} font-mono tabular-nums`}
                  />
                </div>
                <div>
                  <label htmlFor="sf-date-to" className="block text-[11px] text-[#616161] mb-1">
                    종료일
                  </label>
                  <input
                    id="sf-date-to"
                    type="date"
                    value={draft.dateTo}
                    onChange={(e) =>
                      setDraft((prev) => ({
                        ...prev,
                        datePreset: 'custom',
                        dateTo: e.target.value,
                      }))
                    }
                    className={`${FLUENT_INPUT_CLASS} font-mono tabular-nums`}
                  />
                </div>
              </div>
            </div>

            {/* Attachment Checkbox */}
            <div className="pt-2.5 border-t border-[#EDEBE9]">
              <label className="inline-flex items-center gap-2 text-xs text-[#242424] cursor-pointer select-none px-1.5 py-1 rounded-[4px] hover:bg-[#FAF9F8]">
                <input
                  type="checkbox"
                  checked={draft.hasAttachments}
                  onChange={(e) =>
                    setDraft((prev) => ({ ...prev, hasAttachments: e.target.checked }))
                  }
                  className="accent-[#0F6CBD] w-3.5 h-3.5 rounded-[3px]"
                />
                <Paperclip className="w-3.5 h-3.5 text-[#0F6CBD]" strokeWidth={1.75} />
                첨부 파일이 있는 메일만 표시
              </label>
            </div>

            {!validation.valid && validation.error && (
              <div
                role="alert"
                className="rounded-[4px] border border-[#F1BBBC] bg-[#FDE7E9] px-2.5 py-1.5 text-[11px] text-[#A4262C] flex items-center gap-1.5"
              >
                <AlertCircle className="w-3.5 h-3.5 shrink-0 text-[#C50F1F]" strokeWidth={1.75} />
                <span>{validation.error}</span>
              </div>
            )}

            {/* Fluent 2 Action Footer */}
            <div className="pt-3 border-t border-[#EDEBE9] flex items-center justify-between gap-2">
              <button
                type="button"
                onClick={handleReset}
                className="inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs text-[#424242] hover:bg-black/[0.05] active:bg-black/[0.08] rounded-[4px] transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]/60"
              >
                <RotateCcw className="w-3.5 h-3.5" strokeWidth={1.75} />
                필터 초기화
              </button>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => setIsOpen(false)}
                  className="px-3 py-1.5 text-xs font-medium text-[#242424] bg-white border border-[#D1D1D1] hover:bg-[#F5F5F5] active:bg-[#EBEBEB] rounded-[4px] transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]/60"
                >
                  취소
                </button>
                <button
                  type="submit"
                  disabled={!validation.valid}
                  className="px-4 py-1.5 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] active:bg-[#0E4775] disabled:bg-[#F0F0F0] disabled:text-[#A19F9D] rounded-[4px] transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-offset-1 focus-visible:ring-[#0F6CBD]"
                >
                  검색 적용
                </button>
              </div>
            </div>
          </form>
        </div>
      )}
    </div>
  );
};
