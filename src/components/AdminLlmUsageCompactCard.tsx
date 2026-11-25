import React, { useCallback, useEffect, useState } from 'react';
import {
  Activity,
  ChevronDown,
  ChevronUp,
  Cpu,
  RefreshCw,
  RotateCcw,
} from 'lucide-react';
import type { LlmUsageSummary } from '../services/llmUsageTracker';

export const AdminLlmUsageCompactCard: React.FC = () => {
  const [summary, setSummary] = useState<LlmUsageSummary | null>(null);
  const [loading, setLoading] = useState(false);
  const [expanded, setExpanded] = useState(false);

  const fetchUsage = useCallback(async () => {
    setLoading(true);
    try {
      const res = await fetch('/api/admin/llm-usage?limit=8');
      if (res.ok) {
        const data = (await res.json()) as LlmUsageSummary;
        setSummary(data);
      }
    } catch {
      // keep silent in offline or initial load
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    void fetchUsage();
  }, [fetchUsage]);

  const handleReset = async () => {
    setLoading(true);
    try {
      const res = await fetch('/api/admin/llm-usage/reset', { method: 'POST' });
      if (res.ok) {
        const data = (await res.json()) as { summary?: LlmUsageSummary };
        if (data.summary) setSummary(data.summary);
      }
    } finally {
      setLoading(false);
    }
  };

  const totals = summary?.totals || {
    totalCalls: 0,
    successCalls: 0,
    errorCalls: 0,
    inputTokens: 0,
    outputTokens: 0,
    totalTokens: 0,
    totalCostUsd: 0,
    avgLatencyMs: 0,
  };
  const today = summary?.today || {
    date: new Date().toISOString().slice(0, 10),
    calls: 0,
    totalTokens: 0,
    costUsd: 0,
  };
  const byModel = summary?.byModel || [];
  const recentEvents = summary?.recentEvents || [];

  const formatNum = (n: number) => n.toLocaleString('ko-KR');
  const formatUsd = (n: number) =>
    n === 0 ? '$0.0000' : n < 0.0001 ? '<$0.0001' : `$${n.toFixed(4)}`;

  const tierBadge = (tier: string) => {
    switch (tier) {
      case 'high':
        return 'bg-purple-50 text-purple-700 border-purple-200';
      case 'medium':
        return 'bg-[#EBF3FC] text-[#0F6CBD] border-[#B4D6FA]';
      case 'low':
        return 'bg-emerald-50 text-emerald-700 border-emerald-200';
      default:
        return 'bg-slate-100 text-slate-600 border-slate-200';
    }
  };

  return (
    <section
      aria-label="관리자 LLM 사용량 미니 모니터"
      className="bg-white border border-[#E0E0E0] rounded-[8px] px-4 py-3 shadow-[0_1px_2px_rgba(0,0,0,0.03)]"
    >
      {/* Compact Single-Row Summary Bar */}
      <div className="flex flex-wrap items-center justify-between gap-2.5">
        <div className="flex items-center gap-2 min-w-0">
          <span className="inline-flex items-center justify-center w-6 h-6 rounded-[4px] bg-[#EBF3FC] text-[#0F6CBD] shrink-0">
            <Cpu className="w-3.5 h-3.5" strokeWidth={1.8} />
          </span>
          <div className="min-w-0">
            <div className="flex items-center gap-1.5 flex-wrap">
              <h2 className="text-xs font-bold text-[#242424]">
                LLM 토큰 · 모델 사용량 요약
              </h2>
              <span className="text-[10px] font-medium px-1.5 py-0.5 rounded bg-[#FAF9F8] border border-[#E0E0E0] text-[#616161]">
                Claude 5.5 3-티어 · Gemini · Groq
              </span>
            </div>
          </div>
        </div>

        {/* Inline Mini Metrics Pills */}
        <div className="flex flex-wrap items-center gap-1.5 text-[11px]">
          <div className="px-2.5 py-1 rounded-[4px] bg-[#FAF9F8] border border-[#EDEBE9] flex items-center gap-1.5">
            <span className="text-[#616161]">누적 호출</span>
            <strong className="font-mono text-[#242424]">{formatNum(totals.totalCalls)}회</strong>
            {totals.errorCalls > 0 && (
              <span className="text-[10px] text-red-600 font-mono">
                (오류 {totals.errorCalls})
              </span>
            )}
          </div>

          <div className="px-2.5 py-1 rounded-[4px] bg-[#FAF9F8] border border-[#EDEBE9] flex items-center gap-1.5">
            <span className="text-[#616161]">총 토큰</span>
            <strong className="font-mono text-[#0F6CBD]">{formatNum(totals.totalTokens)}</strong>
            <span className="text-[10px] text-[#707070] font-mono">
              (In {formatNum(totals.inputTokens)} / Out {formatNum(totals.outputTokens)})
            </span>
          </div>

          <div className="px-2.5 py-1 rounded-[4px] bg-[#FAF9F8] border border-[#EDEBE9] flex items-center gap-1.5">
            <span className="text-[#616161]">오늘({today.date.slice(5)})</span>
            <strong className="font-mono text-[#242424]">
              {formatNum(today.calls)}회 · {formatNum(today.totalTokens)} tok
            </strong>
          </div>

          <div className="px-2.5 py-1 rounded-[4px] bg-[#FAF9F8] border border-[#EDEBE9] flex items-center gap-1.5">
            <span className="text-[#616161]">추정 비용</span>
            <strong className="font-mono text-emerald-700">{formatUsd(totals.totalCostUsd)}</strong>
          </div>

          <div className="px-2 py-1 rounded-[4px] bg-[#FAF9F8] border border-[#EDEBE9] flex items-center gap-1 text-[10px] text-[#616161] font-mono">
            <Activity className="w-3 h-3 text-[#0F6CBD]" />
            평균 {formatNum(totals.avgLatencyMs)}ms
          </div>

          <button
            type="button"
            onClick={() => void fetchUsage()}
            disabled={loading}
            title="사용량 새로고침"
            className="inline-flex items-center gap-1 px-2 py-1 text-[11px] font-medium text-[#242424] bg-white border border-[#D1D1D1] hover:bg-[#FAF9F8] rounded-[4px] transition-colors"
          >
            <RefreshCw className={`w-3 h-3 ${loading ? 'animate-spin text-[#0F6CBD]' : ''}`} />
          </button>

          <button
            type="button"
            onClick={() => setExpanded((prev) => !prev)}
            className="inline-flex items-center gap-1 px-2 py-1 text-[11px] font-semibold text-[#0F6CBD] bg-[#EBF3FC] hover:bg-[#DCEBFA] rounded-[4px] transition-colors"
          >
            {expanded ? '접기' : '모델별 상세'}
            {expanded ? <ChevronUp className="w-3 h-3" /> : <ChevronDown className="w-3 h-3" />}
          </button>
        </div>
      </div>

      {/* Compact Model Breakdown Pills (always visible if models exist, 1 compact line) */}
      {byModel.length > 0 && !expanded && (
        <div className="mt-2 pt-2 border-t border-[#F0F0F0] flex flex-wrap items-center gap-1.5 text-[10px]">
          <span className="text-[#707070] font-medium mr-1">모델별:</span>
          {byModel.slice(0, 5).map((m) => (
            <span
              key={`${m.provider}-${m.modelId}`}
              className={`inline-flex items-center gap-1 px-2 py-0.5 rounded border font-mono ${tierBadge(m.tier)}`}
            >
              <strong className="font-semibold">{m.modelId}</strong>
              <span>· {formatNum(m.calls)}회</span>
              <span>· {formatNum(m.totalTokens)} tok</span>
              <span>· {formatUsd(m.costUsd)}</span>
            </span>
          ))}
        </div>
      )}

      {/* Expandable Compact Table for Per-Model & Recent Calls */}
      {expanded && (
        <div className="mt-3 pt-3 border-t border-[#EDEBE9] grid grid-cols-1 lg:grid-cols-2 gap-3 text-[11px]">
          {/* Left: Model Breakdown Table */}
          <div className="border border-[#E0E0E0] rounded-[6px] overflow-hidden">
            <div className="px-3 py-1.5 bg-[#FAF9F8] border-b border-[#EDEBE9] flex items-center justify-between">
              <span className="font-bold text-[#242424]">모델별 토큰 · 비용 집계</span>
              <button
                type="button"
                onClick={() => void handleReset()}
                disabled={loading || totals.totalCalls === 0}
                className="inline-flex items-center gap-1 text-[10px] text-[#616161] hover:text-red-600 disabled:opacity-40"
              >
                <RotateCcw className="w-2.5 h-2.5" />
                카운터 초기화
              </button>
            </div>
            {byModel.length === 0 ? (
              <div className="p-3 text-center text-[#707070]">
                아직 기록된 LLM API 호출 내역이 없습니다.
              </div>
            ) : (
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="border-b border-[#EDEBE9] text-[10px] text-[#616161] bg-white">
                      <th className="py-1.5 px-2.5 font-semibold">모델 ID</th>
                      <th className="py-1.5 px-2 font-semibold text-right">호출</th>
                      <th className="py-1.5 px-2 font-semibold text-right">입력/출력 토큰</th>
                      <th className="py-1.5 px-2 font-semibold text-right">추정 비용</th>
                      <th className="py-1.5 px-2.5 font-semibold text-right">지연</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[#F0F0F0] font-mono text-[10px]">
                    {byModel.map((m) => (
                      <tr key={`${m.provider}-${m.modelId}`} className="hover:bg-[#FAF9F8]">
                        <td className="py-1.5 px-2.5">
                          <div className="flex items-center gap-1.5">
                            <span
                              className={`px-1.5 py-0.2 rounded border text-[9px] uppercase ${tierBadge(m.tier)}`}
                            >
                              {m.tier}
                            </span>
                            <span className="font-semibold text-[#242424]">{m.modelId}</span>
                          </div>
                        </td>
                        <td className="py-1.5 px-2 text-right text-[#242424]">
                          {formatNum(m.calls)}
                        </td>
                        <td className="py-1.5 px-2 text-right text-[#424242]">
                          {formatNum(m.inputTokens)} / {formatNum(m.outputTokens)}
                        </td>
                        <td className="py-1.5 px-2 text-right font-semibold text-emerald-700">
                          {formatUsd(m.costUsd)}
                        </td>
                        <td className="py-1.5 px-2.5 text-right text-[#616161]">
                          {formatNum(m.avgLatencyMs)}ms
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>

          {/* Right: Recent Calls Feed */}
          <div className="border border-[#E0E0E0] rounded-[6px] overflow-hidden">
            <div className="px-3 py-1.5 bg-[#FAF9F8] border-b border-[#EDEBE9] flex items-center justify-between">
              <span className="font-bold text-[#242424]">최근 LLM 호출 로그 (최신 8건)</span>
              <span className="text-[10px] text-[#707070] font-mono">
                data/llm-usage-ledger.json
              </span>
            </div>
            {recentEvents.length === 0 ? (
              <div className="p-3 text-center text-[#707070]">
                최근 LLM 호출 기록이 없습니다.
              </div>
            ) : (
              <div className="divide-y divide-[#F0F0F0] max-h-[155px] overflow-y-auto">
                {recentEvents.map((ev) => (
                  <div
                    key={ev.id}
                    className="px-2.5 py-1.5 flex items-center justify-between gap-2 hover:bg-[#FAF9F8] text-[10px]"
                  >
                    <div className="flex items-center gap-1.5 min-w-0">
                      <span
                        className={`w-1.5 h-1.5 rounded-full shrink-0 ${
                          ev.status === 'ok' ? 'bg-emerald-500' : 'bg-red-500'
                        }`}
                      />
                      <span className="font-mono font-semibold text-[#242424] truncate">
                        {ev.modelId}
                      </span>
                      <span className="text-[#616161] truncate">· {ev.feature}</span>
                    </div>
                    <div className="flex items-center gap-2 shrink-0 font-mono text-[#616161]">
                      <span>{formatNum(ev.totalTokens)} tok</span>
                      <span className="text-emerald-700">{formatUsd(ev.costUsd)}</span>
                      <span>{ev.latencyMs}ms</span>
                      <span className="text-[#8A8886]">{ev.timestamp.slice(11, 19)}</span>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      )}
    </section>
  );
};
