import React, { useCallback, useEffect, useState } from 'react';
import {
  Activity,
  BarChart3,
  Check,
  ChevronDown,
  ChevronUp,
  Cpu,
  Gauge,
  RefreshCw,
  RotateCcw,
  Sliders,
} from 'lucide-react';
import type {
  BudgetAlertLevel,
  LlmProviderFamily,
  LlmUsageBudgetStatus,
  LlmUsageLimits,
  LlmUsageSummary,
} from '../services/llmUsageTracker';

export interface AdminLlmRoutingPolicy {
  groqSlotsConfigured: number;
  groqSlotsHealthy: number;
  groqParallelMax: number;
  claudeConfigured: boolean;
  geminiConfigured: boolean;
  defaultFastProvider: string;
  advancedAnalysisProvider: string;
}

interface LlmUsageApiResponse extends LlmUsageSummary {
  limits?: LlmUsageLimits;
  budgetStatus?: LlmUsageBudgetStatus;
  routingPolicy?: AdminLlmRoutingPolicy;
}

export interface AdminLlmUsageCompactCardProps {
  defaultExpanded?: boolean;
}

const FALLBACK_LIMITS: LlmUsageLimits = {
  dailyCallLimit: 300,
  dailyTokenLimit: 500_000,
  monthlyCostBudgetUsd: 15,
  warningThresholdPct: 75,
  criticalThresholdPct: 95,
  groqParallelMax: 5,
  claudeAdvancedMode: 'auto',
  updatedAt: '',
};

const PROVIDER_META: Array<{
  id: LlmProviderFamily;
  label: string;
  barColor: string;
}> = [
  { id: 'groq', label: 'Groq (고속·병렬)', barColor: '#F55036' },
  { id: 'claude', label: 'Claude 5.5 (고급·심층)', barColor: '#7C3AED' },
  { id: 'gemini', label: 'Gemini (백업·멀티)', barColor: '#0F6CBD' },
  { id: 'openai', label: 'OpenAI (보조)', barColor: '#0D9488' },
  { id: 'agent-builder', label: 'Agent Builder RAG', barColor: '#475569' },
];

const LIMIT_PRESETS: Array<{
  label: string;
  patch: Partial<LlmUsageLimits>;
}> = [
  {
    label: '절약형 (100회/$5)',
    patch: {
      dailyCallLimit: 100,
      dailyTokenLimit: 200_000,
      monthlyCostBudgetUsd: 5,
      warningThresholdPct: 70,
      criticalThresholdPct: 90,
      groqParallelMax: 3,
      claudeAdvancedMode: 'auto',
    },
  },
  {
    label: '표준 업무 (300회/$15)',
    patch: {
      dailyCallLimit: 300,
      dailyTokenLimit: 500_000,
      monthlyCostBudgetUsd: 15,
      warningThresholdPct: 75,
      criticalThresholdPct: 95,
      groqParallelMax: 5,
      claudeAdvancedMode: 'auto',
    },
  },
  {
    label: '대용량 집중 (1,000회/$40)',
    patch: {
      dailyCallLimit: 1000,
      dailyTokenLimit: 2_000_000,
      monthlyCostBudgetUsd: 40,
      warningThresholdPct: 80,
      criticalThresholdPct: 95,
      groqParallelMax: 5,
      claudeAdvancedMode: 'claude-only',
    },
  },
];

function levelStyle(level: BudgetAlertLevel): {
  bar: string;
  badge: string;
  label: string;
} {
  if (level === 'critical') {
    return {
      bar: 'bg-rose-600',
      badge: 'bg-rose-50 text-rose-700 border-rose-200',
      label: '위험',
    };
  }
  if (level === 'warning') {
    return {
      bar: 'bg-amber-500',
      badge: 'bg-amber-50 text-amber-800 border-amber-200',
      label: '주의',
    };
  }
  return {
    bar: 'bg-emerald-600',
    badge: 'bg-emerald-50 text-emerald-700 border-emerald-200',
    label: '정상',
  };
}

export const AdminLlmUsageCompactCard: React.FC<AdminLlmUsageCompactCardProps> = ({
  defaultExpanded = false,
}) => {
  const [summary, setSummary] = useState<LlmUsageApiResponse | null>(null);
  const [loading, setLoading] = useState(false);
  const [expanded, setExpanded] = useState(defaultExpanded);
  const [limitsEditorOpen, setLimitsEditorOpen] = useState(false);
  const [savingLimits, setSavingLimits] = useState(false);
  const [limitsNotice, setLimitsNotice] = useState<string | null>(null);
  const [draftLimits, setDraftLimits] = useState<LlmUsageLimits>(FALLBACK_LIMITS);

  const fetchUsage = useCallback(async () => {
    setLoading(true);
    try {
      const res = await fetch('/api/admin/llm-usage?limit=12');
      if (res.ok) {
        const data = (await res.json()) as LlmUsageApiResponse;
        setSummary(data);
        if (data.limits) {
          setDraftLimits(data.limits);
        }
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

  const handleSaveLimits = async () => {
    setSavingLimits(true);
    setLimitsNotice(null);
    try {
      const res = await fetch('/api/admin/llm-usage/limits', {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(draftLimits),
      });
      if (res.ok) {
        const data = (await res.json()) as {
          ok?: boolean;
          limits?: LlmUsageLimits;
          budgetStatus?: LlmUsageBudgetStatus;
        };
        if (data.limits) {
          setDraftLimits(data.limits);
          setSummary((prev) =>
            prev
              ? {
                  ...prev,
                  limits: data.limits,
                  budgetStatus: data.budgetStatus ?? prev.budgetStatus,
                }
              : prev
          );
        }
        await fetchUsage();
        setLimitsNotice('관리자 사용량 기준을 즉시 저장·반영했습니다.');
      } else {
        setLimitsNotice('사용량 기준 저장에 실패했습니다.');
      }
    } catch {
      setLimitsNotice('네트워크 오류로 기준을 저장하지 못했습니다.');
    } finally {
      setSavingLimits(false);
    }
  };

  const handleReset = async () => {
    setLoading(true);
    try {
      const res = await fetch('/api/admin/llm-usage/reset', { method: 'POST' });
      if (res.ok) {
        await fetchUsage();
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
  const limits = summary?.limits || draftLimits;
  const budget = summary?.budgetStatus || {
    month: today.date.slice(0, 7),
    dailyCallsUsed: today.calls,
    dailyCallsLimit: limits.dailyCallLimit,
    dailyCallsPct: Math.round((today.calls / Math.max(1, limits.dailyCallLimit)) * 100),
    dailyCallsLevel: 'ok' as BudgetAlertLevel,
    dailyTokensUsed: today.totalTokens,
    dailyTokensLimit: limits.dailyTokenLimit,
    dailyTokensPct: Math.round((today.totalTokens / Math.max(1, limits.dailyTokenLimit)) * 100),
    dailyTokensLevel: 'ok' as BudgetAlertLevel,
    monthlyCostUsd: totals.totalCostUsd,
    monthlyCostBudgetUsd: limits.monthlyCostBudgetUsd,
    monthlyCostPct: Math.round((totals.totalCostUsd / Math.max(0.01, limits.monthlyCostBudgetUsd)) * 100),
    monthlyCostLevel: 'ok' as BudgetAlertLevel,
    overallStatus: 'ok' as BudgetAlertLevel,
  };

  const byModel = summary?.byModel || [];
  const byProvider = summary?.byProvider;
  const dailyBuckets = summary?.dailyBuckets || [];
  const recentEvents = summary?.recentEvents || [];
  const policy = summary?.routingPolicy;

  const providerRows = PROVIDER_META.map((meta) => {
    const stat = byProvider?.[meta.id] || { calls: 0, totalTokens: 0, costUsd: 0 };
    return { ...meta, ...stat };
  });
  const totalProviderCalls = Math.max(
    1,
    providerRows.reduce((sum, r) => sum + r.calls, 0)
  );
  const chartBuckets = dailyBuckets.slice(-14);
  const maxDailyCalls = Math.max(1, ...chartBuckets.map((b) => b.calls));
  const limitFitsInChart = limits.dailyCallLimit <= Math.max(maxDailyCalls * 2.5, 25);
  const chartScaleMax = limitFitsInChart
    ? Math.max(1, maxDailyCalls, limits.dailyCallLimit)
    : maxDailyCalls;

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

  const overallBadge = levelStyle(budget.overallStatus);
  const gaugeItems = [
    {
      key: 'calls',
      label: '일일 호출량 기준',
      usedText: `${formatNum(budget.dailyCallsUsed)}회`,
      limitText: `${formatNum(budget.dailyCallsLimit)}회/일`,
      pct: budget.dailyCallsPct,
      level: budget.dailyCallsLevel,
    },
    {
      key: 'tokens',
      label: '일일 토큰 한도',
      usedText: `${formatNum(budget.dailyTokensUsed)} tok`,
      limitText: `${formatNum(budget.dailyTokensLimit)} tok/일`,
      pct: budget.dailyTokensPct,
      level: budget.dailyTokensLevel,
    },
    {
      key: 'cost',
      label: `월간 비용 예산 (${budget.month})`,
      usedText: formatUsd(budget.monthlyCostUsd),
      limitText: `$${budget.monthlyCostBudgetUsd.toFixed(2)}/월`,
      pct: budget.monthlyCostPct,
      level: budget.monthlyCostLevel,
    },
  ];

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
                LLM 토큰 · 모델 사용량 및 관리자 기준 모니터
              </h2>
              <span className={`text-[10px] font-bold px-1.5 py-0.5 rounded border ${overallBadge.badge}`}>
                기준 상태: {overallBadge.label}
              </span>
            </div>
          </div>
        </div>

        {/* Inline Mini Metrics Pills + Admin Limit Toggle */}
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
            <span className="text-[#616161]">오늘({today.date.slice(5)})</span>
            <strong className="font-mono text-[#0F6CBD]">
              {formatNum(today.calls)}/{formatNum(limits.dailyCallLimit)}회 ({budget.dailyCallsPct}%)
            </strong>
          </div>

          <div className="px-2.5 py-1 rounded-[4px] bg-[#FAF9F8] border border-[#EDEBE9] flex items-center gap-1.5">
            <span className="text-[#616161]">월간 비용</span>
            <strong className="font-mono text-emerald-700">
              {formatUsd(budget.monthlyCostUsd)} / ${limits.monthlyCostBudgetUsd.toFixed(0)}
            </strong>
          </div>

          <div className="px-2 py-1 rounded-[4px] bg-[#FAF9F8] border border-[#EDEBE9] flex items-center gap-1 text-[10px] text-[#616161] font-mono">
            <Activity className="w-3 h-3 text-[#0F6CBD]" />
            평균 {formatNum(totals.avgLatencyMs)}ms
          </div>

          <button
            type="button"
            onClick={() => setLimitsEditorOpen((v) => !v)}
            aria-expanded={limitsEditorOpen}
            className={`inline-flex items-center gap-1 px-2.5 py-1 text-[11px] font-semibold rounded-[4px] border transition-colors ${
              limitsEditorOpen
                ? 'bg-[#0F6CBD] text-white border-[#0F6CBD]'
                : 'bg-[#FAF9F8] text-[#242424] border-[#C8C6C4] hover:bg-[#EBF3FC] hover:border-[#0F6CBD]'
            }`}
          >
            <Sliders className="w-3 h-3" />
            <span>기준 수시 변경</span>
          </button>

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

      {/* Collapsible Anytime Admin Usage Limits & Routing Policy Editor */}
      {limitsEditorOpen && (
        <div
          data-testid="llm-limits-editor"
          className="mt-2.5 p-3 rounded-[6px] border border-[#B4D6FA] bg-[#F8FBFF] space-y-2.5 text-xs"
        >
          <div className="flex flex-wrap items-center justify-between gap-2">
            <span className="font-bold text-[#0F6CBD] flex items-center gap-1.5">
              <Sliders className="w-3.5 h-3.5" />
              관리자 사용량 기준 · 병렬 · 라우팅 수시 설정 (저장 즉시 적용)
            </span>
            <div className="flex flex-wrap items-center gap-1">
              {LIMIT_PRESETS.map((preset) => (
                <button
                  key={preset.label}
                  type="button"
                  onClick={() => setDraftLimits((prev) => ({ ...prev, ...preset.patch }))}
                  className="px-2 py-0.5 rounded bg-white border border-[#C8C6C4] hover:border-[#0F6CBD] text-[10.5px] font-medium text-[#242424]"
                >
                  {preset.label}
                </button>
              ))}
            </div>
          </div>

          <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-7 gap-2 text-[11px]">
            <label className="space-y-0.5">
              <span className="block text-[10px] font-semibold text-[#424242]">일일 호출 상한</span>
              <input
                id="limit-daily-calls"
                type="number"
                min={10}
                max={50000}
                value={draftLimits.dailyCallLimit}
                onChange={(e) =>
                  setDraftLimits((p) => ({ ...p, dailyCallLimit: Number(e.target.value) || 10 }))
                }
                className="w-full px-2 py-1 rounded border border-[#C8C6C4] bg-white font-mono"
              />
            </label>
            <label className="space-y-0.5">
              <span className="block text-[10px] font-semibold text-[#424242]">일일 토큰 상한</span>
              <input
                id="limit-daily-tokens"
                type="number"
                min={1000}
                max={50000000}
                step={10000}
                value={draftLimits.dailyTokenLimit}
                onChange={(e) =>
                  setDraftLimits((p) => ({ ...p, dailyTokenLimit: Number(e.target.value) || 1000 }))
                }
                className="w-full px-2 py-1 rounded border border-[#C8C6C4] bg-white font-mono"
              />
            </label>
            <label className="space-y-0.5">
              <span className="block text-[10px] font-semibold text-[#424242]">월간 비용 예산($)</span>
              <input
                id="limit-monthly-usd"
                type="number"
                min={0.5}
                max={5000}
                step={0.5}
                value={draftLimits.monthlyCostBudgetUsd}
                onChange={(e) =>
                  setDraftLimits((p) => ({
                    ...p,
                    monthlyCostBudgetUsd: Number(e.target.value) || 0.5,
                  }))
                }
                className="w-full px-2 py-1 rounded border border-[#C8C6C4] bg-white font-mono"
              />
            </label>
            <label className="space-y-0.5">
              <span className="block text-[10px] font-semibold text-[#424242]">주의 임계치(%)</span>
              <input
                id="limit-warn-pct"
                type="number"
                min={10}
                max={95}
                value={draftLimits.warningThresholdPct}
                onChange={(e) =>
                  setDraftLimits((p) => ({ ...p, warningThresholdPct: Number(e.target.value) || 75 }))
                }
                className="w-full px-2 py-1 rounded border border-[#C8C6C4] bg-white font-mono"
              />
            </label>
            <label className="space-y-0.5">
              <span className="block text-[10px] font-semibold text-[#424242]">위험 임계치(%)</span>
              <input
                id="limit-crit-pct"
                type="number"
                min={20}
                max={100}
                value={draftLimits.criticalThresholdPct}
                onChange={(e) =>
                  setDraftLimits((p) => ({ ...p, criticalThresholdPct: Number(e.target.value) || 95 }))
                }
                className="w-full px-2 py-1 rounded border border-[#C8C6C4] bg-white font-mono"
              />
            </label>
            <label className="space-y-0.5">
              <span className="block text-[10px] font-semibold text-[#424242]">Groq 병렬(1~5)</span>
              <input
                id="limit-groq-parallel"
                type="number"
                min={1}
                max={5}
                value={draftLimits.groqParallelMax}
                onChange={(e) =>
                  setDraftLimits((p) => ({ ...p, groqParallelMax: Number(e.target.value) || 5 }))
                }
                className="w-full px-2 py-1 rounded border border-[#C8C6C4] bg-white font-mono"
              />
            </label>
            <label className="space-y-0.5">
              <span className="block text-[10px] font-semibold text-[#424242]">고급 분석 라우팅</span>
              <select
                id="limit-claude-mode"
                value={draftLimits.claudeAdvancedMode}
                onChange={(e) =>
                  setDraftLimits((p) => ({
                    ...p,
                    claudeAdvancedMode: e.target.value as LlmUsageLimits['claudeAdvancedMode'],
                  }))
                }
                className="w-full px-1.5 py-1 rounded border border-[#C8C6C4] bg-white text-[11px]"
              >
                <option value="auto">자동 (고급=Claude)</option>
                <option value="claude-only">항상 Claude 우선</option>
                <option value="groq-only">Groq 5-키 최대 절약</option>
              </select>
            </label>
          </div>

          <div className="flex flex-wrap items-center justify-between gap-2 pt-1">
            <span className="text-[11px] text-[#616161]">
              {limitsNotice ||
                `현재 주의 ${limits.warningThresholdPct}% / 위험 ${limits.criticalThresholdPct}% 임계치로 게이지와 도표가 자동 판정됩니다.`}
            </span>
            <button
              type="button"
              disabled={savingLimits}
              onClick={() => void handleSaveLimits()}
              data-testid="save-llm-limits-btn"
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded bg-[#0F6CBD] hover:bg-[#115EA3] disabled:opacity-60 text-white text-[11px] font-semibold"
            >
              <Check className="w-3.5 h-3.5" />
              {savingLimits ? '저장 중...' : '기준 즉시 저장'}
            </button>
          </div>
        </div>
      )}

      {/* Always-Visible Admin Budget Gauges + Threshold Comparison Table */}
      <div className="mt-2.5 pt-2.5 border-t border-[#F0F0F0] space-y-2.5">
        {policy && (
          <div className="flex flex-wrap items-center gap-1.5 text-[11px]">
            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-[4px] border bg-orange-50 border-orange-200 text-orange-900 font-semibold">
              <span>고속·병렬 처리:</span>
              <span>{policy.defaultFastProvider}</span>
            </span>
            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-[4px] border bg-purple-50 border-purple-200 text-purple-900 font-semibold">
              <span>고급·심층 분석:</span>
              <span>{policy.advancedAnalysisProvider}</span>
            </span>
            <span
              className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-[4px] border text-[10px] font-mono ${
                policy.groqSlotsConfigured > 0
                  ? 'bg-emerald-50 border-emerald-200 text-emerald-800'
                  : 'bg-slate-50 border-slate-200 text-slate-600'
              }`}
            >
              Groq 키 슬롯: {policy.groqSlotsHealthy}/{policy.groqSlotsConfigured} 활성 (최대{' '}
              {policy.groqParallelMax}병렬)
            </span>
          </div>
        )}

        {/* Threshold Gauge Bars + Reference Comparison Table */}
        <div
          data-testid="llm-budget-panel"
          className="grid grid-cols-1 lg:grid-cols-2 gap-3 text-[11px]"
        >
          {/* Gauge Graph vs Admin Thresholds */}
          <div className="p-2.5 rounded-[6px] border border-[#E5E5E5] bg-[#FAF9F8] space-y-2">
            <div className="flex items-center justify-between">
              <span className="font-bold text-[#242424] flex items-center gap-1.5">
                <Gauge className="w-3.5 h-3.5 text-[#0F6CBD]" />
                사용량 기준 대비 실시간 게이지 (주의 {limits.warningThresholdPct}% · 위험{' '}
                {limits.criticalThresholdPct}%)
              </span>
              <span className={`px-1.5 py-0.2 rounded border text-[10px] font-bold ${overallBadge.badge}`}>
                {overallBadge.label}
              </span>
            </div>
            <div className="space-y-2">
              {gaugeItems.map((g) => {
                const st = levelStyle(g.level);
                const fillW = Math.min(100, Math.max(g.pct > 0 ? 3 : 0, g.pct));
                return (
                  <div key={g.key} className="space-y-0.5">
                    <div className="flex items-center justify-between text-[10px]">
                      <span className="font-semibold text-[#242424]">{g.label}</span>
                      <span className="font-mono text-[#424242]">
                        {g.usedText} / <strong className="text-[#0F6CBD]">{g.limitText}</strong> ({g.pct}%)
                      </span>
                    </div>
                    <div className="relative w-full h-2.5 rounded-full bg-white border border-[#D1D1D1] overflow-hidden">
                      <div
                        className={`h-full rounded-full transition-all duration-300 ${st.bar}`}
                        style={{ width: `${fillW}%` }}
                      />
                      <span
                        title={`주의 임계치 ${limits.warningThresholdPct}%`}
                        style={{ left: `${limits.warningThresholdPct}%` }}
                        className="absolute top-0 bottom-0 w-0.5 bg-amber-600/80"
                      />
                      <span
                        title={`위험 임계치 ${limits.criticalThresholdPct}%`}
                        style={{ left: `${Math.min(99, limits.criticalThresholdPct)}%` }}
                        className="absolute top-0 bottom-0 w-0.5 bg-rose-700"
                      />
                    </div>
                  </div>
                );
              })}
            </div>
          </div>

          {/* Admin Threshold Comparison Table */}
          <div className="rounded-[6px] border border-[#E5E5E5] bg-white overflow-hidden">
            <div className="px-2.5 py-1.5 bg-[#FAF9F8] border-b border-[#EDEBE9] flex items-center justify-between">
              <span className="font-bold text-[#242424]">관리자 사용량 기준 도표</span>
              <span className="text-[10px] text-[#616161] font-mono">
                {limits.updatedAt ? `수정: ${limits.updatedAt.slice(0, 16).replace('T', ' ')}` : '기본 기준'}
              </span>
            </div>
            <table data-testid="llm-budget-table" className="w-full text-left border-collapse text-[10.5px]">
              <thead>
                <tr className="border-b border-[#EDEBE9] text-[#616161] bg-[#FAF9F8]/60">
                  <th className="py-1 px-2.5 font-semibold">기준 항목</th>
                  <th className="py-1 px-2 font-semibold text-right">현재</th>
                  <th className="py-1 px-2 font-semibold text-right">설정 기준</th>
                  <th className="py-1 px-2.5 font-semibold text-right">상태</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-[#F0F0F0] font-mono">
                {gaugeItems.map((g) => {
                  const st = levelStyle(g.level);
                  return (
                    <tr key={g.key} className="hover:bg-[#FAF9F8]">
                      <td className="py-1 px-2.5 font-sans font-medium text-[#242424]">{g.label}</td>
                      <td className="py-1 px-2 text-right text-[#424242]">{g.usedText}</td>
                      <td className="py-1 px-2 text-right font-semibold text-[#0F6CBD]">{g.limitText}</td>
                      <td className="py-1 px-2.5 text-right">
                        <span className={`px-1.5 py-0.2 rounded border text-[9.5px] font-sans font-bold ${st.badge}`}>
                          {g.pct}% · {st.label}
                        </span>
                      </td>
                    </tr>
                  );
                })}
                <tr className="hover:bg-[#FAF9F8]">
                  <td className="py-1 px-2.5 font-sans font-medium text-[#242424]">병렬 · 고급 라우팅</td>
                  <td className="py-1 px-2 text-right text-[#424242]">
                    {limits.claudeAdvancedMode === 'auto'
                      ? '고급=Claude'
                      : limits.claudeAdvancedMode === 'claude-only'
                        ? 'Claude 강제'
                        : 'Groq 절약'}
                  </td>
                  <td className="py-1 px-2 text-right font-semibold text-[#0F6CBD]">
                    Groq {limits.groqParallelMax}병렬
                  </td>
                  <td className="py-1 px-2.5 text-right">
                    <span className="px-1.5 py-0.2 rounded border text-[9.5px] font-sans font-semibold bg-[#EBF3FC] text-[#0F6CBD] border-[#B4D6FA]">
                      수시 변경
                    </span>
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>

        {/* Two Side-by-Side Call Volume Charts (Always Visible) */}
        <div
          data-testid="llm-usage-graphs"
          className="grid grid-cols-1 lg:grid-cols-2 gap-3 text-[11px]"
        >
          {/* Graph 1: Provider Call Volume Share Bars */}
          <div className="p-2.5 rounded-[6px] border border-[#E5E5E5] bg-[#FAF9F8]">
            <div className="flex items-center justify-between mb-2">
              <span className="font-bold text-[#242424] flex items-center gap-1.5">
                <BarChart3 className="w-3.5 h-3.5 text-[#0F6CBD]" />
                프로바이더별 AI API 호출량 분포
              </span>
              <span className="text-[10px] font-mono text-[#616161]">
                총 {formatNum(totals.totalCalls)}회
              </span>
            </div>
            <div className="space-y-1.5">
              {providerRows.map((row) => {
                const pct = totals.totalCalls > 0 ? Math.round((row.calls / totalProviderCalls) * 100) : 0;
                const widthPct = row.calls > 0 ? Math.max(4, pct) : 0;
                return (
                  <div key={row.id} className="space-y-0.5">
                    <div className="flex items-center justify-between text-[10px]">
                      <span className="font-semibold text-[#242424]">{row.label}</span>
                      <span className="font-mono text-[#424242]">
                        {formatNum(row.calls)}회 ({pct}%) · {formatNum(row.totalTokens)} tok ·{' '}
                        {formatUsd(row.costUsd)}
                      </span>
                    </div>
                    <div className="w-full h-2 rounded-full bg-white border border-[#E1DFDD] overflow-hidden">
                      <div
                        className="h-full rounded-full transition-all duration-300"
                        style={{ width: `${widthPct}%`, backgroundColor: row.barColor }}
                      />
                    </div>
                  </div>
                );
              })}
            </div>
          </div>

          {/* Graph 2: Daily Call Volume SVG Bar Chart + Daily Limit Reference Line */}
          <div className="p-2.5 rounded-[6px] border border-[#E5E5E5] bg-[#FAF9F8] flex flex-col justify-between">
            <div className="flex items-center justify-between mb-1.5">
              <span className="font-bold text-[#242424] flex items-center gap-1.5">
                <Activity className="w-3.5 h-3.5 text-purple-700" />
                일자별 API 호출 추이 그래프 (최근 {chartBuckets.length || 0}일)
              </span>
              <span className="text-[10px] font-mono text-[#616161]">
                기준선 {formatNum(limits.dailyCallLimit)}회/일 · 최대 {formatNum(maxDailyCalls)}회
              </span>
            </div>
            {chartBuckets.length === 0 ? (
              <div className="h-[92px] flex items-center justify-center text-[11px] text-[#707070] bg-white rounded border border-dashed border-[#D1D1D1]">
                AI 분석 또는 메일 요약을 실행하면 일별 API 호출량 그래프가 표시됩니다.
              </div>
            ) : (
              <div className="bg-white rounded border border-[#E1DFDD] px-2 pt-2 pb-1">
                <svg
                  viewBox="0 0 360 92"
                  role="img"
                  aria-label="일자별 AI API 호출량 그래프"
                  className="w-full h-[92px]"
                >
                  {/* Threshold Reference Line */}
                  {(() => {
                    const lineY = limitFitsInChart
                      ? Math.max(10, 68 - Math.round((limits.dailyCallLimit / chartScaleMax) * 54))
                      : 10;
                    return (
                      <g>
                        <line
                          x1={8}
                          y1={lineY}
                          x2={352}
                          y2={lineY}
                          stroke="#DC2626"
                          strokeWidth="1"
                          strokeDasharray="4 3"
                          opacity={0.75}
                        />
                        <text x={350} y={Math.max(8, lineY - 2)} textAnchor="end" fontSize="7.5" fill="#DC2626">
                          일일 기준 {formatNum(limits.dailyCallLimit)}회
                        </text>
                      </g>
                    );
                  })()}

                  {chartBuckets.map((bucket, idx) => {
                    const slotW = 340 / Math.max(1, chartBuckets.length);
                    const barW = Math.min(22, Math.max(8, slotW * 0.62));
                    const x = 10 + idx * slotW + (slotW - barW) / 2;
                    const barH = Math.max(4, Math.round((bucket.calls / chartScaleMax) * 54));
                    const y = 68 - barH;
                    const barColor =
                      bucket.calls >= limits.dailyCallLimit ? '#DC2626' : '#0F6CBD';
                    return (
                      <g key={bucket.date}>
                        <rect x={x} y={y} width={barW} height={barH} rx={2.5} fill={barColor}>
                          <title>
                            {bucket.date}: {bucket.calls}회 · {formatNum(bucket.totalTokens)} tok
                          </title>
                        </rect>
                        <text
                          x={x + barW / 2}
                          y={Math.max(10, y - 3)}
                          textAnchor="middle"
                          fontSize="8.5"
                          fill="#242424"
                          fontWeight="600"
                        >
                          {bucket.calls}
                        </text>
                        <text
                          x={x + barW / 2}
                          y={84}
                          textAnchor="middle"
                          fontSize="8"
                          fill="#616161"
                        >
                          {bucket.date.slice(5)}
                        </text>
                      </g>
                    );
                  })}
                </svg>
              </div>
            )}
          </div>
        </div>
      </div>

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
              <span className="font-bold text-[#242424]">최근 LLM 호출 로그 (최신 12건)</span>
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
