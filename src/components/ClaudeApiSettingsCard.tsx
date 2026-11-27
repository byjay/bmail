import React, { useCallback, useEffect, useState } from 'react';
import {
  CheckCircle2,
  Eye,
  EyeOff,
  KeyRound,
  Loader2,
  Save,
  Sparkles,
  Trash2,
  Zap,
} from 'lucide-react';
import type { AllowedClaudeModel } from '../utils/localClaudeEnv';

interface ClaudeConfigView {
  configured: boolean;
  maskedPreview: string;
  source: 'env' | 'emv' | 'unconfigured';
  claudeModel: AllowedClaudeModel;
  routingMode: 'claude-only' | 'auto' | 'groq-only';
  groqConfiguredCount: number;
}

export interface ClaudeApiSettingsCardProps {
  compact?: boolean;
  onUpdated?: () => void;
}

export const ClaudeApiSettingsCard: React.FC<ClaudeApiSettingsCardProps> = ({
  compact = false,
  onUpdated,
}) => {
  const [status, setStatus] = useState<ClaudeConfigView | null>(null);
  const [apiKeyInput, setApiKeyInput] = useState('');
  const [showKey, setShowKey] = useState(false);
  const [claudeModel, setClaudeModel] = useState<AllowedClaudeModel>('auto');
  const [routingMode, setRoutingMode] = useState<'claude-only' | 'auto' | 'groq-only'>('claude-only');
  const [saving, setSaving] = useState(false);
  const [testing, setTesting] = useState(false);
  const [clearing, setClearing] = useState(false);
  const [notice, setNotice] = useState<{ type: 'success' | 'error'; text: string } | null>(null);

  const fetchConfig = useCallback(async () => {
    try {
      const res = await fetch('/api/local/ai/claude-config');
      if (!res.ok) return;
      const data = (await res.json()) as ClaudeConfigView;
      setStatus(data);
      setClaudeModel(data.claudeModel || 'auto');
      setRoutingMode(data.routingMode || 'claude-only');
    } catch {
      // ignore offline
    }
  }, []);

  useEffect(() => {
    void fetchConfig();
  }, [fetchConfig]);

  const handleSave = async (e: React.FormEvent) => {
    e.preventDefault();
    setSaving(true);
    setNotice(null);
    try {
      const payload: Record<string, unknown> = {
        claudeModel,
        routingMode,
      };
      if (apiKeyInput.trim()) {
        payload.apiKey = apiKeyInput.trim();
      }
      const res = await fetch('/api/local/ai/claude-config', {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      const data = (await res.json().catch(() => ({}))) as ClaudeConfigView & {
        ok?: boolean;
        error?: string;
      };
      if (!res.ok) {
        setNotice({ type: 'error', text: data.error || `저장 실패 (HTTP ${res.status})` });
        return;
      }
      setStatus(data);
      setClaudeModel(data.claudeModel || claudeModel);
      setRoutingMode(data.routingMode || routingMode);
      setApiKeyInput('');
      setNotice({
        type: 'success',
        text: data.configured
          ? `Claude API 설정 저장 및 즉시 반영 완료 (${data.maskedPreview} · 모드: ${
              data.routingMode === 'claude-only'
                ? 'Claude API 우선'
                : data.routingMode === 'groq-only'
                  ? 'Groq 전용'
                  : '자동 혼합'
            })`
          : 'AI 엔진 라우팅 설정을 저장했습니다.',
      });
      onUpdated?.();
    } catch (err) {
      setNotice({
        type: 'error',
        text: err instanceof Error ? err.message : 'Claude API 설정 저장 중 오류가 발생했습니다.',
      });
    } finally {
      setSaving(false);
    }
  };

  const handleTest = async () => {
    setTesting(true);
    setNotice(null);
    try {
      const res = await fetch('/api/local/ai/claude-test', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          ...(apiKeyInput.trim() ? { apiKey: apiKeyInput.trim() } : {}),
          claudeModel,
        }),
      });
      const data = (await res.json().catch(() => ({}))) as {
        ok?: boolean;
        model?: string;
        latencyMs?: number;
        error?: string;
      };
      if (!res.ok || !data.ok) {
        setNotice({
          type: 'error',
          text: data.error || `Claude API 연결 테스트 실패 (HTTP ${res.status})`,
        });
        return;
      }
      setNotice({
        type: 'success',
        text: `Claude API 연결 성공 · 모델: ${data.model || 'Claude 5.5'} · 응답 ${data.latencyMs ?? 0}ms`,
      });
    } catch (err) {
      setNotice({
        type: 'error',
        text: err instanceof Error ? err.message : 'Claude API 연결 테스트 실패',
      });
    } finally {
      setTesting(false);
    }
  };

  const handleClearKey = async () => {
    setClearing(true);
    setNotice(null);
    try {
      const res = await fetch('/api/local/ai/claude-config', {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ clearApiKey: true }),
      });
      const data = (await res.json().catch(() => ({}))) as ClaudeConfigView & { error?: string };
      if (!res.ok) {
        setNotice({ type: 'error', text: data.error || '키 삭제 실패' });
        return;
      }
      setStatus(data);
      setApiKeyInput('');
      setNotice({ type: 'success', text: '저장된 Claude API 키를 삭제했습니다.' });
      onUpdated?.();
    } finally {
      setClearing(false);
    }
  };

  return (
    <section
      data-testid="claude-api-settings-card"
      aria-label="Claude API 직접 설정"
      className={`bg-white border border-[#B4D6FA] rounded-[8px] shadow-[0_1px_2px_rgba(0,0,0,0.04)] ${
        compact ? 'p-4 space-y-3.5' : 'p-6 space-y-5'
      }`}
    >
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 pb-3 border-b border-[#EDEBE9]">
        <div className="space-y-0.5">
          <div className="flex items-center gap-2">
            <Sparkles className="w-4 h-4 text-[#0F6CBD]" strokeWidth={2} />
            <h2 className="text-sm font-bold text-[#111827]">
              Claude API 직접 설정 (Anthropic · 메일 분석/답안/검증)
            </h2>
          </div>
          <p className="text-xs text-[#4B5563]">
            Groq 연동 없이도 Claude API 키(<code className="font-mono">sk-ant-...</code>)만 입력하면 즉시 실제 AI 메일 분석·답안·번역이 동작합니다. 키는 로컬 <code className="font-mono">.env</code>에만 저장되며 외부로 노출되지 않습니다.
          </p>
        </div>
        <div className="flex items-center gap-1.5 shrink-0 flex-wrap">
          <span
            data-testid="claude-config-status-badge"
            className={`inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-[11px] font-bold border ${
              status?.configured
                ? 'bg-emerald-50 border-emerald-200 text-emerald-800'
                : 'bg-amber-50 border-amber-200 text-amber-900'
            }`}
          >
            {status?.configured ? (
              <>
                <CheckCircle2 className="w-3.5 h-3.5" />
                <span>Claude 연결됨 ({status.maskedPreview})</span>
              </>
            ) : (
              <span>Claude API 키 미설정</span>
            )}
          </span>
          <span className="px-2 py-1 rounded-full text-[10.5px] font-semibold bg-slate-100 text-[#374151] border border-[#E1DFDD]">
            Groq 풀: {status?.groqConfiguredCount ?? 0}개
          </span>
        </div>
      </div>

      <form onSubmit={handleSave} className="space-y-4">
        <div className="grid grid-cols-1 md:grid-cols-2 gap-3.5">
          {/* Claude API Key Input */}
          <div className="md:col-span-2 space-y-1.5">
            <label
              htmlFor="claude-api-key-input"
              className="flex items-center gap-1.5 text-xs font-bold text-[#111827]"
            >
              <KeyRound className="w-3.5 h-3.5 text-[#0F6CBD]" />
              <span>Anthropic Claude API 키 (ANTHROPIC_API_KEY)</span>
            </label>
            <div className="relative flex items-center gap-2">
              <input
                id="claude-api-key-input"
                data-testid="claude-api-key-input"
                type={showKey ? 'text' : 'password'}
                value={apiKeyInput}
                onChange={(e) => setApiKeyInput(e.target.value)}
                autoComplete="off"
                spellCheck={false}
                placeholder={
                  status?.configured
                    ? `현재 저장됨 (${status.maskedPreview}) — 교체할 때만 새 sk-ant-... 키 입력`
                    : 'sk-ant-api03-...'
                }
                className="flex-1 px-3 py-2 text-xs font-mono text-[#111827] bg-white border border-[#D1D1D1] border-b-[#616161] rounded-[4px] focus:outline-none focus:border-[#0F6CBD] focus:ring-2 focus:ring-[#0F6CBD]"
              />
              <button
                type="button"
                onClick={() => setShowKey((v) => !v)}
                aria-label={showKey ? 'API 키 숨기기' : 'API 키 보기'}
                className="p-2 rounded-[4px] border border-[#D1D1D1] bg-white hover:bg-[#F3F2F1] text-[#374151]"
              >
                {showKey ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
              </button>
            </div>
          </div>

          {/* Routing Mode */}
          <div className="space-y-1.5">
            <label
              htmlFor="claude-routing-mode"
              className="block text-xs font-bold text-[#111827]"
            >
              기본 AI 분석 엔진 모드
            </label>
            <select
              id="claude-routing-mode"
              data-testid="claude-routing-mode-select"
              value={routingMode}
              onChange={(e) =>
                setRoutingMode(e.target.value as 'claude-only' | 'auto' | 'groq-only')
              }
              className="w-full px-3 py-2 text-xs text-[#111827] bg-white border border-[#D1D1D1] border-b-[#616161] rounded-[4px] focus:outline-none focus:border-[#0F6CBD] focus:ring-2 focus:ring-[#0F6CBD]"
            >
              <option value="claude-only">Claude API 우선 (권장 · Groq 미사용/단독 실행)</option>
              <option value="auto">자동 혼합 (일반 Groq + 긴급·정밀 Claude)</option>
              <option value="groq-only">Groq 5-키 전용</option>
            </select>
          </div>

          {/* Claude Model */}
          <div className="space-y-1.5">
            <label
              htmlFor="claude-model-select"
              className="block text-xs font-bold text-[#111827]"
            >
              Claude 모델 선택
            </label>
            <select
              id="claude-model-select"
              data-testid="claude-model-select"
              value={claudeModel}
              onChange={(e) => setClaudeModel(e.target.value as AllowedClaudeModel)}
              className="w-full px-3 py-2 text-xs text-[#111827] bg-white border border-[#D1D1D1] border-b-[#616161] rounded-[4px] focus:outline-none focus:border-[#0F6CBD] focus:ring-2 focus:ring-[#0F6CBD]"
            >
              <option value="auto">자동 난이도 라우팅 (Haiku / Sonnet / Opus 5.5)</option>
              <option value="claude-haiku-5-5">Claude Haiku 5.5 (초고속 · 저비용)</option>
              <option value="claude-sonnet-5-5">Claude Sonnet 5.5 (설계·선급 균형)</option>
              <option value="claude-opus-5-5">Claude Opus 5.5 (최고 정밀 분석)</option>
            </select>
          </div>
        </div>

        {notice && (
          <div
            data-testid="claude-config-notice"
            className={`px-3 py-2 rounded-[4px] text-xs border ${
              notice.type === 'success'
                ? 'bg-emerald-50 border-emerald-200 text-emerald-900'
                : 'bg-rose-50 border-rose-200 text-rose-900'
            }`}
          >
            {notice.text}
          </div>
        )}

        <div className="flex flex-wrap items-center gap-2">
          <button
            type="submit"
            data-testid="btn-save-claude-config"
            disabled={saving}
            className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-[4px] bg-[#0F6CBD] hover:bg-[#115EA3] text-white text-xs font-semibold disabled:opacity-60"
          >
            {saving ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <Save className="w-3.5 h-3.5" />}
            <span>{saving ? '저장 중...' : '저장 및 즉시 적용'}</span>
          </button>

          <button
            type="button"
            data-testid="btn-test-claude-config"
            disabled={testing || (!status?.configured && !apiKeyInput.trim())}
            onClick={() => void handleTest()}
            className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-[4px] border border-[#0F6CBD] bg-[#EBF3FC] hover:bg-[#DCEBFA] text-[#0F6CBD] text-xs font-semibold disabled:opacity-50"
          >
            {testing ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <Zap className="w-3.5 h-3.5" />}
            <span>{testing ? '연결 확인 중...' : 'Claude 연결 테스트'}</span>
          </button>

          {status?.configured && (
            <button
              type="button"
              disabled={clearing}
              onClick={() => void handleClearKey()}
              className="inline-flex items-center gap-1 px-2.5 py-2 rounded-[4px] border border-rose-200 bg-white hover:bg-rose-50 text-rose-700 text-xs font-medium disabled:opacity-50 ml-auto"
            >
              <Trash2 className="w-3.5 h-3.5" />
              <span>키 삭제</span>
            </button>
          )}
        </div>
      </form>
    </section>
  );
};
