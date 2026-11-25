import type {
  AgentExecutionStep,
  AiProviderId,
  AnalysisOptions,
  AnalysisResult,
  AppConfig,
  MailItem,
} from '../types/mail';
import { buildDeterministicWorkLog, type ProjectDefinition } from '../utils/deterministicAnalyzer';
import { filterEmails, getAvailableDates, maskSensitiveData, SCOPE_LABELS } from '../utils/mailFilter';
import { analyzeEmailsWithProvider, getProviderStatuses } from './aiService';

import type { EmvKeyName } from './emvVaultService';

export interface AutonomousAgentRequest {
  emails: MailItem[];
  options: AnalysisOptions;
  config: AppConfig;
  projects?: ProjectDefinition[];
  agentGoal?: string;
  autoPickLatestDateIfEmpty?: boolean;
  onStepUpdate?: (steps: AgentExecutionStep[]) => void;
  envOverrides?: Partial<Record<EmvKeyName, string>>;
}

/**
 * Masks sensitive fields on shallow copies of emails for offline deterministic analysis.
 */
export function sanitizeEmailsForOffline(emails: MailItem[], shouldMask: boolean): MailItem[] {
  if (!shouldMask) return emails;
  return emails.map((mail) => ({
    ...mail,
    subject: maskSensitiveData(mail.subject),
    snippet: maskSensitiveData(mail.snippet),
    body: maskSensitiveData(mail.body),
    fromName: maskSensitiveData(mail.fromName),
    fromEmail: maskSensitiveData(mail.fromEmail),
  }));
}

export interface PrivacyAuditSummary {
  redactedCount: number;
  sanitizedEmails: MailItem[];
}

/**
 * Audits and redacts sensitive items across emails, returning exact count of redactions.
 */
export function auditAndMaskEmails(emails: MailItem[], shouldMask: boolean): PrivacyAuditSummary {
  if (!shouldMask) {
    return { redactedCount: 0, sanitizedEmails: emails };
  }

  let redactedCount = 0;
  const sanitizedEmails = emails.map((mail) => {
    const maskedBody = maskSensitiveData(mail.body);
    const markerRe = /\[(이메일 비공개|연락처 비공개|금액 비공개|IP 비공개|API키 비공개)\]/g;
    const bodyMarkers = (maskedBody.match(markerRe) ?? []).length;
    const rawMarkers = (mail.body.match(markerRe) ?? []).length;
    redactedCount += Math.max(0, bodyMarkers - rawMarkers);
    return mail;
  });

  return { redactedCount, sanitizedEmails };
}

/**
 * Resolves the best available AI provider based on user choice, server secrets, and config.json fallback rules.
 */
export function resolveAgentProvider(
  requested: AiProviderId,
  config: AppConfig,
  envOverrides?: Partial<Record<EmvKeyName, string>>
): { provider: AiProviderId; switchedReason: string | null } {
  const statuses = getProviderStatuses(config, envOverrides);
  const reqStatus = statuses.find((s) => s.id === requested);
  if (reqStatus?.configured) {
    return { provider: requested, switchedReason: null };
  }

  if (config.ai.fallbackToBuiltIn) {
    const fallbackOrder: AiProviderId[] = ['builtin', 'groq', 'claude', 'openai'];
    const ready = fallbackOrder.find((id) => statuses.some((s) => s.id === id && s.configured));
    if (ready) {
      const readyName = statuses.find((s) => s.id === ready)?.name ?? ready;
      return {
        provider: ready,
        switchedReason: `${reqStatus?.name ?? requested} 키 미등록 → config.json 폴백 규칙에 따라 ${readyName}으로 자동 전환`,
      };
    }
  }

  return { provider: requested, switchedReason: null };
}

/**
 * Autonomous multi-step AI Agent pipeline:
 * 1) Target email discovery & automatic date/scope resolution
 * 2) Privacy & sensitive-data audit
 * 3) Engine selection & multi-provider LLM work analysis
 * 4) Work log + blog synthesis & completeness verification
 * 5) Archive & export preparation
 */
export async function runAutonomousMailAgent(req: AutonomousAgentRequest): Promise<AnalysisResult> {
  const steps: AgentExecutionStep[] = [
    {
      id: 'scope_resolution',
      title: '1. 대상 메일 자동 선별 및 스레드 분류',
      status: 'pending',
      detail: '지정 일자·기간 및 수신/발신/선택 범위에 맞는 실제 메일을 선별합니다.',
    },
    {
      id: 'privacy_audit',
      title: '2. 민감 정보 자동 감사 및 비식별화',
      status: 'pending',
      detail: '이메일·연락처·계약 금액 노출 여부를 검사합니다.',
    },
    {
      id: 'ai_analysis',
      title: '3. AI 엔진 업무 분석 · 일지 · 블로그 자동 작성',
      status: 'pending',
      detail: 'config.json 모델 정책과 연동된 AI 엔진으로 문서를 일괄 생성합니다.',
    },
    {
      id: 'quality_check',
      title: '4. 결과 품질 검증 및 내보내기 준비',
      status: 'pending',
      detail: '업무 일지 항목과 블로그 섹션 완성도를 자동 확인합니다.',
    },
  ];

  const emit = (index: number, patch: Partial<AgentExecutionStep>) => {
    steps[index] = {
      ...steps[index],
      ...patch,
      timestamp: new Date().toISOString(),
    };
    req.onStepUpdate?.(steps.map((s) => ({ ...s })));
  };

  // Step 1: Scope resolution (with intelligent auto-date fallback if requested)
  emit(0, { status: 'running', detail: '연동된 실제 메일함에서 분석 대상을 선별하는 중...' });

  if (!req.emails || req.emails.length === 0) {
    emit(0, { status: 'error', detail: '연동된 실제 메일이 없습니다. PST/EML 파일을 열거나 Outlook 계정을 연결해 주세요.' });
    throw new Error('연동된 실제 메일이 없습니다. 먼저 Outlook .pst/.eml 파일을 연동하거나 Microsoft 365 메일함을 연결해 주세요.');
  }

  let effectiveOptions = { ...req.options };
  let matched = filterEmails(req.emails, effectiveOptions);
  let autoAdjustedNote = '';

  if (matched.length === 0 && req.autoPickLatestDateIfEmpty !== false && effectiveOptions.scope !== 'selected') {
    const dates = getAvailableDates(req.emails);
    if (dates.length > 0) {
      effectiveOptions = {
        ...effectiveOptions,
        dateMode: 'single',
        singleDate: dates[0].date,
        scope: 'all',
      };
      matched = filterEmails(req.emails, effectiveOptions);
      autoAdjustedNote = ` (지정 조건 메일 없음 → 최근 메일 일자 ${dates[0].date} 자동 선택)`;
    }
  }

  if (matched.length === 0) {
    const msg =
      effectiveOptions.scope === 'selected'
        ? '선택한 메일이 없습니다. 메일 목록에서 분석할 메일을 체크해 주세요.'
        : '조건에 일치하는 실제 메일이 없습니다.';
    emit(0, { status: 'error', detail: msg });
    throw new Error(msg);
  }

  const maxEmails = req.config.ai.maxEmailsPerRun || 40;
  const trimmed = matched.slice(0, maxEmails);
  const receivedCount = trimmed.filter((m) => m.folder === 'inbox').length;
  const sentCount = trimmed.filter((m) => m.folder === 'sent').length;

  let periodLabel: string;
  if (effectiveOptions.scope === 'selected') {
    const sortedDates = trimmed.map((m) => m.date).sort();
    periodLabel =
      sortedDates[0] === sortedDates[sortedDates.length - 1]
        ? sortedDates[0]
        : `${sortedDates[0]} ~ ${sortedDates[sortedDates.length - 1]}`;
  } else {
    periodLabel =
      effectiveOptions.dateMode === 'single'
        ? effectiveOptions.singleDate
        : `${effectiveOptions.startDate} ~ ${effectiveOptions.endDate}`;
  }
  const scopeLabel = SCOPE_LABELS[effectiveOptions.scope];

  emit(0, {
    status: 'completed',
    detail: `총 ${trimmed.length}건 선별 완료 (수신 ${receivedCount}건 · 발신 ${sentCount}건 · ${periodLabel})${autoAdjustedNote}`,
  });

  // Step 2: Privacy audit
  emit(1, { status: 'running', detail: '민감 정보(이메일·전화번호·금액) 패턴을 스캔하는 중...' });
  const audit = auditAndMaskEmails(trimmed, effectiveOptions.maskSensitiveInfo);
  emit(1, {
    status: 'completed',
    detail: effectiveOptions.maskSensitiveInfo
      ? `민감 정보 보호 활성화 — 본문 내 ${audit.redactedCount}건의 개인/금액 패턴 자동 비식별화`
      : '민감 정보 마스킹 비활성화 (원문 기준 정밀 분석 모드)',
  });

  // Step 3: Provider resolution & AI synthesis
  const { provider, switchedReason } = resolveAgentProvider(
    effectiveOptions.provider,
    req.config,
    req.envOverrides
  );
  emit(2, {
    status: 'running',
    detail: switchedReason
      ? `${switchedReason} · 업무 분석 및 문서 생성 중...`
      : `${provider.toUpperCase()} 엔진으로 수행 업무 분석, 업무 일지, 블로그 초안을 생성하는 중...`,
  });

  const combinedFocus = [
    req.config.agent.customSystemDirective,
    req.agentGoal,
    effectiveOptions.customFocus,
  ]
    .map((s) => (s || '').trim())
    .filter(Boolean)
    .join(' / ');

  let result: AnalysisResult;
  let usedOfflineFallback = false;
  try {
    result = await analyzeEmailsWithProvider({
      emails: trimmed,
      provider,
      periodLabel,
      scopeLabel,
      blogTone: effectiveOptions.blogTone,
      maskSensitiveInfo: effectiveOptions.maskSensitiveInfo,
      includeActionItems: effectiveOptions.includeActionItems,
      customFocus: combinedFocus,
      modelOverrides: req.config.ai.models,
      envOverrides: req.envOverrides,
    });
  } catch (err) {
    if (req.config.ai.fallbackToBuiltIn !== false) {
      const offlineEmails = sanitizeEmailsForOffline(trimmed, effectiveOptions.maskSensitiveInfo);
      result = buildDeterministicWorkLog({
        emails: offlineEmails,
        projects: req.projects ?? [],
        periodLabel,
        templateName: 'standard',
      });
      usedOfflineFallback = true;
    } else {
      const errMsg = err instanceof Error ? err.message : 'AI 분석 실패';
      emit(2, { status: 'error', detail: errMsg });
      throw err;
    }
  }

  emit(2, {
    status: 'completed',
    detail: usedOfflineFallback
      ? `외부 AI 키 미등록/호출 불가 → 로컬 결정론적 오프라인 분석 완료 — 수행 업무 ${result.workPerformed.length}건, 주요 수신 ${result.incomingHighlights.length}건 추출`
      : `${result.providerModel} 분석 완료 — 수행 업무 ${result.workPerformed.length}건, 주요 수신 ${result.incomingHighlights.length}건 추출`,
  });

  // Step 4: Quality verification & export readiness
  emit(3, { status: 'running', detail: '업무 일지 항목과 블로그 초안 구조를 검증하는 중...' });
  const totalTasks =
    result.workLog.completedTasks.length + result.workLog.inProgressTasks.length;
  const blogSections = result.blogPost.sections.length;
  emit(3, {
    status: 'completed',
    detail: `업무 일지(${totalTasks}개 항목) 및 블로그 초안(${blogSections}개 섹션, 약 ${result.blogPost.readingTimeMinutes}분 분량) 생성·검증 완료`,
  });

  return {
    ...result,
    agentSteps: steps,
    redactedCount: audit.redactedCount,
  };
}
