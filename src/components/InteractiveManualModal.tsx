import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
  BookOpen,
  CheckCircle2,
  ChevronRight,
  ExternalLink,
  FolderOpen,
  History,
  Inbox,
  Layers,
  MessageSquare,
  MousePointer,
  Paperclip,
  Pause,
  Play,
  RefreshCw,
  Repeat,
  Search,
  Send,
  Settings,
  SkipBack,
  SkipForward,
  Sparkles,
  Star,
  X,
} from 'lucide-react';
import {
  MANUAL_CHANGELOG,
  MANUAL_ICONS,
  MANUAL_MOTION_SCENES,
  MANUAL_UPDATED_AT,
  MANUAL_VERSION,
  MANUAL_ZONES,
  type ManualIconEntry,
  type ManualPreviewStateId,
  type ManualZoneId,
} from '../data/manualCatalog';

export interface InteractiveManualModalProps {
  isOpen: boolean;
  onClose: () => void;
  standalonePopup?: boolean;
}

export const InteractiveManualModal: React.FC<InteractiveManualModalProps> = ({
  isOpen,
  onClose,
  standalonePopup = false,
}) => {
  const [mode, setMode] = useState<'motion' | 'explorer' | 'changelog'>('motion');
  const [sceneIdx, setSceneIdx] = useState(0);
  const [isPlaying, setIsPlaying] = useState(true);
  const [speed, setSpeed] = useState<0.75 | 1 | 1.5>(1);
  const [loop, setLoop] = useState(true);
  const [elapsedInScene, setElapsedInScene] = useState(0);
  const [selectedZone, setSelectedZone] = useState<'ALL' | ManualZoneId>('ALL');
  const [activeIconId, setActiveIconId] = useState<string>(MANUAL_ICONS[0].id);
  const [clickPulse, setClickPulse] = useState(false);
  const [highlightedRange, setHighlightedRange] = useState<'1-10' | '11-20' | '21-end' | null>(null);
  const iconItemRefs = useRef<Record<string, HTMLButtonElement | null>>({});

  const currentScene = MANUAL_MOTION_SCENES[sceneIdx] || MANUAL_MOTION_SCENES[0];
  const totalDurationSec = useMemo(
    () => MANUAL_MOTION_SCENES.reduce((acc, s) => acc + s.durationSec, 0),
    []
  );

  const cumulativeBeforeScene = useMemo(
    () => MANUAL_MOTION_SCENES.slice(0, sceneIdx).reduce((acc, s) => acc + s.durationSec, 0),
    [sceneIdx]
  );

  const globalElapsedSec = Math.min(totalDurationSec, cumulativeBeforeScene + elapsedInScene);

  useEffect(() => {
    if (!isOpen || mode !== 'motion' || !isPlaying) return;
    const tickMs = 100;
    const timer = window.setInterval(() => {
      setElapsedInScene((prev) => {
        const next = prev + (tickMs / 1000) * speed;
        if (next >= currentScene.durationSec) {
          if (sceneIdx + 1 < MANUAL_MOTION_SCENES.length) {
            setSceneIdx((s) => s + 1);
          } else if (loop) {
            setSceneIdx(0);
          } else {
            setIsPlaying(false);
          }
          return 0;
        }
        return next;
      });
    }, tickMs);
    return () => window.clearInterval(timer);
  }, [isOpen, mode, isPlaying, speed, loop, sceneIdx, currentScene.durationSec]);

  useEffect(() => {
    if (mode === 'motion') {
      setActiveIconId(currentScene.targetIconId);
      setClickPulse(false);
      const t = window.setTimeout(() => setClickPulse(true), 600);
      return () => window.clearTimeout(t);
    }
  }, [mode, sceneIdx, currentScene.targetIconId]);

  const activeIcon: ManualIconEntry = useMemo(
    () => MANUAL_ICONS.find((i) => i.id === activeIconId) || MANUAL_ICONS[0],
    [activeIconId]
  );

  const filteredIcons = useMemo(() => {
    if (selectedZone === 'ALL') return MANUAL_ICONS;
    return MANUAL_ICONS.filter((i) => i.zone === selectedZone);
  }, [selectedZone]);

  useEffect(() => {
    setHighlightedRange(null);
  }, [selectedZone]);

  const iconJumpRanges = [
    { key: '1-10' as const, label: '1~10', start: 1, end: 10 },
    { key: '11-20' as const, label: '11~20', start: 11, end: 20 },
    { key: '21-end' as const, label: '21~끝', start: 21, end: Number.POSITIVE_INFINITY },
  ];

  const handleJumpToRange = (range: (typeof iconJumpRanges)[number]) => {
    const firstInRange = filteredIcons.find(
      (ic) => ic.number >= range.start && ic.number <= range.end
    );
    if (!firstInRange) return;
    setHighlightedRange(range.key);
    const el = iconItemRefs.current[firstInRange.id];
    if (el) {
      el.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  };

  if (!isOpen) return null;

  const previewState: ManualPreviewStateId =
    mode === 'motion'
      ? currentScene.previewState
      : activeIcon.zone === 'dashboard'
        ? 'executive-dashboard'
        : activeIcon.zone === 'sgw-chat'
          ? 'sgw-login-chat'
          : activeIcon.zone === 'attach-dock' && activeIcon.id.includes('batch')
            ? 'batch-folder-export'
            : activeIcon.zone === 'attach-dock'
              ? 'attach-preview'
              : activeIcon.zone === 'ai-dock'
                ? 'ai-dock-reply'
                : 'outlook-overview';

  const handleOpenStandaloneHtmlPopup = () => {
    window.open(
      '/?manual=1',
      'bmail_interactive_manual',
      'width=1420,height=880,menubar=no,toolbar=no,location=no,status=no,resizable=yes,scrollbars=yes'
    );
  };

  const formatTime = (sec: number) => {
    const s = Math.max(0, Math.floor(sec));
    const mm = String(Math.floor(s / 60)).padStart(2, '0');
    const ss = String(s % 60).padStart(2, '0');
    return `${mm}:${ss}`;
  };

  const cursorPos =
    mode === 'motion'
      ? currentScene.cursor
      : { x: activeIcon.xPercent, y: activeIcon.yPercent };

  const wrapperCls = standalonePopup
    ? 'w-full min-h-[100dvh] bg-[#0B0F17] text-white flex flex-col overflow-hidden'
    : 'fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-2 sm:p-4';

  const dialogCls = standalonePopup
    ? 'flex-1 flex flex-col min-h-0 overflow-hidden'
    : 'w-full max-w-[1440px] h-[92dvh] bg-[#0F172A] text-white rounded-xl border border-slate-700/80 shadow-2xl flex flex-col overflow-hidden';

  return (
    <div className={wrapperCls} role="dialog" aria-modal="true" aria-label="Bmail 인터랙티브 HTML 매뉴얼">
      <div className={dialogCls}>
        {/* Top Modal Header */}
        <header className="px-4 py-2.5 bg-[#0F6CBD] border-b border-white/15 flex flex-wrap items-center justify-between gap-2 shrink-0">
          <div className="flex items-center gap-2.5 min-w-0">
            <span className="w-7 h-7 rounded-md bg-white/15 flex items-center justify-center shrink-0">
              <BookOpen className="w-4 h-4 text-white" />
            </span>
            <div className="min-w-0">
              <div className="flex items-center gap-2">
                <h1 className="text-sm font-bold tracking-tight text-white truncate">
                  Bmail Outlook AI 인터랙티브 HTML 매뉴얼 · 모션그래픽 가이드
                </h1>
                <span className="px-2 py-0.5 text-[10px] font-bold bg-emerald-400/20 text-emerald-100 border border-emerald-300/40 rounded-full">
                  v{MANUAL_VERSION} 자동동기화 ({MANUAL_UPDATED_AT})
                </span>
              </div>
              <p className="text-[11px] text-white/80 truncate">
                실제 화면 1:1 레이아웃 위에서 각 아이콘 기능을 동영상 모션그래픽과 인터랙티브 핀으로 즉시 확인합니다.
              </p>
            </div>
          </div>

          <div className="flex items-center gap-1.5 shrink-0">
            <div className="inline-flex bg-black/25 p-0.5 rounded-lg border border-white/15">
              <button
                type="button"
                onClick={() => {
                  setMode('motion');
                  setIsPlaying(true);
                }}
                className={`px-2.5 py-1 rounded-md text-xs font-semibold flex items-center gap-1.5 transition-colors ${
                  mode === 'motion' ? 'bg-white text-[#0F6CBD]' : 'text-white/85 hover:bg-white/10'
                }`}
              >
                <Play className="w-3.5 h-3.5" />
                모션그래픽 가이드 (동영상)
              </button>
              <button
                type="button"
                onClick={() => {
                  setMode('explorer');
                  setIsPlaying(false);
                }}
                className={`px-2.5 py-1 rounded-md text-xs font-semibold flex items-center gap-1.5 transition-colors ${
                  mode === 'explorer' ? 'bg-white text-[#0F6CBD]' : 'text-white/85 hover:bg-white/10'
                }`}
              >
                <MousePointer className="w-3.5 h-3.5" />
                화면 아이콘 탐색 ({MANUAL_ICONS.length})
              </button>
              <button
                type="button"
                onClick={() => {
                  setMode('changelog');
                  setIsPlaying(false);
                }}
                className={`px-2.5 py-1 rounded-md text-xs font-semibold flex items-center gap-1.5 transition-colors ${
                  mode === 'changelog' ? 'bg-white text-[#0F6CBD]' : 'text-white/85 hover:bg-white/10'
                }`}
              >
                <History className="w-3.5 h-3.5" />
                기능 업데이트 이력
              </button>
            </div>

            {!standalonePopup && (
              <button
                type="button"
                onClick={handleOpenStandaloneHtmlPopup}
                title="독립된 새 HTML 브라우저 팝업창으로 매뉴얼 띄우기"
                className="px-2.5 py-1.5 rounded-md bg-white/15 hover:bg-white/25 border border-white/25 text-white text-xs font-semibold inline-flex items-center gap-1"
              >
                <ExternalLink className="w-3.5 h-3.5" />
                <span className="hidden md:inline">새 HTML 팝업창</span>
              </button>
            )}

            <button
              type="button"
              onClick={onClose}
              aria-label="매뉴얼 닫기"
              className="w-8 h-8 rounded-md bg-black/20 hover:bg-rose-600 text-white flex items-center justify-center transition-colors"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        </header>

        {/* Body */}
        {mode === 'changelog' ? (
          <div className="flex-1 overflow-y-auto p-6 bg-slate-950 space-y-5">
            <div className="max-w-4xl mx-auto space-y-4">
              <div className="p-4 rounded-xl bg-sky-950/50 border border-sky-500/30 flex items-center justify-between">
                <div>
                  <div className="text-xs font-bold text-sky-300">단일 진실 원천(Single-Source) 매뉴얼 자동 연동</div>
                  <p className="text-xs text-slate-300 mt-0.5">
                    Bmail 기능이 업데이트될 때마다 <code className="text-sky-200">src/data/manualCatalog.ts</code>와 버전 검증 테스트가 함께 갱신되어 매뉴얼 화면과 아이콘 설명이 항상 최신 상태로 유지됩니다.
                  </p>
                </div>
                <span className="px-3 py-1 rounded-full bg-sky-500/20 text-sky-200 text-xs font-bold">
                  현재 v{MANUAL_VERSION}
                </span>
              </div>

              {MANUAL_CHANGELOG.map((entry) => (
                <div
                  key={entry.version}
                  className="p-5 rounded-xl bg-slate-900/90 border border-slate-800 space-y-3"
                >
                  <div className="flex items-center justify-between border-b border-slate-800 pb-2.5">
                    <div className="flex items-center gap-2">
                      <span className="px-2.5 py-0.5 rounded bg-[#0F6CBD] text-white text-xs font-bold">
                        v{entry.version}
                      </span>
                      <h2 className="text-sm font-bold text-white">{entry.headline}</h2>
                    </div>
                    <span className="text-xs text-slate-400 font-mono">{entry.date}</span>
                  </div>
                  <ul className="space-y-2 text-xs text-slate-200">
                    {entry.highlights.map((h, i) => (
                      <li key={i} className="flex items-start gap-2 leading-relaxed">
                        <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0 mt-0.5" />
                        <span>{h}</span>
                      </li>
                    ))}
                  </ul>
                </div>
              ))}
            </div>
          </div>
        ) : (
          <div className="flex-1 flex flex-col lg:flex-row min-h-0 overflow-hidden">
            {/* Left / Center: 1:1 Live HTML Visual Screen Replica + Motion Graphic Stage */}
            <div className="flex-1 flex flex-col min-w-0 min-h-0 bg-[#070B12] p-3 overflow-hidden">
              {/* Zone Filter Pills (when in explorer mode) or Scene Step Strip (when in motion mode) */}
              {mode === 'explorer' ? (
                <div className="flex items-center gap-1.5 overflow-x-auto pb-2 shrink-0">
                  <button
                    type="button"
                    onClick={() => setSelectedZone('ALL')}
                    className={`px-2.5 py-1 rounded-full text-[11px] font-semibold whitespace-nowrap ${
                      selectedZone === 'ALL'
                        ? 'bg-[#0F6CBD] text-white'
                        : 'bg-slate-800 text-slate-300 hover:bg-slate-700'
                    }`}
                  >
                    전체 화면 아이콘 ({MANUAL_ICONS.length})
                  </button>
                  {MANUAL_ZONES.map((z) => (
                    <button
                      key={z.id}
                      type="button"
                      onClick={() => setSelectedZone(z.id)}
                      className={`px-2.5 py-1 rounded-full text-[11px] font-semibold whitespace-nowrap ${
                        selectedZone === z.id
                          ? 'bg-sky-500 text-white'
                          : 'bg-slate-800/90 text-slate-300 hover:bg-slate-700'
                      }`}
                    >
                      {z.label}
                    </button>
                  ))}
                </div>
              ) : (
                <div className="flex items-center gap-1.5 overflow-x-auto pb-2 shrink-0">
                  {MANUAL_MOTION_SCENES.map((sc, idx) => (
                    <button
                      key={sc.id}
                      type="button"
                      onClick={() => {
                        setSceneIdx(idx);
                        setElapsedInScene(0);
                      }}
                      className={`px-2.5 py-1 rounded-md text-[11px] font-semibold whitespace-nowrap flex items-center gap-1.5 transition-all ${
                        idx === sceneIdx
                          ? 'bg-[#0F6CBD] text-white ring-1 ring-sky-300 shadow'
                          : 'bg-slate-900 text-slate-400 hover:bg-slate-800 hover:text-slate-200'
                      }`}
                    >
                      <span className="w-4 h-4 rounded-full bg-black/30 text-[10px] flex items-center justify-center font-bold">
                        {sc.step}
                      </span>
                      <span>{sc.title.slice(0, 18)}</span>
                    </button>
                  ))}
                </div>
              )}

              {/* 1:1 Interactive HTML Screen Replica Stage */}
              <div className="relative flex-1 min-h-0 rounded-lg border border-slate-700/90 bg-[#F3F2F1] text-[#242424] shadow-inner overflow-hidden select-none flex flex-col">
                {/* Replica 1: Top Blue Outlook Bar */}
                <div className="h-9 bg-[#0F6CBD] text-white px-3 flex items-center justify-between text-[11px] shrink-0">
                  <div className="flex items-center gap-2">
                    <span className="font-bold tracking-tight">::: Outlook Bmail</span>
                    <span className="px-1.5 py-0.5 rounded bg-white/15 text-[10px]">
                      SQLite 영구보관 24,626건
                    </span>
                  </div>
                  <div className="w-72 bg-white text-slate-700 rounded px-2 py-0.5 text-[10px] flex items-center gap-1">
                    <Search className="w-3 h-3 text-[#0F6CBD]" />
                    <span>검색 (제목, 보낸 사람, 22K FD, 도면·본문)</span>
                  </div>
                  <div className="flex items-center gap-2">
                    <span className="px-2 py-0.5 rounded bg-white/20 font-semibold text-[10px]">
                      매뉴얼
                    </span>
                    <span>POP3</span>
                    <span>SGW채팅</span>
                    <span>일정</span>
                    <Settings className="w-3.5 h-3.5" />
                  </div>
                </div>

                {/* Replica 2: Classic Outlook Grouped Icon Ribbon */}
                <div className="h-10 bg-white border-b border-[#E0E0E0] px-3 flex items-center justify-between text-[11px] shrink-0">
                  <div className="flex items-center gap-2">
                    <span className="px-2 py-1 rounded bg-[#0F6CBD] text-white font-semibold text-[10px]">
                      + 새 메일
                    </span>
                    <span className="px-1.5 py-0.5 rounded hover:bg-slate-100">회신</span>
                    <span className="px-1.5 py-0.5 rounded hover:bg-slate-100">전체 회신</span>
                    <span className="px-1.5 py-0.5 rounded hover:bg-slate-100">전달</span>
                    <span className="h-4 w-px bg-slate-200" />
                    <span
                      className={`px-2 py-0.5 rounded font-semibold flex items-center gap-1 ${
                        previewState === 'live-sync'
                          ? 'bg-emerald-600 text-white animate-pulse'
                          : 'bg-[#EBF3FC] text-[#0F6CBD]'
                      }`}
                    >
                      <RefreshCw className="w-3 h-3" />
                      신규 메일 증분 동기화
                    </span>
                    <span className="px-2 py-0.5 rounded bg-purple-50 text-purple-700 border border-purple-200 font-semibold">
                      우측 2분할 (AI · 첨부)
                    </span>
                    <span className="px-2 py-0.5 rounded bg-rose-50 text-rose-700 border border-rose-200 font-semibold">
                      주간·월간 분석글 대시보드
                    </span>
                  </div>
                  <span className="text-[10px] text-slate-500">
                    연결된 데이터 파일: 김봉정.pst · 백업.pst (실시간 COM 연동)
                  </span>
                </div>

                {/* Replica Main Body: Switch between Executive Dashboard View vs 4-Column Classic Outlook Workbench */}
                {previewState === 'executive-dashboard' ? (
                  <div className="flex-1 bg-[#FAF9F8] p-3 flex flex-col gap-2.5 overflow-hidden text-[11px]">
                    <div className="bg-white border border-slate-200 rounded-lg p-2.5 flex items-center justify-between">
                      <div className="flex items-center gap-2">
                        <span className="px-2 py-0.5 rounded bg-[#0F6CBD] text-white font-bold">
                          주간 · 월간 종합 분석 대시보드
                        </span>
                        <span className="text-slate-600">
                          기간: 2026-11 (월간) · 프로젝트: 전체 · 주제: 전체 (영구저장 24,626건)
                        </span>
                      </div>
                      <div className="flex items-center gap-1.5">
                        <span className="px-2 py-0.5 rounded bg-emerald-600 text-white font-semibold">
                          신규 메일만 증분 업데이트 완료 (+4건)
                        </span>
                        <span className="px-2 py-0.5 rounded bg-purple-600 text-white font-semibold">
                          Claude 5.5 분석글 생성
                        </span>
                      </div>
                    </div>

                    <div className="grid grid-cols-5 gap-2">
                      <div className="bg-white p-2 rounded border border-slate-200">
                        <div className="text-[10px] text-slate-500">총 보관 메일</div>
                        <div className="text-sm font-bold text-slate-900">24,626건</div>
                      </div>
                      <div className="bg-white p-2 rounded border border-slate-200">
                        <div className="text-[10px] text-slate-500">받은 메일 / 보낸 메일</div>
                        <div className="text-sm font-bold text-[#0F6CBD]">19,410 / 5,216</div>
                      </div>
                      <div className="bg-rose-50 p-2 rounded border border-rose-200">
                        <div className="text-[10px] text-rose-700 font-semibold">내게 왔는데 미회신건</div>
                        <div className="text-sm font-bold text-rose-600">18건 (긴급 5건)</div>
                      </div>
                      <div className="bg-white p-2 rounded border border-slate-200">
                        <div className="text-[10px] text-slate-500">실첨부 문서 (도면·POS)</div>
                        <div className="text-sm font-bold text-emerald-700">6,842개</div>
                      </div>
                      <div className="bg-white p-2 rounded border border-slate-200">
                        <div className="text-[10px] text-slate-500">활성 프로젝트 · 주제</div>
                        <div className="text-sm font-bold text-purple-700">12개 호선 · 7대 주제</div>
                      </div>
                    </div>

                    <div className="grid grid-cols-12 gap-2 flex-1 min-h-0">
                      <div className="col-span-5 bg-white border border-slate-200 rounded-lg p-2.5 flex flex-col justify-between">
                        <div>
                          <div className="font-bold text-[#0F6CBD] mb-1">
                            1. 서술형 종합 분석글 (주간·월간 브리핑)
                          </div>
                          <p className="text-[10px] text-slate-700 leading-relaxed">
                            금주(2026-W47) 총 312건(수신 248건, 발신 64건, 실첨부 94개)이 집계되었습니다. 핵심 비중 1위 프로젝트는 <strong>22K-FD (KYTS-1037)</strong>이며 수정 POS 발행 및 한국선급 승인 미결 도면 조율이 집중되어 있습니다.
                          </p>
                          <p className="text-[10px] text-rose-700 mt-1.5 font-medium">
                            • 내게 온 미회신 최우선 조치: [KYTS-1037] 22K TLC 수정 POS 발행 요청 (최길용, D+3) 외 17건
                          </p>
                        </div>
                        <div className="text-[10px] text-slate-500 border-t pt-1">
                          스냅샷 자동 저장됨 · Markdown / Excel CSV / HTML 즉시 출력
                        </div>
                      </div>

                      <div className="col-span-7 bg-white border border-slate-200 rounded-lg p-2.5 flex flex-col gap-1.5 overflow-hidden">
                        <div className="font-bold text-slate-800">
                          2. 주간/월간 집계 테이블 & 프로젝트·주제별 미회신 현황
                        </div>
                        <table className="w-full text-[10px] border-collapse">
                          <thead>
                            <tr className="bg-slate-100 text-slate-700">
                              <th className="p-1 text-left">구분 (주차/프로젝트/주제)</th>
                              <th className="p-1 text-right">받은 메일</th>
                              <th className="p-1 text-right">보낸 메일</th>
                              <th className="p-1 text-right text-rose-600">내 미회신</th>
                              <th className="p-1 text-right">첨부파일</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y">
                            <tr>
                              <td className="p-1 font-semibold">2026-W47 (11/16~11/23) · 22K FD</td>
                              <td className="p-1 text-right">142건</td>
                              <td className="p-1 text-right">38건</td>
                              <td className="p-1 text-right font-bold text-rose-600">7건</td>
                              <td className="p-1 text-right">54개</td>
                            </tr>
                            <tr>
                              <td className="p-1 font-semibold">H042 Exmar · 3D 모델링·생산설계</td>
                              <td className="p-1 text-right">86건</td>
                              <td className="p-1 text-right">29건</td>
                              <td className="p-1 text-right font-bold text-rose-600">4건</td>
                              <td className="p-1 text-right">31개</td>
                            </tr>
                            <tr>
                              <td className="p-1 font-semibold">선급·검사·미결 (KR/ABS 승인도)</td>
                              <td className="p-1 text-right">64건</td>
                              <td className="p-1 text-right">19건</td>
                              <td className="p-1 text-right font-bold text-rose-600">5건</td>
                              <td className="p-1 text-right">28개</td>
                            </tr>
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                ) : (
                  <div className="flex-1 flex min-h-0 overflow-hidden">
                    {/* Col 1: Classic Outlook Favorites & Store Tree (15%) */}
                    <div className="w-[15%] bg-[#F0F0F0] border-r border-[#E0E0E0] p-2 text-[10px] flex flex-col justify-between shrink-0">
                      <div className="space-y-2">
                        <div>
                          <div className="font-bold text-slate-700 mb-1">⌄ 즐겨찾기</div>
                          <div className="space-y-0.5 pl-1.5">
                            <div className="flex justify-between bg-[#CFE4FA] text-[#0F6CBD] font-bold px-1.5 py-0.5 rounded">
                              <span>받은 편지함</span>
                              <span>3</span>
                            </div>
                            <div className="flex justify-between px-1.5 py-0.5">
                              <span>유일</span>
                            </div>
                            <div className="flex justify-between px-1.5 py-0.5">
                              <span>선보</span>
                            </div>
                            <div className="flex justify-between px-1.5 py-0.5">
                              <span>가스앤텍</span>
                            </div>
                          </div>
                        </div>

                        <div>
                          <div className="font-bold text-slate-800 mb-1">⌄ 김봉정</div>
                          <div className="space-y-0.5 pl-1.5">
                            <div className="flex justify-between font-bold">
                              <span>⌄ 받은 편지함</span>
                              <span className="text-[#0F6CBD]">3</span>
                            </div>
                            <div className="pl-2 space-y-0.5 text-slate-700">
                              <div>선보</div>
                              <div>가스앤텍</div>
                              <div className="font-semibold text-[#0F6CBD]">⌄ 22K FD</div>
                              <div className="pl-2 text-slate-600">22K POS</div>
                              <div className="pl-2 text-slate-600">선급승도도 제출</div>
                              <div>› 한라ims</div>
                              <div>› DOCK</div>
                            </div>
                            <div className="flex justify-between">
                              <span>임시 보관함</span>
                              <span className="text-[#0F6CBD]">[23]</span>
                            </div>
                            <div className="flex justify-between">
                              <span>› 지운 편지함</span>
                              <span className="text-[#0F6CBD] font-bold">390</span>
                            </div>
                          </div>
                        </div>

                        <div>
                          <div className="font-bold text-slate-800">⌄ 백업</div>
                          <div className="pl-1.5 text-slate-600 space-y-0.5">
                            <div>› 1.35K FD</div>
                            <div>› 1120 TEU CONT</div>
                            <div>› 한진 PMU</div>
                          </div>
                        </div>
                      </div>
                      <div className="text-[9px] text-slate-500 border-t border-slate-300 pt-1">
                        항목: 815 · 읽지 않음: 3
                      </div>
                    </div>

                    {/* Col 2: Classic Today/Last Week Mail List (16%) */}
                    <div className="w-[16%] bg-white border-r border-[#E0E0E0] flex flex-col text-[10px] shrink-0">
                      <div className="px-2 py-1 bg-slate-50 border-b font-bold text-slate-700 flex justify-between">
                        <span>모든 항목 · 읽지 않음</span>
                        <span>날짜 ↓</span>
                      </div>
                      <div className="p-1.5 space-y-1 overflow-hidden">
                        <div className="text-[9px] font-bold text-slate-500">⌄ 오늘</div>
                        <div className="p-1.5 rounded bg-[#EBF3FC] border-l-[3px] border-[#0F6CBD]">
                          <div className="flex justify-between font-bold">
                            <span>박봉근</span>
                            <span className="text-[#0F6CBD]">오전 11:26</span>
                          </div>
                          <div className="font-bold text-[#0F6CBD] truncate">
                            [FW][한국선급 승인 미결] KYTS1037
                          </div>
                          <div className="text-slate-600 truncate">내일(11/24)까지 회신 부탁...</div>
                        </div>

                        <div className="p-1.5 rounded hover:bg-slate-50 border-l-[3px] border-[#0F6CBD]">
                          <div className="flex justify-between font-bold">
                            <span>MNC최길용</span>
                            <span>14:28</span>
                          </div>
                          <div className="font-semibold truncate">
                            KYTS-1037, 22K TLC 수정POS
                          </div>
                        </div>

                        <div className="text-[9px] font-bold text-slate-500 pt-1">⌄ 지난 주</div>
                        <div className="p-1.5 rounded hover:bg-slate-50">
                          <div className="flex justify-between">
                            <span className="font-semibold">김정인</span>
                            <span>11-20 (금)</span>
                          </div>
                          <div className="truncate">[RE]H042 Exmar 생산설계 자료</div>
                        </div>
                      </div>
                    </div>

                    {/* Col 3: Reading Pane with Rich HTML + Inline cid: 3D Image (33%) */}
                    <div className="w-[33%] bg-white border-r border-[#E0E0E0] p-3 flex flex-col justify-between text-[10px] overflow-hidden">
                      <div className="space-y-2">
                        <div className="text-xs font-bold text-slate-900">
                          RE: [RE]H042 Exmar Project_생산설계 관련 자료 제공의 건_261102
                        </div>
                        <div className="flex items-center justify-between text-slate-600 border-b pb-1.5">
                          <span>김정인 &lt;jikim@seastargo.com&gt; → 석주선 대표님, 김봉정 부장</span>
                          <span>2026-11-20 14:28</span>
                        </div>

                        {/* Real attachments only (signature image001.png hidden!) */}
                        <div className="flex flex-wrap gap-1">
                          <span className="px-2 py-0.5 rounded bg-slate-100 border border-slate-300 font-medium flex items-center gap-1">
                            <Paperclip className="w-3 h-3 text-rose-600" />
                            22K_TLC_FLOATING_DOCK_2604.pdf
                          </span>
                          <span className="px-2 py-0.5 rounded bg-slate-100 border border-slate-300 font-medium flex items-center gap-1">
                            <Paperclip className="w-3 h-3 text-emerald-600" />
                            KYTS-1037_POS_Deviation.xlsx
                          </span>
                        </div>

                        {/* Rich HTML Body + Inline cid: 3D CAD rendering preview */}
                        <div className="p-2 rounded bg-slate-50 border border-slate-200 space-y-1.5">
                          <p className="text-slate-800">
                            수신 : 김봉정 부장 / SEASTAR · 아래 Z 좌표는 1ST DECK CAMBER 상부 기준이니 참고 바랍니다.
                          </p>
                          <div className="bg-yellow-200/80 px-1.5 py-0.5 font-semibold text-slate-900 rounded">
                            아래 사진을 첨부하오니 어제 사진과 비교해 모듈 상단 인양 러그 간섭 확인 바랍니다.
                          </div>
                          <div className="h-28 rounded bg-[#141824] border border-slate-700 p-2 flex items-center justify-between text-white">
                            <div className="space-y-1">
                              <span className="px-1.5 py-0.5 rounded bg-sky-500/30 text-sky-200 text-[9px] font-bold">
                                cid:inline-3d-cad@outlook (본문 인라인 렌더링)
                              </span>
                              <div className="text-[10px] font-semibold">
                                H042 Exmar Recondenser Module SKID 3D Modeling
                              </div>
                            </div>
                            <div className="w-24 h-20 rounded bg-gradient-to-br from-fuchsia-600/40 via-purple-600/40 to-emerald-500/40 border border-pink-400/50 flex items-center justify-center text-[9px] font-mono">
                              [3D CAD SKID]
                            </div>
                          </div>
                        </div>
                      </div>

                      <div className="text-[9px] text-emerald-700 bg-emerald-50 border border-emerald-200 px-2 py-1 rounded">
                        ✓ 원본 HTML 표·하이라이트·cid 인라인 이미지 복원 완료 (서명 로고 6개 자동 분리)
                      </div>
                    </div>

                    {/* Col 4 (Split 1): AI Analysis & Reply Dock (18%) */}
                    <div className="w-[18%] bg-[#FAF9F8] border-r border-[#E0E0E0] p-2 flex flex-col justify-between text-[10px] shrink-0">
                      <div className="space-y-2">
                        <div className="flex items-center justify-between border-b pb-1">
                          <span className="font-bold text-purple-700 flex items-center gap-1">
                            <Sparkles className="w-3 h-3" /> AI 분석 · 답안
                          </span>
                          <span className="px-1.5 py-0.5 rounded bg-purple-100 text-purple-800 text-[9px] font-bold">
                            Sonnet 5.5
                          </span>
                        </div>
                        <div className="grid grid-cols-5 gap-0.5 bg-slate-200 p-0.5 rounded text-[9px] font-semibold text-center">
                          <span className={previewState !== 'ai-dock-reply' ? 'bg-white rounded py-0.5 text-purple-700' : 'py-0.5'}>분석</span>
                          <span className={previewState === 'ai-dock-reply' ? 'bg-white rounded py-0.5 text-purple-700' : 'py-0.5'}>답안</span>
                          <span className="py-0.5">대화</span>
                          <span className="py-0.5">번역</span>
                          <span className="py-0.5">액션</span>
                        </div>

                        {previewState === 'ai-dock-reply' ? (
                          <div className="space-y-1.5">
                            <div className="p-1.5 rounded bg-white border border-purple-200">
                              <div className="font-bold text-purple-700">1. 기술·도면 검토 회신</div>
                              <p className="text-slate-600 mt-0.5">
                                보내주신 H042 1ST DECK CAMBER 기준 Z 좌표 및 Lifting lug 간섭 여부 확인 후 금일 중 회신드리겠습니다.
                              </p>
                            </div>
                            <div className="px-2 py-1 rounded bg-purple-600 text-white text-center font-semibold">
                              회신 창에 바로 삽입
                            </div>
                          </div>
                        ) : (
                          <div className="space-y-1.5">
                            <div className="p-1.5 rounded bg-white border">
                              <div className="font-bold text-slate-800">핵심 요약</div>
                              <p className="text-slate-600 mt-0.5">
                                • H042 Exmar SKID 3D 모델링 Z좌표 기준 전달 및 Lifting lug 구조 해석 대기
                              </p>
                            </div>
                            <div className="p-1.5 rounded bg-white border">
                              <div className="font-bold text-slate-800">식별된 도면 · 기한</div>
                              <p className="text-[#0F6CBD] font-mono">H042 · KYTS-1037 (오탐 0건)</p>
                            </div>
                          </div>
                        )}
                      </div>
                      <div className="text-[9px] text-slate-500">Claude 5.5 3단계 자동 라우팅</div>
                    </div>

                    {/* Col 5 (Split 2): Attachment Preview & Folder Batch Exporter Dock (18%) */}
                    <div className="w-[18%] bg-white p-2 flex flex-col justify-between text-[10px] shrink-0">
                      <div className="space-y-2">
                        <div className="grid grid-cols-2 gap-1 bg-slate-100 p-0.5 rounded text-[9px] font-bold text-center">
                          <span
                            className={
                              previewState !== 'batch-folder-export'
                                ? 'bg-[#0F6CBD] text-white rounded py-1'
                                : 'py-1 text-slate-600'
                            }
                          >
                            첨부 미리보기·분석
                          </span>
                          <span
                            className={
                              previewState === 'batch-folder-export'
                                ? 'bg-emerald-600 text-white rounded py-1'
                                : 'py-1 text-slate-600'
                            }
                          >
                            폴더 저장·리포트
                          </span>
                        </div>

                        {previewState === 'batch-folder-export' ? (
                          <div className="space-y-1.5">
                            <div className="p-1.5 rounded bg-emerald-50 border border-emerald-200">
                              <div className="font-bold text-emerald-800">저장 폴더 지정</div>
                              <div className="font-mono text-[9px] text-emerald-950 truncate">
                                D:\Bmail_Attachments\22K_FD
                              </div>
                            </div>
                            <div className="p-1.5 rounded bg-slate-50 border text-[9px] space-y-0.5">
                              <div>• 범위: 기간 + 보낸사람 + 프로젝트</div>
                              <div>• 폴더구조: 프로젝트별 / 날짜_발신자_원본</div>
                              <div>• 서명 이미지(image001.png) 자동 제외</div>
                            </div>
                            <div className="px-2 py-1 rounded bg-emerald-600 text-white text-center font-bold">
                              일괄 저장 + MD/CSV/HTML 리포트 생성 완료
                            </div>
                          </div>
                        ) : (
                          <div className="space-y-1.5">
                            <div className="h-24 rounded bg-slate-100 border border-slate-300 p-1.5 flex flex-col justify-between">
                              <span className="text-[9px] font-bold text-rose-700">
                                PDF 실시간 미리보기: 22K_TLC_2604.pdf
                              </span>
                              <div className="text-[9px] text-slate-600 bg-white p-1 rounded border">
                                CAISSON FLOATING DOCK EL. SYSTEM POS Deviation List...
                              </div>
                            </div>
                            <div className="p-1.5 rounded bg-sky-50 border border-sky-200 text-[9px]">
                              <div className="font-bold text-sky-800">AI 첨부 문서 분석</div>
                              <div className="text-slate-700">
                                초기 POS 대비 최종 변경 견적 및 SPEC DEVIATION 정산 필요
                              </div>
                            </div>
                          </div>
                        )}
                      </div>

                      {/* Bottom-Right Pinned Compact Mini Translation Window Replica */}
                      <div className="p-1.5 rounded-lg bg-[#0F6CBD] text-white shadow-md border border-white/20">
                        <div className="flex items-center justify-between text-[9px] font-bold">
                          <span>번역·요약 미니창 (우측 하단 고정)</span>
                          <span>무오류 200 OK</span>
                        </div>
                        <div className="text-[8px] text-sky-100 truncate">
                          {previewState === 'sgw-login-chat'
                            ? 'SGW 공식 로그인 창(https://sgw.seastar.work) 완전 연동됨'
                            : '좌측 트리를 가리지 않는 330px 우측 하단 컴팩트 모드'}
                        </div>
                      </div>
                    </div>
                  </div>
                )}

                {/* Spotlight Overlay in Motion Mode */}
                {mode === 'motion' && (
                  <div
                    className="pointer-events-none absolute border-2 border-sky-400 rounded-lg shadow-[0_0_0_9999px_rgba(2,6,23,0.42)] transition-all duration-700 ease-out"
                    style={{
                      left: `${currentScene.spotlight.x}%`,
                      top: `${currentScene.spotlight.y}%`,
                      width: `${currentScene.spotlight.w}%`,
                      height: `${currentScene.spotlight.h}%`,
                    }}
                  >
                    <span className="absolute -top-6 left-0 px-2 py-0.5 rounded bg-sky-500 text-white text-[10px] font-bold whitespace-nowrap shadow">
                      {currentScene.badgeText}
                    </span>
                  </div>
                )}

                {/* Interactive Numbered Icon Pins (Always clickable!) */}
                {filteredIcons.map((ic) => {
                  const isSelected = ic.id === activeIcon.id;
                  return (
                    <button
                      key={ic.id}
                      type="button"
                      onClick={() => {
                        setActiveIconId(ic.id);
                        if (mode === 'motion') setIsPlaying(false);
                      }}
                      onMouseEnter={() => {
                        if (mode === 'explorer') setActiveIconId(ic.id);
                      }}
                      title={`${ic.number}. ${ic.label}`}
                      style={{
                        left: `${ic.xPercent}%`,
                        top: `${ic.yPercent}%`,
                      }}
                      className={`absolute -translate-x-1/2 -translate-y-1/2 z-20 w-6 h-6 rounded-full text-[10px] font-extrabold flex items-center justify-center transition-all duration-200 ${
                        isSelected
                          ? 'bg-amber-400 text-slate-950 scale-125 ring-4 ring-amber-300/60 shadow-lg'
                          : 'bg-[#0F6CBD] text-white hover:scale-110 ring-2 ring-white shadow'
                      }`}
                    >
                      {ic.number}
                    </button>
                  );
                })}

                {/* Animated Motion-Graphic Virtual Cursor + Click Ripple */}
                <div
                  className="pointer-events-none absolute z-30 transition-all duration-700 ease-out"
                  style={{
                    left: `${cursorPos.x}%`,
                    top: `${cursorPos.y}%`,
                  }}
                >
                  {clickPulse && (
                    <span className="absolute -left-4 -top-4 w-8 h-8 rounded-full bg-amber-400/60 animate-ping" />
                  )}
                  <div className="relative flex items-center gap-1.5">
                    <svg
                      width="24"
                      height="24"
                      viewBox="0 0 24 24"
                      className="drop-shadow-[0_2px_6px_rgba(0,0,0,0.6)]"
                    >
                      <path
                        d="M4 3L11 20L14 13L21 10L4 3Z"
                        fill="#FBBF24"
                        stroke="#0F172A"
                        strokeWidth="2"
                      />
                    </svg>
                    <span className="px-2 py-0.5 rounded-full bg-slate-950/90 text-amber-300 border border-amber-400/50 text-[10px] font-bold whitespace-nowrap shadow-lg">
                      #{activeIcon.number} {activeIcon.label}
                    </span>
                  </div>
                </div>

                {/* Cinematic Motion-Graphic Subtitle Overlay */}
                {mode === 'motion' && (
                  <div className="absolute bottom-3 left-1/2 -translate-x-1/2 z-30 w-[92%] max-w-3xl bg-slate-950/92 backdrop-blur-md text-white px-4 py-2.5 rounded-xl border border-sky-400/40 shadow-2xl">
                    <div className="flex items-center justify-between gap-2 mb-0.5">
                      <span className="text-[11px] font-bold text-sky-300">
                        {currentScene.badgeText} — {currentScene.title}
                      </span>
                      <span className="text-[10px] font-mono text-slate-400">
                        씬 {sceneIdx + 1} / {MANUAL_MOTION_SCENES.length}
                      </span>
                    </div>
                    <p className="text-xs text-slate-100 leading-relaxed break-keep">
                      {currentScene.narration}
                    </p>
                  </div>
                )}
              </div>

              {/* Video Transport Control Bar */}
              <div className="mt-2.5 px-3 py-2 rounded-lg bg-slate-900 border border-slate-800 flex flex-wrap items-center justify-between gap-3 shrink-0">
                <div className="flex items-center gap-1.5">
                  <button
                    type="button"
                    onClick={() => {
                      setMode('motion');
                      setIsPlaying((p) => !p);
                    }}
                    className="px-3 py-1.5 rounded-md bg-[#0F6CBD] hover:bg-[#115EA3] text-white text-xs font-bold flex items-center gap-1.5"
                  >
                    {isPlaying && mode === 'motion' ? (
                      <>
                        <Pause className="w-3.5 h-3.5" /> 일시정지
                      </>
                    ) : (
                      <>
                        <Play className="w-3.5 h-3.5" /> 모션 재생
                      </>
                    )}
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      setMode('motion');
                      setSceneIdx((s) => (s - 1 + MANUAL_MOTION_SCENES.length) % MANUAL_MOTION_SCENES.length);
                      setElapsedInScene(0);
                    }}
                    title="이전 씬"
                    className="p-1.5 rounded bg-slate-800 hover:bg-slate-700 text-slate-200"
                  >
                    <SkipBack className="w-3.5 h-3.5" />
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      setMode('motion');
                      setSceneIdx((s) => (s + 1) % MANUAL_MOTION_SCENES.length);
                      setElapsedInScene(0);
                    }}
                    title="다음 씬"
                    className="p-1.5 rounded bg-slate-800 hover:bg-slate-700 text-slate-200"
                  >
                    <SkipForward className="w-3.5 h-3.5" />
                  </button>
                </div>

                {/* Progress Scrubber */}
                <div className="flex-1 min-w-[180px] flex items-center gap-2">
                  <span className="text-[11px] font-mono text-sky-300">
                    {formatTime(globalElapsedSec)}
                  </span>
                  <div className="flex-1 h-2 bg-slate-800 rounded-full overflow-hidden flex">
                    {MANUAL_MOTION_SCENES.map((sc, idx) => {
                      const pct =
                        idx < sceneIdx
                          ? 100
                          : idx === sceneIdx
                            ? Math.min(100, (elapsedInScene / sc.durationSec) * 100)
                            : 0;
                      return (
                        <button
                          key={sc.id}
                          type="button"
                          onClick={() => {
                            setMode('motion');
                            setSceneIdx(idx);
                            setElapsedInScene(0);
                          }}
                          style={{ width: `${(sc.durationSec / totalDurationSec) * 100}%` }}
                          className="h-full border-r border-slate-950 bg-slate-800 relative"
                          title={`${sc.step}. ${sc.title}`}
                        >
                          <span
                            className="block h-full bg-gradient-to-r from-sky-500 to-amber-400 transition-all duration-100"
                            style={{ width: `${pct}%` }}
                          />
                        </button>
                      );
                    })}
                  </div>
                  <span className="text-[11px] font-mono text-slate-400">
                    {formatTime(totalDurationSec)}
                  </span>
                </div>

                {/* Speed & Loop Controls */}
                <div className="flex items-center gap-1.5">
                  {([0.75, 1, 1.5] as const).map((sp) => (
                    <button
                      key={sp}
                      type="button"
                      onClick={() => setSpeed(sp)}
                      className={`px-2 py-1 rounded text-[11px] font-mono font-bold ${
                        speed === sp
                          ? 'bg-sky-500/30 text-sky-200 border border-sky-400/50'
                          : 'bg-slate-800 text-slate-400 hover:text-white'
                      }`}
                    >
                      {sp}x
                    </button>
                  ))}
                  <button
                    type="button"
                    onClick={() => setLoop((l) => !l)}
                    title="반복 재생"
                    className={`p-1.5 rounded text-xs ${
                      loop ? 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/40' : 'bg-slate-800 text-slate-400'
                    }`}
                  >
                    <Repeat className="w-3.5 h-3.5" />
                  </button>
                </div>
              </div>
            </div>

            {/* Right Inspector Column: Detailed Icon Function & Live Spec Card */}
            <aside className="w-full lg:w-[360px] bg-slate-900 border-t lg:border-t-0 lg:border-l border-slate-800 p-4 flex flex-col justify-between overflow-y-auto shrink-0">
              <div className="space-y-4">
                <div className="p-3.5 rounded-xl bg-slate-950 border border-sky-500/40 space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="px-2 py-0.5 rounded-full bg-amber-400 text-slate-950 text-[11px] font-extrabold">
                      아이콘 #{activeIcon.number}
                    </span>
                    {activeIcon.shortcut && (
                      <kbd className="px-2 py-0.5 rounded bg-slate-800 text-sky-200 border border-slate-700 text-[10px] font-mono">
                        {activeIcon.shortcut}
                      </kbd>
                    )}
                  </div>
                  <h2 className="text-sm font-bold text-white leading-snug">
                    {activeIcon.label}
                  </h2>
                  <p className="text-xs font-semibold text-sky-300 leading-relaxed">
                    {activeIcon.summary}
                  </p>
                  <p className="text-xs text-slate-300 leading-relaxed break-keep">
                    {activeIcon.detail}
                  </p>
                  <div className="pt-2 border-t border-slate-800 space-y-1 text-[11px]">
                    <div className="flex items-center justify-between text-slate-400">
                      <span>연결 API / 라우트</span>
                      <code className="text-emerald-300 font-mono text-[10px]">
                        {activeIcon.backendEndpoint}
                      </code>
                    </div>
                    <div className="p-2 rounded bg-sky-950/60 border border-sky-500/20 text-sky-200 text-[11px]">
                      <strong>실무 팁:</strong> {activeIcon.tip}
                    </div>
                  </div>
                </div>

                {/* All Icons Quick Directory List */}
                <div className="space-y-2">
                  <div className="text-xs font-bold text-slate-300 px-1">
                    화면 아이콘 전체 목록 (총 {filteredIcons.length}개 · 클릭 시 위치 이동)
                  </div>
                  <div className="flex flex-wrap items-center gap-1.5 px-1">
                    {iconJumpRanges.map((range) => {
                      const hasMatch = filteredIcons.some(
                        (ic) => ic.number >= range.start && ic.number <= range.end
                      );
                      const isActiveRange = highlightedRange === range.key;
                      return (
                        <button
                          key={range.key}
                          type="button"
                          disabled={!hasMatch}
                          onClick={() => handleJumpToRange(range)}
                          className={`px-2 py-0.5 rounded-full text-[11px] font-semibold border transition-colors ${
                            !hasMatch
                              ? 'bg-slate-900 text-slate-600 border-slate-800 cursor-not-allowed'
                              : isActiveRange
                                ? 'bg-sky-500/25 text-sky-200 border-sky-400/60'
                                : 'bg-slate-800/90 text-slate-300 border-slate-700 hover:bg-slate-700 hover:text-white'
                          }`}
                        >
                          {range.label}
                        </button>
                      );
                    })}
                  </div>
                  <div className="space-y-1 pr-1">
                    {filteredIcons.map((ic) => {
                      const activeRangeObj = iconJumpRanges.find((r) => r.key === highlightedRange);
                      const inHighlightedRange =
                        !!activeRangeObj &&
                        ic.number >= activeRangeObj.start &&
                        ic.number <= activeRangeObj.end;
                      return (
                        <button
                          key={ic.id}
                          ref={(el) => {
                            iconItemRefs.current[ic.id] = el;
                          }}
                          type="button"
                          onClick={() => {
                            setActiveIconId(ic.id);
                            setIsPlaying(false);
                          }}
                          className={`w-full text-left px-2.5 py-2 rounded-lg text-xs flex items-center justify-between gap-2 transition-colors ${
                            ic.id === activeIcon.id
                              ? 'bg-[#0F6CBD] text-white font-semibold'
                              : inHighlightedRange
                                ? 'bg-sky-950/60 hover:bg-slate-800 text-sky-100 ring-1 ring-sky-400/40'
                                : 'bg-slate-950/60 hover:bg-slate-800 text-slate-300'
                          }`}
                        >
                          <div className="flex items-center gap-2 min-w-0">
                            <span className="w-5 h-5 rounded-full bg-black/30 text-[10px] font-bold flex items-center justify-center shrink-0">
                              {ic.number}
                            </span>
                            <span className="truncate">{ic.label}</span>
                          </div>
                          <ChevronRight className="w-3.5 h-3.5 shrink-0 opacity-75" />
                        </button>
                      );
                    })}
                  </div>
                </div>
              </div>

              <div className="mt-4 pt-3 border-t border-slate-800 text-[11px] text-slate-400 flex items-center justify-between shrink-0">
                <span>Bmail Outlook AI v{MANUAL_VERSION}</span>
                <span>기능 업데이트 자동 연동됨</span>
              </div>
            </aside>
          </div>
        )}
      </div>
    </div>
  );
};
