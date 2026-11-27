/**
 * Single-Source-of-Truth Feature, Icon & Motion-Graphic Manual Catalog
 *
 * Whenever a UI button, dock, or workflow is added/updated in Bmail Outlook AI,
 * updating this catalog automatically updates:
 *   1. In-app Interactive HTML Manual Modal (`InteractiveManualModal.tsx`)
 *   2. Standalone HTML Popup Window (`/?manual=1`)
 *   3. Motion-Graphic Video Walkthrough Player (13 scenes)
 *   4. Versioned Release Changelog (`MANUAL_CHANGELOG`)
 *
 * Enforced by `src/data/manualCatalog.test.ts` (verifies `MANUAL_VERSION === package.json.version`).
 */

export const MANUAL_VERSION = '1.3.1';
export const MANUAL_UPDATED_AT = '2026-11-27';

export type ManualZoneId =
  | 'header'
  | 'ribbon'
  | 'folders'
  | 'maillist'
  | 'reading'
  | 'ai-dock'
  | 'attach-dock'
  | 'mini-assist'
  | 'dashboard'
  | 'sgw-chat';

export interface ManualZoneInfo {
  id: ManualZoneId;
  label: string;
  description: string;
  accentColor: string;
}

export interface ManualIconEntry {
  id: string;
  number: number;
  zone: ManualZoneId;
  label: string;
  shortcut?: string;
  /** X coordinate (0-100%) on the 1:1 visual screen replica */
  xPercent: number;
  /** Y coordinate (0-100%) on the 1:1 visual screen replica */
  yPercent: number;
  summary: string;
  detail: string;
  backendEndpoint: string;
  tip: string;
  addedInVersion: string;
}

export type ManualPreviewStateId =
  | 'outlook-overview'
  | 'live-sync'
  | 'rich-reading'
  | 'ai-dock-reply'
  | 'attach-preview'
  | 'batch-folder-export'
  | 'executive-dashboard'
  | 'sgw-login-chat'
  | 'outlook-companion';

export interface ManualMotionScene {
  id: string;
  step: number;
  title: string;
  subtitle: string;
  narration: string;
  durationSec: number;
  targetIconId: string;
  zone: ManualZoneId;
  cursor: { x: number; y: number };
  spotlight: { x: number; y: number; w: number; h: number };
  previewState: ManualPreviewStateId;
  badgeText: string;
}

export interface ManualChangelogEntry {
  version: string;
  date: string;
  headline: string;
  highlights: string[];
}

export const MANUAL_ZONES: ManualZoneInfo[] = [
  {
    id: 'header',
    label: '상단 글로벌 바',
    description: '통합 검색, 보내기/받기 원클릭(300ms 자동), Outlook 동반 검증(기본 시작), 매뉴얼, POP3/IMAP, SGW 채팅, 설정 진입(비연동 5개 메뉴 정리)',
    accentColor: '#0F6CBD',
  },
  {
    id: 'ribbon',
    label: 'Classic Outlook 평면 리본',
    description: '오리지널 무채색 평면 1px 리본(새 메일, 회신/전달, 증분 동기화, 미확인 브리핑 토글, 우측 도크)',
    accentColor: '#115EA3',
  },
  {
    id: 'folders',
    label: '즐겨찾기 · 저장소 폴더 트리',
    description: '하위 폴더 아이콘 간소화 · 경계선 가로 드래그 크기 조절(bmail_pane_widths_v1) 및 Outlook 즐겨찾기·평면 트리',
    accentColor: '#0B7A75',
  },
  {
    id: 'maillist',
    label: '평면 1px 메일 목록 · 경계선 크기 조절',
    description: '4개 구역 경계선 마우스 드래그 폭 조절·더블클릭 초기화(bmail_pane_widths_v1) · 아바타 OFF 1px 목록',
    accentColor: '#2563EB',
  },
  {
    id: 'reading',
    label: '본문 읽기 창 (HTML · AI 요약 보기 토글)',
    description: '경계선 가로 드래그 폭 조절 + 원본 서식·인라인(cid:) 도면 렌더링 및 인라인 AI 요약 토글',
    accentColor: '#4F46E5',
  },
  {
    id: 'ai-dock',
    label: '우측 1분할: AI 분석 · 0ms 요약 도크',
    description: '경계선 가로 드래그 폭 조절 + 메일 클릭 0ms 즉시 요약 · 수동/자동 토글 · Groq/Claude 라우팅',
    accentColor: '#7C3AED',
  },
  {
    id: 'attach-dock',
    label: '우측 2분할: 첨부 미리보기 · 폴더 저장',
    description: 'PDF/이미지/엑셀 즉시 미리보기·AI 문서분석 및 지정 폴더 일괄 저장+3종 리포트',
    accentColor: '#059669',
  },
  {
    id: 'mini-assist',
    label: '우측 하단: 14일 미확인 자동 알림 · 미니창',
    description: '최근 14일 미확인 수신 메일 자동 알림 토스트([30분 뒤] · [×] 재알림) 및 무오류 번역 미니창',
    accentColor: '#D97706',
  },
  {
    id: 'dashboard',
    label: '종합 분석글 · 주간/월간 대시보드',
    description: '1회 로드 영구저장 + 증분 업데이트 기반 분석글·주간/월간 테이블·내게 온 미회신건·프로젝트/주제별 정리',
    accentColor: '#DC2626',
  },
  {
    id: 'sgw-chat',
    label: 'SGW 공식 로그인 창 메신저 연동',
    description: 'SGW 그룹웨어 실제 로그인 창(https://sgw.seastar.work) 직결 세션 연동 및 실시간 채팅',
    accentColor: '#0284C7',
  },
];

export const MANUAL_ICONS: ManualIconEntry[] = [
  {
    id: 'icon-app-launcher',
    number: 1,
    zone: 'header',
    label: '앱 런처 (정리된 네비게이션 · 기본 시작: 동반 검증)',
    xPercent: 2.2,
    yPercent: 4.5,
    summary: '기본 시작 화면 Outlook 동반 검증(companion) · Outlook 메일함·종합 대시보드·SGW 허브 전환 (비연동 5개 메뉴 정리)',
    detail: 'v1.3.1부터 기본 시작 화면을 Outlook 동반 검증(companion)으로 전환하고, Outlook 비연동·중복 5개 메뉴(일정·달력, 할 일·To-Do, 기록함, AI 브리핑·Ask AI, 플러그인 설치)를 네비게이션에서 깔끔하게 정리했습니다. 로컬 서버는 127.0.0.1 바인딩, 루프백/Origin 검증 및 민감 경로 차단으로 보안이 강화되었습니다.',
    backendEndpoint: 'GET /api/local/mail/bootstrap',
    tip: '개발용 내부 문서(HANDOFF/소스 다운로드)와 미연동 메뉴는 완제품 화면에서 숨김 처리되었습니다.',
    addedInVersion: '1.2.1',
  },
  {
    id: 'icon-global-search',
    number: 2,
    zone: 'header',
    label: 'FTS5 초고속 통합 검색 · 필터',
    shortcut: 'Ctrl + K',
    xPercent: 34.0,
    yPercent: 4.5,
    summary: '24,600+ 메일 제목·발신자·호선(22K FD, H042)·도면번호·본문 즉시 검색',
    detail: 'SQLite FTS5 전문 검색 엔진과 연동되어 한글·영문 키워드, 기간, 첨부 유무별로 0.05초 내에 필터링합니다.',
    backendEndpoint: 'GET /api/local/mail/emails?q=...',
    tip: '우측 필터 버튼에서 기간·발신자·첨부파일 조건을 조합할 수 있습니다.',
    addedInVersion: '1.1.0',
  },
  {
    id: 'icon-interactive-manual',
    number: 3,
    zone: 'header',
    label: '인터랙티브 HTML 매뉴얼 (모션 가이드)',
    shortcut: 'F1',
    xPercent: 77.5,
    yPercent: 4.5,
    summary: '현재 화면을 1:1 그대로 보여주는 HTML 팝업 매뉴얼 + 동영상 모션그래픽 투어',
    detail: '각 아이콘의 기능을 실제 화면 배치 그대로 설명하며, 우측 전체 목록(1~10 · 11~20 · 21~끝 빠른 이동)과 모션그래픽 자동 재생·독립 HTML 팝업창(/?manual=1)을 지원합니다.',
    backendEndpoint: 'GET /?manual=1',
    tip: '기능 업데이트 시 단일 카탈로그(manualCatalog.ts)에서 자동 동기화되며 잘림 없이 25개+ 전체 아이콘을 탐색할 수 있습니다.',
    addedInVersion: '1.2.1',
  },
  {
    id: 'icon-pop3-connect',
    number: 4,
    zone: 'header',
    label: '메일플러그 · POP3/IMAP 연결',
    xPercent: 83.2,
    yPercent: 4.5,
    summary: '사내 메일플러그·표준 POP3/IMAP 서버 자동 감지 및 백그라운드 수신',
    detail: 'Outlook 레지스트리의 계정 정보를 자동 인식하여 서버 메일을 실시간 동기화합니다.',
    backendEndpoint: 'POST /api/pop3/connect',
    tip: '초록색 펄스 점은 백그라운드 동기화가 정상 작동 중임을 뜻합니다.',
    addedInVersion: '1.1.0',
  },
  {
    id: 'icon-sgw-chat-header',
    number: 5,
    zone: 'sgw-chat',
    label: 'SGW 공식 로그인 창 메신저 연동',
    xPercent: 88.4,
    yPercent: 4.5,
    summary: 'SGW(https://sgw.seastar.work) 공식 로그인 창을 띄워 완전 로그인 연동',
    detail: '전용 로그인 브라우저 창에서 로그인하면 인증 토큰/쿠키 세션을 AES-256-GCM으로 안전하게 연결하여 사내 대화방·쪽지·메일 공유를 즉시 수행합니다.',
    backendEndpoint: 'POST /api/sgw-chat/session-connect',
    tip: '비밀번호를 저장하지 않고 SGW 공식 로그인 창 세션을 그대로 연동합니다.',
    addedInVersion: '1.2.1',
  },
  {
    id: 'icon-calendar-header',
    number: 6,
    zone: 'header',
    label: 'Outlook 일정 · 액션 상태 SQLite 저장 (네비게이션 정리)',
    xPercent: 91.5,
    yPercent: 4.5,
    summary: '비연동 독립 일정 메뉴는 네비게이션에서 정리되고, AI 도크 액션 탭·SQLite 상태 저장(PUT /api/local/state/:key)으로 연동 유지',
    detail: 'v1.3.1에서 Outlook 비연동·중복 5개 메뉴(일정·달력, 할 일·To-Do, 기록함, AI 브리핑·Ask AI, 플러그인 설치)를 상단/런처 네비게이션에서 숨기고, 메일별 후속 일정과 상태 저장은 PUT /api/local/state/:key SQLite 저장소를 통해 안정적으로 유지합니다.',
    backendEndpoint: 'PUT /api/local/state/:key',
    tip: '메일 본문 AI 도크의 액션 탭에서 추출된 납기일과 후속 조치를 즉시 저장할 수 있습니다.',
    addedInVersion: '1.1.0',
  },
  {
    id: 'icon-outlook-settings',
    number: 7,
    zone: 'header',
    label: 'Outlook 환경 · 규칙 · 서명 · AI Claude API 설정',
    xPercent: 96.5,
    yPercent: 4.5,
    summary: 'PC Outlook 전체 설정 탐색·적용 · AI · Claude API 탭(Claude API 직접 설정 카드) & Groq 5-키 일괄 설정',
    detail: 'Windows 레지스트리 및 %APPDATA%\\Microsoft\\Outlook 전체 설정을 자동 탐색·적용하고 명함 HTML 로고 이미지를 base64 인라인으로 반영하며, [AI · Claude API] 탭 및 설정(ConfigEnvManager) 최상단의 Claude API 직접 설정 카드와 npm run groq:keys / groq-keys.txt 5-키 설정을 제공합니다.',
    backendEndpoint: 'POST /api/local/outlook/sync-incremental · /api/local/ai/claude-config',
    tip: '단일 줄 리본, 폴더 창 고정, 명함 HTML 서명 로고는 물론 ANTHROPIC_API_KEY(sk-ant-...) 저장까지 한 곳에서 설정합니다.',
    addedInVersion: '1.2.1',
  },
  {
    id: 'icon-ribbon-new-mail',
    number: 8,
    zone: 'ribbon',
    label: '새 메일 작성 (오리지널 Outlook 작성창 안전 위임)',
    shortcut: 'Ctrl + N',
    xPercent: 6.5,
    yPercent: 11.2,
    summary: '직접 SMTP/COM 발송 대신 항상 오리지널 Classic Outlook 작성창(MailItem.Display())으로 안전하게 위임',
    detail: '새 메일 작성 시 직접 백그라운드 발송하지 않고 오리지널 Classic Outlook 작성창(MailItem.Display())을 열어 마스터의 Outlook 서명·보안·발신 계정을 그대로 사용하여 발송 사고를 원천 방지합니다.',
    backendEndpoint: 'POST /api/local/outlook/send-via-com (action: display)',
    tip: '실제 발송 전 클래식 Outlook 창에서 최종 확인 후 직접 [보내기]를 누를 수 있어 안전합니다.',
    addedInVersion: '1.1.0',
  },
  {
    id: 'icon-ribbon-reply-all',
    number: 9,
    zone: 'ribbon',
    label: '회신 · 전체 회신 · 전달 (MailItem.Display() 위임)',
    shortcut: 'Ctrl + R',
    xPercent: 16.5,
    yPercent: 11.2,
    summary: 'AI 답안·초안을 오리지널 Classic Outlook 작성창(MailItem.Display())으로 안전 위임',
    detail: '우측 AI 답안 도크 및 동반 검증 패널에서 생성한 회신 초안을 직접 SMTP/COM 발송하지 않고 오리지널 Classic Outlook 작성창(MailItem.Display(), POST /api/local/outlook/send-via-com action: display)에 채워 열어줍니다.',
    backendEndpoint: 'POST /api/local/outlook/send-via-com',
    tip: '원본 스레드·CC·서명·첨부가 오리지널 Outlook 작성창에 그대로 유지됩니다.',
    addedInVersion: '1.1.0',
  },
  {
    id: 'icon-ribbon-live-sync',
    number: 10,
    zone: 'ribbon',
    label: '보내기/받기 · 증분 동기화 (300ms 자동 실행)',
    shortcut: 'F9',
    xPercent: 32.5,
    yPercent: 11.2,
    summary: '앱 시작 즉시(300ms) Outlook 보내기/받기(sendAndReceive: true) 및 비동기 증분 동기화 실행',
    detail: '앱 시작 후 300ms 시점에 Outlook 보내기/받기와 새 메일 탐색(POST /api/local/outlook/sync-incremental, sendAndReceive: true)을 자동 실행하며, 상단 헤더 및 동반 검증 패널의 [보내기/받기](btn-send-receive-refresh) 버튼으로도 언제든 원클릭 동기화할 수 있습니다.',
    backendEndpoint: 'POST /api/local/outlook/sync-incremental',
    tip: '24,626건 전체를 다시 읽지 않고 새 메일만 즉시 반영해 대시보드·UI 지연이 없습니다.',
    addedInVersion: '1.2.1',
  },
  {
    id: 'icon-ribbon-toggle-docks',
    number: 11,
    zone: 'ribbon',
    label: '우측 2분할 도크 (AI 분석창 · 첨부 미리보기창) 토글',
    xPercent: 48.0,
    yPercent: 11.2,
    summary: '오른쪽 세로분할 AI 분석 도크(272~296px) 상시 고정 및 첨부파일 클릭 우측 미리보기 전환',
    detail: '본문 우측에 272~296px 폭의 AI 분석 도크를 상시 고정해 메일 선택 즉시 핵심 분석과 답안을 확인하고, 첨부파일 클릭 시 우측 미리보기 도크로 즉시 전환합니다.',
    backendEndpoint: 'POST /api/local/mail/assistant',
    tip: '미확인 메일은 읽지않음 상태를 유지한 채 상단 UnreadBriefingBanner로 중요도순 브리핑을 제공합니다.',
    addedInVersion: '1.2.1',
  },
  {
    id: 'icon-ribbon-dashboard-jump',
    number: 12,
    zone: 'dashboard',
    label: '주간·월간 분석글 · 미회신 · 사용량 기준 대시보드',
    xPercent: 62.0,
    yPercent: 11.2,
    summary: '분석글·주간/월간 테이블·미회신건·관리자 사용량 기준 게이지 & [기준 수시 변경] 대시보드 열기',
    detail: '저장된 전체 메일 DB 기반의 기간·프로젝트·주제별 분석글 및 미회신 건과 함께, 관리자 LLM 사용량 기준 게이지·일별 호출 SVG 기준선·[기준 수시 변경] 제어를 한눈에 보여줍니다.',
    backendEndpoint: 'GET /api/local/dashboard/executive',
    tip: '대시보드 상단에서 기준 대비 사용량과 라우팅 모드(절약형/표준/대용량)를 실시간으로 확인·변경할 수 있습니다.',
    addedInVersion: '1.2.1',
  },
  {
    id: 'icon-tree-favorites',
    number: 13,
    zone: 'folders',
    label: 'Outlook 즐겨찾기 (New2026.xml 완벽 연동)',
    xPercent: 7.5,
    yPercent: 22.0,
    summary: '실제 Outlook에 설정된 즐겨찾기 폴더(받은 편지함, 보낸 편지함, 선보, 유일 등) 자동 표시',
    detail: '%APPDATA%\\Microsoft\\Outlook\\New2026.xml 파일을 직접 파싱하여 마스터의 Outlook 즐겨찾기 구성을 그대로 복원합니다.',
    backendEndpoint: 'GET /api/local/mail/bootstrap',
    tip: '폴더 우측 별표 아이콘으로 즐겨찾기를 추가·해제할 수 있습니다.',
    addedInVersion: '1.2.1',
  },
  {
    id: 'icon-tree-stores',
    number: 14,
    zone: 'folders',
    label: '저장소별 평면 1px 폴더 트리 (김봉정 · 백업)',
    xPercent: 7.5,
    yPercent: 45.0,
    summary: '하위 폴더 아이콘을 간소화한 오리지널 클래식 평면 트리 · 파란색 미읽음 카운트',
    detail: '받은 편지함 하위(22K FD / 22K POS / 선급승도도 제출, 한라ims, DOCK 등) 다단계 폴더의 불필요한 아이콘·둥근 박스를 정리하고 1px 경계선 기반 클래식 구조와 미읽음 수 표기를 재현합니다.',
    backendEndpoint: 'GET /api/local/mail/emails?folder=...',
    tip: '상단 폴더 빠른 검색창에 호선명(예: 22K, 선보)을 입력하면 SQLite 복합 인덱스(idx_emails_active_ts)로 즉시 조회됩니다.',
    addedInVersion: '1.2.1',
  },
  {
    id: 'icon-maillist-row',
    number: 15,
    zone: 'maillist',
    label: '평면 1px 그룹 메일 리스트 (아바타 기본 OFF)',
    xPercent: 22.5,
    yPercent: 34.0,
    summary: '둥근 박스·원형 아바타를 기본 OFF로 정리한 오리지널 Classic Outlook 1px 경계선 목록',
    detail: '발신자 원형 아바타와 과도한 둥근 배지를 제거해 한 화면 표시 밀도를 높였으며, 미읽음 메일은 좌측 #0F6CBD 파란 수직선과 굵은 제목으로 명확히 구분됩니다.',
    backendEndpoint: 'PATCH /api/local/mail/emails/:id',
    tip: 'idx_emails_unread_inbox 복합 인덱스로 미확인 메일 필터링과 목록 전환이 즉시 반응합니다.',
    addedInVersion: '1.2.1',
  },
  {
    id: 'icon-reading-rich-html',
    number: 16,
    zone: 'reading',
    label: '원본 HTML 서식 · 인라인(cid:) 도면 이미지 본문',
    xPercent: 46.0,
    yPercent: 48.0,
    summary: '본문 내 3D 모델링 캡처 이미지(cid:), 표, 글자 색상/하이라이트 원본 그대로 렌더링',
    detail: 'PST 내부의 HTMLBody와 cid: 인라인 이미지를 자동 매핑하여 본문 안에 선명하게 표시하고, 메일 명함·네임태그·서명 인라인 이미지(image.png, image001.png)가 첨부파일 칩을 어지럽히지 않도록 자동 제외합니다.',
    backendEndpoint: 'GET /api/local/emails/:id/render',
    tip: '도면 번호 정규식이 개선되어 일반 영단어(MESSAGE 등)는 도면으로 오탐되지 않습니다.',
    addedInVersion: '1.2.1',
  },
  {
    id: 'icon-reading-attachment-chips',
    number: 17,
    zone: 'reading',
    label: '실첨부 전용 파일 칩 (1클릭 우측 미리보기 연동)',
    xPercent: 46.0,
    yPercent: 28.0,
    summary: '명함·네임태그 인라인 이미지(image.png)를 제외한 실제 업무 첨부파일만 깔끔하게 표시',
    detail: '첨부 칩을 클릭하면 다운로드 없이도 우측 미리보기 도크에서 즉시 내용을 확인하고 AI로 분석할 수 있으며, 서명·네임태그 인라인 이미지는 자동 필터링됩니다.',
    backendEndpoint: 'GET /api/local/pst/attachment?inline=1',
    tip: '우측 다운로드 아이콘을 누르면 원본 파일명(UTF-8) 그대로 저장됩니다.',
    addedInVersion: '1.2.1',
  },
  {
    id: 'icon-aidock-analyze',
    number: 18,
    zone: 'ai-dock',
    label: 'AI 분석 탭 (실제 LLM 자동 분석 · Groq AI / Claude 배지)',
    xPercent: 68.0,
    yPercent: 21.0,
    summary: '기본 Groq 5-키 풀 실제 LLM 자동 분석 · 긴급/선급/도면 및 [Claude 정밀] 시 Claude 5.5 에스컬레이션',
    detail: 'MailAiDockPanel 및 동반 검증 패널에서 상투적 오프라인 문구 없이 실제 LLM 자동 분석을 수행합니다. 기본은 Groq 5-키 풀(Groq AI 배지)로 고속 분석하고 긴급·선급·도면 메일이나 [Claude 정밀] 클릭 시 Claude 5.5(Claude 배지)로 자동 에스컬레이션하며, 이메일 주소(elcon0955@naver.com 등) 도면 번호 오탐도 교정되었습니다.',
    backendEndpoint: 'POST /api/local/mail/assistant',
    tip: '상단 엔진 배지(Groq AI / Claude)로 현재 실제 사용 중인 LLM 프로바이더를 투명하게 확인할 수 있습니다.',
    addedInVersion: '1.2.1',
  },
  {
    id: 'icon-aidock-reply',
    number: 19,
    zone: 'ai-dock',
    label: 'AI 답안 탭 (3종 톤앤매너 맞춤 회신)',
    xPercent: 71.5,
    yPercent: 21.0,
    summary: '정중·공식 / 기술·도면검토 / 간결·일정확인 3가지 회신 초안 생성 및 회신창 삽입',
    detail: '추가 지시어(예: "내일 오전까지 수정 POS 송부 예정")를 입력해 맞춤 답안을 다시 생성하고 1클릭으로 회신 창에 넣을 수 있습니다.',
    backendEndpoint: 'POST /api/local/mail/assistant',
    tip: '영문 메일일 경우 한/영 병기 또는 비즈니스 영문 답안을 자동 작성합니다.',
    addedInVersion: '1.2.1',
  },
  {
    id: 'icon-aidock-chat-translate-action',
    number: 20,
    zone: 'ai-dock',
    label: 'AI 대화 · 번역 · 액션 상태 탭',
    xPercent: 76.5,
    yPercent: 21.0,
    summary: '메일별 다중 턴 AI 질의응답, 4개국어 정밀 번역, 후속 조치 상태(대기/진행/완료) 저장',
    detail: '메일마다 나눈 AI 대화 이력과 액션 상태가 로컬 SQLite에 영구 보존되어 언제 다시 열어도 이어서 작업할 수 있습니다.',
    backendEndpoint: 'PUT /api/local/state/mail_action_states',
    tip: '액션 탭에서 마감일을 지정해 Outlook 일정 달력으로 바로 보낼 수 있습니다.',
    addedInVersion: '1.2.1',
  },
  {
    id: 'icon-attachdock-preview',
    number: 21,
    zone: 'attach-dock',
    label: '첨부 미리보기 · AI 문서 분석 탭',
    xPercent: 86.0,
    yPercent: 21.0,
    summary: '우측 세로 2번째 분할창에서 PDF·이미지·Excel·Word 즉시 미리보기 및 AI 사양 분석',
    detail: 'PDF 도면/POS는 내장 뷰어로, Excel/Word는 표·텍스트 구조로 즉시 렌더링하며 "AI 첨부 문서 정밀 분석"으로 핵심 스펙과 리스크를 추출합니다.',
    backendEndpoint: 'POST /api/local/pst/attachment/parse',
    tip: '메일 명함·네임태그 인라인 이미지(image.png)는 기본 자동 제외되며, 필요 시 상단 토글로 전체 확인이 가능합니다.',
    addedInVersion: '1.2.1',
  },
  {
    id: 'icon-attachdock-batch-export',
    number: 22,
    zone: 'attach-dock',
    label: '지정 폴더 일괄 첨부 저장 · 3종 리포트 생성 탭',
    xPercent: 93.5,
    yPercent: 21.0,
    summary: '기간·보낸사람·프로젝트·전체/일부 모든 옵션으로 지정 폴더에 첨부 저장 + 리포트 작성',
    detail: '로컬 폴더를 지정한 뒤 범위(전체/현재폴더/선택메일/프로젝트/기간), 발신자, 파일종류, 폴더 구조(프로젝트별/발신자별/월별/원본폴더), 파일명 규칙을 선택해 일괄 추출하고 Markdown·Excel CSV(UTF-8 BOM)·HTML 리포트를 자동 생성합니다.',
    backendEndpoint: 'POST /api/local/attachments/export-run',
    tip: '동일 파일 자동 스킵(Skip-Identical)과 서명 이미지 자동 제외 기능으로 깔끔하게 정리됩니다.',
    addedInVersion: '1.2.1',
  },
  {
    id: 'icon-mini-assist-pinned',
    number: 23,
    zone: 'mini-assist',
    label: '우측 하단 소형 고정 번역 · 요약 미니창',
    xPercent: 91.0,
    yPercent: 88.0,
    summary: '좌측 폴더 트리를 가리지 않도록 우측 하단(330px)에 작게 고정된 무오류 번역·요약 창',
    detail: '번역·할 일 / 회신 초안 / 영업·핵심수치 3개 탭을 제공하며, 오프라인이나 서버 재시작 중에도 로컬 엔진으로 폴백되어 Failed to fetch 에러가 절대 발생하지 않습니다.',
    backendEndpoint: 'POST /api/mail/mini-assist',
    tip: '최소화 버튼을 누르면 우측 하단 작은 알약 버튼으로 접혀 화면을 넓게 쓸 수 있습니다.',
    addedInVersion: '1.2.1',
  },
  {
    id: 'icon-exec-narrative-report',
    number: 24,
    zone: 'dashboard',
    label: '종합 분석글 · 주간/월간 보고서 · 사용량 기준 게이지',
    xPercent: 52.0,
    yPercent: 18.0,
    summary: '서술형 종합 분석글 + 관리자 사용량 기준 게이지(일일 호출·토큰·월간 비용 예산)·비교 도표·SVG 기준선 & [기준 수시 변경]',
    detail: '기간 추이·집중 프로젝트·미회신 리스크를 서술형 분석글과 주간/월간 테이블로 생성하며, 관리자 사용량 기준 게이지(일일 호출·토큰·월간 비용 예산)·기준 대비 비교 도표·일별 호출 SVG 기준선을 상시 표시하고 [기준 수시 변경](PUT /api/admin/llm-usage/limits)으로 절약형/표준/대용량 프리셋 및 Groq 병렬 슬롯·고급 분석 모드를 수시 전환합니다.',
    backendEndpoint: 'PUT /api/admin/llm-usage/limits',
    tip: '[기준 수시 변경]에서 저장한 한도와 프리셋·Groq 병렬 슬롯·고급 분석 모드는 즉시 실시간 AI 라우팅에 반영됩니다.',
    addedInVersion: '1.2.1',
  },
  {
    id: 'icon-exec-unreplied-tracker',
    number: 25,
    zone: 'dashboard',
    label: '내게 왔는데 미회신건 정리 · 프로젝트/주제별 분류',
    xPercent: 58.0,
    yPercent: 18.0,
    summary: '수신 메일 중 내가 아직 답장하지 않은 미회신 건만 자동 추출 + 프로젝트·주제별 교차 필터',
    detail: '전체 대화 스레드(conversationKey)의 수신 시각과 내 발신(보낸 편지함) 시각을 전수 대조하여 실제 미회신 건만 정확히 골라내고 경과일(D+N)과 긴급도 순으로 정리합니다.',
    backendEndpoint: 'GET /api/local/dashboard/executive',
    tip: '최근 14일(maxAgeDays=14) 기준 UnreadBriefingBanner 플로팅 알약/사이드 팝업과 연동되어 오래된 과거 메일 노이즈 없이 긴급 건만 집중 관리합니다.',
    addedInVersion: '1.2.1',
  },
  {
    id: 'icon-ribbon-unread-briefing-toggle',
    number: 26,
    zone: 'ribbon',
    label: '리본 [미확인 브리핑] 토글 (최근 14일 기준)',
    xPercent: 76.0,
    yPercent: 11.2,
    summary: '무채색 평면 리본에서 최근 14일(maxAgeDays=14) 미확인 브리핑 팝업을 즉시 켜고 끄기',
    detail: '4년 전 등 장기 과거 미열람 메일을 자동 제외한 최근 14일 이내 미확인 메일 브리핑(UnreadBriefingBanner)을 리본 버튼 클릭 한 번으로 표시하거나 접습니다.',
    backendEndpoint: 'GET /api/local/mail/unread-briefing?maxAgeDays=14',
    tip: '우측 하단 플로팅 알약으로도 접힌 브리핑 건수를 상시 확인할 수 있습니다.',
    addedInVersion: '1.2.9',
  },
  {
    id: 'icon-reading-ai-summary-toggle',
    number: 27,
    zone: 'reading',
    label: '읽기 창 [AI 요약 보기] 토글',
    xPercent: 59.0,
    yPercent: 21.0,
    summary: '본문 상단 인라인 AI 요약을 기본 숨김으로 정리하고 필요 시 [AI 요약 보기]로 즉시 펼치기',
    detail: '오리지널 클래식 Outlook 읽기 경험을 해치지 않도록 본문 내 인라인 AI 요약 박스를 기본 숨김 처리하고, 상단 [AI 요약 보기] 토글을 클릭할 때만 본문 위에 요약을 펼쳐 보여줍니다.',
    backendEndpoint: 'POST /api/local/mail/assistant',
    tip: '기본 본문을 넓게 읽다가 핵심 파악이 필요할 때만 토글해 보세요.',
    addedInVersion: '1.2.9',
  },
  {
    id: 'icon-aidock-instant-manual-toggle',
    number: 28,
    zone: 'ai-dock',
    label: 'AI 도크 0ms 즉시 요약 · [수동/자동] 토글',
    xPercent: 79.5,
    yPercent: 17.0,
    summary: '메일 클릭 즉시 0ms 결정론적 요약 표시 + AI 심층 분석 [수동/자동] 전환 스위치',
    detail: '메일을 클릭하자마자 대기 시간 없이 0ms 즉시 요약을 띄우며, 상단 [수동/자동] 토글로 자동 심층 AI 분석 실행 여부를 제어해 API 토큰과 속도를 최적화합니다.',
    backendEndpoint: 'POST /api/local/mail/assistant',
    tip: '빠르게 메일을 훑을 때는 수동 모드로 전환하면 불필요한 LLM 호출 없이 0ms 요약만 확인할 수 있습니다.',
    addedInVersion: '1.2.9',
  },
  {
    id: 'icon-floating-unread-briefing-pill',
    number: 29,
    zone: 'mini-assist',
    label: '최근 14일 미확인 자동 알림 팝업 · [30분 뒤] · [×] 재알림',
    xPercent: 85.0,
    yPercent: 88.0,
    summary: '최근 14일 이내 미확인 수신 메일 자동 알림 토스트(우측 하단) + [30분 뒤] 타이머 & [×] 후 새 메일 자동 재알림',
    detail: '최근 14일 이내 미확인 수신 메일이 감지되면 우측 하단에 비간섭 자동 알림 토스트 팝업이 열립니다. [30분 뒤]를 누르면 30분 다시 알림 타이머가 작동하고, [×] 닫기 후 새 미확인 메일이 도착하면 자동으로 재알림되며 기존 2023 장기 과거 메일 제외 정책을 유지합니다.',
    backendEndpoint: 'GET /api/local/mail/unread-briefing?maxAgeDays=14',
    tip: '읽음 상태를 건드리지 않고(읽지않음 유지) 새 미확인 수신 건만 놓치지 않도록 알립니다.',
    addedInVersion: '1.2.9',
  },
  {
    id: 'icon-dashboard-usage-limits',
    number: 30,
    zone: 'dashboard',
    label: '관리자 AI 사용량 상시 게이지 · SVG 기준선 · [기준 수시 변경]',
    xPercent: 72.0,
    yPercent: 26.0,
    summary: '일일 호출·토큰·월간 예산 상시 게이지 + 기준 대비 비교표 + SVG 기준선 & 실시간 [기준 수시 변경]',
    detail: 'AdminLlmUsageCompactCard에서 일일 호출·토큰·월간 비용 예산 게이지와 기준 대비 비교표, 일별 호출 SVG 기준선을 상시 확인하고 [기준 수시 변경](PUT /api/admin/llm-usage/limits)으로 절약형/표준/대용량 프리셋과 Groq 병렬 슬롯·고급 분석 모드를 즉시 전환합니다.',
    backendEndpoint: 'PUT /api/admin/llm-usage/limits',
    tip: '저장 즉시 서버 라우팅 한도와 병렬 슬롯 설정에 반영되어 비용과 처리량을 유연하게 제어합니다.',
    addedInVersion: '1.2.8',
  },
  {
    id: 'icon-pane-resize-splitter',
    number: 31,
    zone: 'maillist',
    label: '4개 구역 경계선 가로 드래그 크기 조절 · 더블클릭 초기화',
    xPercent: 31.0,
    yPercent: 54.0,
    summary: '폴더 창·메일 목록·읽기 창·AI 분석/첨부 도크 경계선 드래그 폭 조절 · 더블클릭 초기화 · 자동 저장(bmail_pane_widths_v1)',
    detail: 'Outlook 워크스페이스 4개 구역(폴더 창·메일 목록·읽기 창·AI 분석/첨부 도크) 사이의 경계선을 마우스로 좌우 드래그해 원하는 너비로 즉시 조절할 수 있으며, 경계선을 더블클릭하면 기본 너비로 초기화되고 변경된 폭은 브라우저(bmail_pane_widths_v1)에 자동 저장됩니다.',
    backendEndpoint: 'localStorage: bmail_pane_widths_v1',
    tip: '읽기 창이나 메일 목록 폭을 업무 스타일에 맞게 한 번 맞춰두면 다음 실행 시에도 그대로 복원됩니다.',
    addedInVersion: '1.2.10',
  },
  {
    id: 'icon-header-outlook-companion',
    number: 32,
    zone: 'header',
    label: 'Outlook 동반 검증 (기본 시작 · 사이드카 모드)',
    xPercent: 70.5,
    yPercent: 4.5,
    summary: 'v1.3.1 기본 시작 화면 · 300ms 자동 보내기/받기 · Groq AI/Claude 실제 LLM 검증 · MailItem.Display() 위임',
    detail: 'v1.3.1부터 기본 시작 화면(companion)으로 즉시 열리며, 300ms 자동 [보내기/받기](sendAndReceive: true)와 ActiveInspector/ActiveExplorer 2.5초 실시간 감지로 발송 전 첨부 누락·수신자/제목 누락·외부 민감정보 유출·도면/기한을 검증합니다. Groq 5-키 풀 및 [Claude 정밀 재분석]을 지원하고 회신은 항상 오리지널 Outlook 작성창(MailItem.Display())으로 안전 위임합니다.',
    backendEndpoint: 'POST /api/local/outlook/verify-item',
    tip: '우측 사이드카 도킹·항상 위(bmail:setCompanionDock)를 켜두면 클래식 Outlook 옆에서 발송 직전 위험을 즉시 차단합니다.',
    addedInVersion: '1.3.0',
  },
  {
    id: 'icon-header-send-receive',
    number: 33,
    zone: 'header',
    label: '보내기/받기 원클릭 동기화 (btn-send-receive-refresh)',
    xPercent: 61.5,
    yPercent: 4.5,
    summary: '상단 헤더·동반 검증 패널 [보내기/받기] 원클릭 실행 + 앱 시작 300ms 자동 동기화(sendAndReceive: true)',
    detail: '상단 헤더(btn-send-receive-refresh)와 동반 검증 패널 상단의 [보내기/받기] 버튼을 클릭하면 클래식 Outlook의 보내기/받기를 실행하고 새로 도착한 메일을 즉시 탐색·동기화합니다. 앱 실행 직후(300ms)에도 자동으로 최초 동기화가 수행됩니다.',
    backendEndpoint: 'POST /api/local/outlook/sync-incremental (sendAndReceive: true)',
    tip: '새 메일이 안 보일 때 Outlook 창으로 이동하지 않고 Bmail 상단 [보내기/받기] 한 번으로 즉시 가져옵니다.',
    addedInVersion: '1.3.1',
  },
  {
    id: 'icon-aidock-claude-deep',
    number: 34,
    zone: 'ai-dock',
    label: '[Claude 정밀] · 실제 LLM 자동 분석 (Groq AI / Claude)',
    xPercent: 79.5,
    yPercent: 25.0,
    summary: '기본 Groq 5-키 풀 자동 분석 + 긴급/선급/도면·[Claude 정밀] 클릭 시 Claude 5.5 에스컬레이션 + 오프라인 배너 [Claude API 설정 열기]',
    detail: 'MailAiDockPanel의 [Claude 정밀] 및 동반 검증 패널의 [Claude 정밀 재분석] 버튼으로 고난도 메일을 Claude 5.5로 즉시 승격 분석합니다. 키 미설정/오프라인 배너에서 [Claude API 설정 열기] 원클릭 이동을 제공하며, Groq AI / Claude 엔진 배지를 정확히 표기하고 elcon0955@naver.com 등 이메일 주소 도면 오탐도 교정되었습니다.',
    backendEndpoint: 'POST /api/local/mail/assistant',
    tip: '키 미설정 시 오프라인 배너의 [Claude API 설정 열기]를 누르면 Claude API 직접 설정 카드로 바로 이동합니다.',
    addedInVersion: '1.3.1',
  },
  {
    id: 'icon-claude-api-settings',
    number: 35,
    zone: 'header',
    label: 'Claude API 직접 설정 · 즉시 연결 테스트',
    xPercent: 93.8,
    yPercent: 4.5,
    summary: '⚙️ Outlook 설정 → AI · Claude API 탭 & 🛡️ 설정(ConfigEnvManager) 최상단 Claude API 직접 설정 카드',
    detail: 'Groq 연동 없이도 ANTHROPIC_API_KEY(sk-ant-...)를 직접 입력·.env 저장하고 즉시 연결 테스트(POST /api/local/ai/claude-test)를 실행합니다. 기본 엔진 모드(Claude API 우선 / 자동 혼합 / Groq 전용)와 모델(자동 난이도 / Haiku 5.5 / Sonnet 5.5 / Opus 5.5) 선택(GET/PUT /api/local/ai/claude-config)을 지원하며, AI 분석 도크·동반 검증 패널 오프라인 배너에서 [Claude API 설정 열기]로 1클릭 이동합니다.',
    backendEndpoint: 'GET/PUT /api/local/ai/claude-config · POST /api/local/ai/claude-test',
    tip: 'Groq 키가 없어도 sk-ant-... 키를 저장하고 [즉시 연결 테스트]를 통과하면 Claude API 우선 모드로 바로 동작합니다.',
    addedInVersion: '1.3.1',
  },
];

export const MANUAL_MOTION_SCENES: ManualMotionScene[] = [
  {
    id: 'scene-01-outlook-tree',
    step: 1,
    title: 'Classic Outlook 즐겨찾기 · 저장소 트리 완벽 동기화',
    subtitle: 'New2026.xml 즐겨찾기와 김봉정·백업 다단계 폴더 및 미읽음 카운트 재현',
    narration:
      'Outlook 로컬 설정(New2026.xml 및 레지스트리)을 자동 인식하여 즐겨찾기(선보, 유일, 가스앤텍)와 김봉정·백업 저장소의 모든 하위 폴더, 파란색 미읽음 숫자를 그대로 표시합니다.',
    durationSec: 6,
    targetIconId: 'icon-tree-stores',
    zone: 'folders',
    cursor: { x: 8.5, y: 42.0 },
    spotlight: { x: 1.5, y: 15.0, w: 15.0, h: 76.0 },
    previewState: 'outlook-overview',
    badgeText: 'STEP 1 · 폴더·즐겨찾기 동기화',
  },
  {
    id: 'scene-02-incremental-sync',
    step: 2,
    title: '1회 불러오기 영구 저장 + 신규 메일만 초고속 증분 업데이트',
    subtitle: '24,626건 재색인 없이 PST 워터마크 비교 & 실시간 Outlook COM 연동',
    narration:
      '한 번 불러온 메일은 로컬 SQLite에 영구 보관됩니다. [신규 메일 증분 동기화] 클릭 시 변경된 PST 노와 현재 열려 있는 Outlook의 오늘자 신규 메일만 1~2초 만에 업데이트합니다.',
    durationSec: 6,
    targetIconId: 'icon-ribbon-live-sync',
    zone: 'ribbon',
    cursor: { x: 32.5, y: 11.2 },
    spotlight: { x: 25.0, y: 8.0, w: 18.0, h: 7.0 },
    previewState: 'live-sync',
    badgeText: 'STEP 2 · 증분 델타 업데이트',
  },
  {
    id: 'scene-03-rich-html-cid',
    step: 3,
    title: '원본 HTML 서식 · 인라인(cid:) 3D 도면 캡처 완벽 렌더링',
    subtitle: '본문 내 삽입 이미지 선명 표시 + 서명 로고(image001.png) 자동 분리',
    narration:
      '메일 본문의 표, 하이라이트, cid: 인라인 3D 모델링 캡처 이미지를 원본 그대로 렌더링하며, 서명용 작은 로고 이미지는 첨부파일 칩에서 자동으로 숨겨 실첨부 문서만 깔끔하게 보여줍니다.',
    durationSec: 6,
    targetIconId: 'icon-reading-rich-html',
    zone: 'reading',
    cursor: { x: 46.0, y: 48.0 },
    spotlight: { x: 31.0, y: 16.0, w: 33.0, h: 75.0 },
    previewState: 'rich-reading',
    badgeText: 'STEP 3 · HTML + 인라인 이미지',
  },
  {
    id: 'scene-04-ai-dock-5tabs',
    step: 4,
    title: '우측 세로 1분할: AI 분석 · 3종 답안 · 대화 · 번역 · 액션 도크',
    subtitle: 'Claude 5.5 (Haiku → Sonnet → Opus) 난이도 자동 라우팅 적용',
    narration:
      '본문 바로 옆 첫 번째 세로 분할창에서 메일 난이도에 따라 Claude 5.5 모델을 자동 선택하여 핵심 요약, 3가지 톤의 회신 답안, 다중 턴 대화, 정밀 번역, 액션 상태를 즉시 처리합니다.',
    durationSec: 7,
    targetIconId: 'icon-aidock-reply',
    zone: 'ai-dock',
    cursor: { x: 71.5, y: 21.0 },
    spotlight: { x: 64.5, y: 16.0, w: 17.0, h: 75.0 },
    previewState: 'ai-dock-reply',
    badgeText: 'STEP 4 · AI 분석·답안 도크',
  },
  {
    id: 'scene-05-attach-preview-ai',
    step: 5,
    title: '우측 세로 2분할: 첨부파일 즉시 미리보기 & AI 문서 사양 분석',
    subtitle: 'PDF 도면·POS 내장 뷰어 + Excel/Word 구조 파싱 + AI 핵심 스펙 추출',
    narration:
      '두 번째 세로 분할창에서는 첨부파일을 클릭하는 즉시 PDF·이미지·엑셀·워드 내용을 미리 보여주고, [AI 첨부 문서 분석]으로 도면 리비전과 POS 편차 항목을 자동 요약합니다.',
    durationSec: 6,
    targetIconId: 'icon-attachdock-preview',
    zone: 'attach-dock',
    cursor: { x: 86.0, y: 21.0 },
    spotlight: { x: 81.8, y: 16.0, w: 17.0, h: 75.0 },
    previewState: 'attach-preview',
    badgeText: 'STEP 5 · 첨부 미리보기·분석',
  },
  {
    id: 'scene-06-batch-folder-export',
    step: 6,
    title: '지정 폴더 첨부파일 일괄 저장 & MD·CSV·HTML 3종 리포트 작성',
    subtitle: '기간·발신자·프로젝트·전체/일부 필터링 + 폴더 자동 분류 + 리포트 생성',
    narration:
      '[폴더 저장 · 리포트] 탭에서 저장할 로컬 폴더를 지정하고 기간·보낸사람·프로젝트·파일유형 옵션을 선택하면, 첨부파일을 폴더 구조별로 깔끔하게 저장하고 엑셀 호환 CSV·MD·HTML 리포트를 동시에 작성합니다.',
    durationSec: 7,
    targetIconId: 'icon-attachdock-batch-export',
    zone: 'attach-dock',
    cursor: { x: 93.5, y: 21.0 },
    spotlight: { x: 81.8, y: 16.0, w: 17.0, h: 75.0 },
    previewState: 'batch-folder-export',
    badgeText: 'STEP 6 · 폴더 일괄저장·리포트',
  },
  {
    id: 'scene-07-executive-dashboard',
    step: 7,
    title: '종합 분석글 · 관리자 사용량 상시 게이지 · [기준 수시 변경]',
    subtitle: '일일 호출·토큰·월간 예산 게이지 + 비교표 + SVG 기준선 & PUT /api/admin/llm-usage/limits',
    narration:
      '대시보드에서는 주간·월간 서술형 분석글과 미회신 현황은 물론, 관리자 LLM 사용량 상시 게이지·기준 대비 비교표·일별 호출 SVG 기준선을 확인하고 [기준 수시 변경]으로 프리셋과 병렬 슬롯을 즉시 전환합니다.',
    durationSec: 7,
    targetIconId: 'icon-dashboard-usage-limits',
    zone: 'dashboard',
    cursor: { x: 72.0, y: 26.0 },
    spotlight: { x: 15.0, y: 14.0, w: 72.0, h: 76.0 },
    previewState: 'executive-dashboard',
    badgeText: 'STEP 7 · 분석 대시보드 & 사용량 게이지',
  },
  {
    id: 'scene-08-sgw-and-mini-assist',
    step: 8,
    title: '최근 14일 미확인 자동 알림 토스트([30분 뒤] · [×] 재알림) & SGW',
    subtitle: '2023 장기 과거 제외(maxAgeDays=14) + 우측 하단 비간섭 자동 알림 토스트 & 타이머',
    narration:
      '최근 14일 이내 미확인 수신 메일이 감지되면 우측 하단 비간섭 토스트 팝업이 자동 표시됩니다. [30분 뒤] 다시 알림 타이머와 [×] 닫기 후 새 메일 도착 시 자동 재알림을 지원하며 기존 2023 장기 과거 메일 제외 정책을 유지합니다.',
    durationSec: 6,
    targetIconId: 'icon-floating-unread-briefing-pill',
    zone: 'mini-assist',
    cursor: { x: 85.0, y: 88.0 },
    spotlight: { x: 76.0, y: 65.0, w: 22.5, h: 31.0 },
    previewState: 'sgw-login-chat',
    badgeText: 'STEP 8 · 14일 미확인 자동 알림',
  },
  {
    id: 'scene-09-classic-3pane-and-ai-toggles',
    step: 9,
    title: '오리지널 클래식 Outlook 평면 1px 3-Pane & [AI 요약 보기] 토글',
    subtitle: '무채색 평면 리본 · 아바타 기본 OFF · 읽기 창 AI 요약 기본 숨김 + 도크 0ms 수동/자동 토글',
    narration:
      'v1.2.9는 무채색 평면 리본, 둥근 박스 제거, 발신자 원형 아바타 기본 OFF, 하위 폴더 아이콘 간소화로 오리지널 Outlook 밀도를 복원하고, 읽기 창 [AI 요약 보기] 토글과 AI 도크 0ms 즉시 요약·수동/자동 스위치를 제공합니다.',
    durationSec: 7,
    targetIconId: 'icon-reading-ai-summary-toggle',
    zone: 'reading',
    cursor: { x: 59.0, y: 21.0 },
    spotlight: { x: 31.0, y: 15.0, w: 50.0, h: 75.0 },
    previewState: 'rich-reading',
    badgeText: 'STEP 9 · 클래식 평면 3-Pane & 토글',
  },
  {
    id: 'scene-10-sqlite-speed-and-aidock-instant',
    step: 10,
    title: 'SQLite 복합 인덱스 고속화 · 비블로킹 부트스트랩 · 0ms 즉시 요약',
    subtitle: 'idx_emails_active_ts / idx_emails_unread_inbox · /api/local/bootstrap 비블로킹 탐색',
    narration:
      'SQLite 복합 인덱스(idx_emails_active_ts, idx_emails_unread_inbox)와 /api/local/bootstrap 비블로킹 설정 탐색으로 초기 로딩과 폴더 전환을 고속화하고, AI 도크의 [수동/자동] 토글과 메일 클릭 0ms 즉시 요약을 제공합니다.',
    durationSec: 6,
    targetIconId: 'icon-aidock-instant-manual-toggle',
    zone: 'ai-dock',
    cursor: { x: 79.5, y: 17.0 },
    spotlight: { x: 64.5, y: 15.0, w: 17.0, h: 76.0 },
    previewState: 'outlook-overview',
    badgeText: 'STEP 10 · DB 인덱스 고속화 & 0ms 요약',
  },
  {
    id: 'scene-11-pane-resize-and-auto-toast',
    step: 11,
    title: '4개 구역 경계선 가로 드래그 크기 조절 · 더블클릭 초기화',
    subtitle: '폴더·메일목록·읽기창·AI/첨부 도크 폭 자유 조절 & 브라우저 자동 저장(bmail_pane_widths_v1)',
    narration:
      '폴더 창·메일 목록·읽기 창·AI 분석/첨부 도크 사이의 경계선을 마우스로 좌우 드래그해 원하는 크기로 즉시 조절하고, 더블클릭 시 기본 너비로 초기화되며 변경 폭은 bmail_pane_widths_v1에 자동 저장됩니다.',
    durationSec: 6,
    targetIconId: 'icon-pane-resize-splitter',
    zone: 'maillist',
    cursor: { x: 31.0, y: 54.0 },
    spotlight: { x: 14.5, y: 15.0, w: 50.0, h: 76.0 },
    previewState: 'outlook-overview',
    badgeText: 'STEP 11 · 경계선 드래그 폭 조절',
  },
  {
    id: 'scene-12-outlook-companion-sidecar',
    step: 12,
    title: '기본 시작: Outlook 동반 검증 · MailItem.Display() 안전 위임',
    subtitle: 'v1.3.1 기본 시작 화면 · 300ms 자동 보내기/받기 · Groq AI/Claude 검증 & 오리지널 작성창 위임',
    narration:
      'v1.3.1부터 앱 실행 즉시 Outlook 동반 검증(companion) 모드로 시작하며 300ms 자동 보내기/받기를 수행합니다. 클래식 Outlook 옆에서 2.5초마다 메일을 감지·검증하고 회신은 항상 MailItem.Display()로 오리지널 작성창에 안전 위임합니다.',
    durationSec: 7,
    targetIconId: 'icon-header-outlook-companion',
    zone: 'header',
    cursor: { x: 70.5, y: 4.5 },
    spotlight: { x: 63.0, y: 1.5, w: 36.0, h: 95.0 },
    previewState: 'outlook-companion',
    badgeText: 'STEP 12 · 기본 시작: 동반 검증',
  },
  {
    id: 'scene-13-send-receive-and-claude-deep',
    step: 13,
    title: '[보내기/받기] 원클릭 & [Claude 정밀] 실제 LLM 분석 · 보안 강화',
    subtitle: 'btn-send-receive-refresh · Groq AI/Claude 배지 · 비연동 5메뉴 정리 · 127.0.0.1 로컬 보안',
    narration:
      '상단 [보내기/받기](btn-send-receive-refresh)로 새 메일을 즉시 수신하고, AI 도크의 [Claude 정밀]과 Groq AI/Claude 배지로 실제 LLM 분석을 확인합니다. 비연동 5개 메뉴가 정리되었고 로컬 서버는 127.0.0.1 바인딩과 Origin/경로 검증으로 안전하게 보호됩니다.',
    durationSec: 7,
    targetIconId: 'icon-header-send-receive',
    zone: 'header',
    cursor: { x: 61.5, y: 4.5 },
    spotlight: { x: 56.0, y: 1.5, w: 42.0, h: 28.0 },
    previewState: 'live-sync',
    badgeText: 'STEP 13 · 보내기/받기 & Claude 정밀',
  },
];

export const MANUAL_CHANGELOG: ManualChangelogEntry[] = [
  {
    version: '1.3.1',
    date: '2026-11-27',
    headline:
      '기본 시작 Outlook 동반 검증 · [보내기/받기] 원클릭 & 300ms 자동 동기화 · 실제 LLM 자동 분석(Groq AI/Claude 정밀) · 네비게이션 정리 & 로컬 보안 강화',
    highlights: [
      '기본 시작 화면을 Outlook 동반 검증(companion) 모드로 전환하고, 직접 SMTP/COM 발송 대신 항상 오리지널 Classic Outlook 작성창(MailItem.Display())으로 안전하게 위임',
      '상단 헤더 및 동반 검증 패널에 [보내기/받기](btn-send-receive-refresh) 원클릭 버튼 추가 + 앱 시작 즉시(300ms) Outlook 보내기/받기 및 새 메일 자동 탐색·동기화(POST /api/local/outlook/sync-incremental sendAndReceive: true)',
      '⚙️ Outlook 설정 → AI · Claude API 탭 및 🛡️ 설정(ConfigEnvManager) 최상단에 Claude API 직접 설정 카드 추가(GET/PUT /api/local/ai/claude-config, POST /api/local/ai/claude-test): Groq 연동 없이도 ANTHROPIC_API_KEY(sk-ant-...) 직접 입력·.env 저장·즉시 연결 테스트·기본 엔진 모드(Claude API 우선/자동 혼합/Groq 전용)·모델 선택(자동 난이도/Haiku 5.5/Sonnet 5.5/Opus 5.5) 및 AI 도크·동반 검증 오프라인 배너 [Claude API 설정 열기] 원클릭 이동 지원',
      '메일 AI 분석 도크(MailAiDockPanel) 및 동반 검증 패널의 실제 LLM 자동 분석 활성화(기본 Groq 5-키 풀, 긴급/선급/도면 메일 및 [Claude 정밀] 버튼 시 Claude 5.5 자동 에스컬레이션, Groq AI/Claude 엔진 배지 정확 표기, 이메일 주소 elcon0955@naver.com 도면 오탐 수정, 상투적 오프라인 문구 제거)',
      'Outlook 비연동/중복 5개 메뉴(일정·달력, 할 일·To-Do, 기록함, AI 브리핑·Ask AI, 플러그인 설치)를 네비게이션에서 정리(숨김)',
      '로컬 서버 보안 강화(127.0.0.1 바인딩, /api/local/*·/api/admin/* 루프백/Origin 검증, /api/local/files/parse-paths 민감 파일 경로 차단) 및 PUT /api/local/state/:key SQLite 저장 404 수정',
    ],
  },
  {
    version: '1.3.0',
    date: '2026-11-26',
    headline:
      'Outlook 동반 검증 사이드카 모드(2.5s 실시간 감지 · 발송 전 위험 검증 · MailItem.Display() 원창 열기 · Electron 우측 도킹)',
    highlights: [
      'Outlook 동반 검증 사이드카 모드 추가: 아웃룩을 완전 대체하지 않고 클래식 Outlook 옆에 붙어 ActiveInspector/ActiveExplorer 선택·작성 메일을 2.5s 실시간 감지',
      '발송 전 실시간 위험 검증(POST /api/local/outlook/verify-item): 첨부 누락·수신자/제목 누락·외부 수신자 민감정보 유출·조선 도면/기한 자동 추출',
      '회신 초안 오리지널 Outlook 작성창 직접 열기(MailItem.Display(), POST /api/local/outlook/send-via-com action: display) 및 Electron 우측 사이드카 도킹·항상 위(bmail:setCompanionDock) 지원',
    ],
  },
  {
    version: '1.2.10',
    date: '2026-11-25',
    headline:
      '4개 구역 경계선 마우스 가로 드래그 크기 조절(bmail_pane_widths_v1) + 최근 14일 미확인 수신 메일 자동 알림 토스트([30분 뒤] · [×] 재알림)',
    highlights: [
      'Outlook 워크스페이스 4개 구역(폴더 창·메일 목록·읽기 창·AI 분석/첨부 도크) 경계선 마우스 가로 드래그 크기 조절 · 더블클릭 기본 너비 초기화 · 브라우저 자동 저장(bmail_pane_widths_v1)',
      '최근 14일 이내 미확인 수신 메일 자동 알림 팝업 발생(우측 하단 비간섭 토스트, [30분 뒤] 다시 알림 타이머, [×] 닫기 후 새 미확인 수신 시 자동 재알림, 기존 2023 장기 과거 메일 제외 유지)',
    ],
  },
  {
    version: '1.2.9',
    date: '2026-11-25',
    headline:
      '오리지널 클래식 Outlook 평면 1px 3-Pane 정리 + 최근 14일 미확인 브리핑 플로팅 팝업 + SQLite 복합 인덱스·0ms 즉시 요약 고속화',
    highlights: [
      '오리지널 클래식 Outlook 평면 1px 3-Pane 정리: 무채색 평면 리본, 둥근 박스 제거, 발신자 원형 아바타 기본 OFF, 하위 폴더 아이콘 간소화, 읽기 창 인라인 AI 요약 기본 숨김 및 [AI 요약 보기] 토글 추가',
      '최근 14일(maxAgeDays=14) 미확인 메일 브리핑 기준 적용(4년 전 등 장기 과거 미열람 메일 자동 제외) + 우측 하단 플로팅 알약/사이드 팝업 배치 및 리본 [미확인 브리핑] 토글 연동',
      '속도 고속화: SQLite 복합 인덱스(idx_emails_active_ts / idx_emails_unread_inbox) 적용, /api/local/bootstrap 비블로킹 설정 탐색, AI 도크 메일 클릭 0ms 즉시 요약 및 [수동/자동] 토글 탑재',
    ],
  },
  {
    version: '1.2.8',
    date: '2026-11-25',
    headline:
      '관리자 LLM 사용량 기준 게이지·비교 도표·SVG 기준선 상시 표시 & [기준 수시 변경] 실시간 저장·라우팅 반영',
    highlights: [
      '관리자 사용량 기준 게이지(일일 호출·토큰·월간 비용 예산) 및 기준 대비 비교 도표를 대시보드에 상시 표시',
      '일별·프로바이더별 API 호출량 그래프에 일별 호출 한도 SVG 기준선 추가로 초과 위험 즉시 가시화',
      '[기준 수시 변경](PUT /api/admin/llm-usage/limits) UI를 통해 절약형·표준·대용량 프리셋과 Groq 병렬 슬롯·고급 분석 모드를 실시간 저장하고 즉시 AI 라우팅에 반영',
    ],
  },
  {
    version: '1.2.7',
    date: '2026-11-25',
    headline:
      'Classic Outlook 리본 작성·발송 100% 패리티 + Groq 5-키 병렬 배치·Claude 고급 분석 분리 + 대시보드 API 호출량 그래프 & npm run groq:keys',
    highlights: [
      'npm run groq:keys 명령으로 쉼표 구분 5개 Groq API 키를 터미널에서 한 번에 일괄 설정하고 즉시 라운드로빈 풀 반영',
      '고속·병렬 배치 작업은 Groq 5-키 풀로, 고급·심층 분석은 Claude 5.5로 분리하는 이원화 AI 라우팅 엔진 적용',
      '종합 대시보드에 API 라우팅 정책 배지(고속·병렬 Groq 5-키 vs 고급·심층 Claude 5.5) 및 일별·프로바이더별 API 호출량 그래프(AdminLlmUsageCompactCard) 탑재',
      'Classic Outlook 리본 작성·발송 100% 패리티 완성: 메시지/삽입/옵션 3탭 리본, 주소록 152명 선택기·자동완성, 숨은 참조(Bcc), 파일 첨부+드래그앤드롭, 읽음 확인 요청, 발송 지연 취소 카운트다운, 명함 HTML 서명 및 multipart/mixed+cid: SMTP/PC Outlook COM 발송',
    ],
  },
  {
    version: '1.2.6',
    date: '2026-11-25',
    headline:
      '비동기 증분 동기화 · 미확인 긴급 브리핑(UnreadBriefingBanner) · 우측 AI 도크 상시 고정 · 독립 HTML 새 창 & 명함 cid: 발송',
    highlights: [
      '비동기 증분 동기화(pollLiveOutlookInboxAsync)로 대시보드·UI 지연 원천 해소',
      '미확인 메일 읽지않음 유지 전수 분석 · 프로젝트별 중요도순 브리핑 & 반드시 긴급회신 알림(UnreadBriefingBanner)',
      '오른쪽 세로분할 AI 분석 도크(272~296px) 상시 고정 및 첨부파일 클릭 우측 미리보기 + 메일 명함·네임태그 인라인 이미지(image.png) 자동 제외',
      'PC Outlook 전체 설정 탐색·적용(아웃룩 설정 탐색·적용) 및 명함 HTML 로고 이미지 base64 자동 인라인 반영',
      'Outlook HTML 답장·전달·새 메일 독립 새 창 작성(OutlookComposeWindow) 및 cid: 인라인 명함 SMTP/COM 발송',
      'groq-keys.txt 간편 5-키 입력 파일 지원 및 매뉴얼 전체 목록 표시 개선',
    ],
  },
  {
    version: '1.2.5',
    date: '2026-11-24',
    headline:
      '항상 로그인 유지(AES-256-GCM 세션) + 백그라운드 선행 정리 + 미열람 시 SGW 채팅(모바일) 요약 알림',
    highlights: [
      'POP3/IMAP(pop3-sessions.enc.json) 및 Microsoft OAuth(ms-sessions.enc.json) 세션을 AES-256-GCM 암호화 저장하고 앱 재시작·포트 변경 시 자동 복원 + 리프레시 토큰 자동 연장',
      '백그라운드 상시 데몬(BackgroundMailMonitor)이 앱을 보지 않아도 60초마다 메일을 먼저 수신·요약하고 대시보드 미회신 요약을 자동 선행 정리',
      '지정 시간 동안 메일을 열람하지 않으면 SGW 사내 채팅(내 계정 쪽지)으로 요약·도면·미회신 건수를 즉시 전송해 핸드폰 알림 수신 지원 + 트레이 미니창(380×168) 동시 표시',
    ],
  },
  {
    version: '1.2.4',
    date: '2026-11-24',
    headline:
      'Groq 5-키 순환 기본 엔진(openai/gpt-oss-120b 스트리밍) + Claude 고성능·장애 에스컬레이션 & Outlook 기능·설정 완전 복제',
    highlights: [
      '기본 AI 엔진을 Groq openai/gpt-oss-120b 스트리밍 5-키 라운드로빈 풀(GROQ_API_KEY~5)로 전환하고, 장애·고난도 분석 시 Claude(ANTHROPIC_API_KEY)로 자동 전환',
      'Outlook 사람(주소록 · vCard/CSV) 및 할 일(To-Do · 플래그 메일 자동 연동) 모듈 추가',
      '검색 폴더(안읽음·플래그·첨부·대용량·다시 알림), 빠른 단계(Quick Steps), 색 범주 태그, 발신자 일괄 정리(Sweep), 다시 알림(Snooze) 구현',
      'HKCU 레지스트리(근무시간·싱글라인 리본·미니멀 헤더) 이관 및 읽음 지연/삭제 후 이동/기본 중요도/읽음 확인 요청 설정 완비',
    ],
  },
  {
    version: '1.2.3',
    date: '2026-11-24',
    headline:
      '초경량 코드 스플리팅(React.lazy) · 심플 3단 기본 화면 · 앱스타일 스퀘어클 아이콘 · 무간섭 100% 자동 이관 & 상시 증분 누적',
    highlights: [
      '초기 번들 17개 화면/도크 React.lazy 분할 및 초기 부트스트랩 100건 경량화로 즉각 기동(전체 2.4만+ 메일은 SQLite FTS5 즉시 검색 유지)',
      '기본 화면을 3단 클래식 Outlook 구조로 심플화(AI 분석 도크·첨부 미리보기·요약 펼침 1-클릭 전환)',
      '앱스타일 스퀘어클 그라데이션 아이콘(favicon.svg · 256px icon.png · 다중 DIB icon.ico) 헤더/작업표시줄/설치본 일괄 적용',
      '설치·실행 즉시 로컬 Outlook 설정·PST·메일 100% 자동 이관 + 백그라운드 3분/창복귀 무음 증분 누적 유지',
    ],
  },
  {
    version: '1.2.2',
    date: '2026-11-24',
    headline:
      '관리자 LLM 사용량 소형 모니터 + .ics 양방향 달력 + 스티커 메모 + 지연 발송 취소 큐 + .msg 직송 파서 + 일일 SQLite 백업 + 트레이 1-클릭 알림 완성',
    highlights: [
      '관리자 설정 페이지 소형 LLM 사용량 모니터 바(호출/토큰/비용/지연 + 모델별 상세) 및 일 1회 SQLite gzip 자동 백업 스냅샷(최대 7세대·SHA-256 검증) 탑재',
      'Outlook 일정표 Windows .ics 양방향 내보내기·가져오기 어댑터(RFC 5545) 구현',
      '상단 헤더·플로팅 스티커 메모 위젯(Sticky Notes: 4색 스와치·상단 고정·현재 메일 메모 1클릭 연결) 추가',
      '메일 작성 창 발송 지연 대기 시간 선택(0초/10초/30초/1분/5분) 및 좌하단 카운트다운 발송 취소/즉시 발송 큐 구현',
      '순수 Node.js OLE2/CFBF 바이너리 .msg 파서 및 워크벤치 드래그앤드롭(.msg/.eml/.pst) 즉시 파싱 지원',
      'Windows 트레이 우하단 미니 메일 알림창 1-클릭 [즉시 읽음] · [즉시 회신] · [열기] 액션 연결',
    ],
  },
  {
    version: '1.2.1',
    date: '2026-11-23',
    headline:
      'Classic Outlook 100% 패리티 · 우측 2분할 AI/첨부 도크 · 주간/월간 분석글 대시보드 · SGW 공식 로그인 창 연동 · 모션그래픽 매뉴얼',
    highlights: [
      'Claude 5.5 3단계 난이도 자동 라우팅 탑재 (기본 claude-haiku-5-5 → 중난도 claude-sonnet-5-5 → 고난도 claude-opus-5-5, Anthropic 15개 전 모델 라이브 검증 통과)',
      '본문 우측 세로 2분할 도크 구현: [1분할 AI 분석·답안·대화·번역·액션] + [2분할 첨부 미리보기·AI 문서분석 및 지정 폴더 일괄 저장·3종 리포트 생성]',
      '종합 대시보드 업그레이드: 서술형 종합 분석글 + 주간/월간 보고서 테이블 + 내게 온 미회신건 자동 정리 + 프로젝트/주제/기간별 받은·보낸 메일 및 첨부 통합 조회',
      '1회 불러오기 SQLite 영구 저장 + PST 워터마크 비교 & 실시간 Outlook COM 델타 증분 업데이트 구현',
      'Classic Outlook 즐겨찾기(New2026.xml), 김봉정/백업 다단계 폴더 트리, 파란색 미읽음 카운트, 본문 HTML + cid: 인라인 이미지 렌더링, 서명 이미지 분리 및 도면: MESSAGE 오탐 수정',
      'SGW 실시간 채팅: 실제 SGW 로그인 창(https://sgw.seastar.work) 팝업 완전 로그인 세션 연동 구현',
      '우측 하단 소형 고정(330px) 무오류 번역·요약 미니창 및 개발용 상단 버튼(HANDOFF/전체 코드 다운로드) 완전 제거',
      '화면 1:1 재현 인터랙티브 HTML 팝업 매뉴얼(모션그래픽 동영상 플레이어 + 25개 아이콘 탐색기) 및 bmail.seastar.work 20챕터 공개 설치 사이트 구축',
    ],
  },
  {
    version: '1.2.0',
    date: '2026-11-22',
    headline: '오프라인 결정론적 폴백 엔진 및 로컬 SQLite FTS5 아키텍처 강화',
    highlights: [
      '클라우드 API 키 미설정 시 업무일지·회의록·문서분석 100% 로컬 결정론적 엔진 자동 폴백',
      'Windows 독립 실행형 Electron + Node 24 내장 SQLite 메일함 인덱서 안정화',
    ],
  },
];
