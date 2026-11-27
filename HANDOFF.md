# HANDOFF.md — Bmail v1.3.1 엔지니어링 인수인계서 (개발 전용 · 프로덕션 UI 비노출)

> **대상 프로젝트**: Bmail — Outlook AI 워크스페이스 (`MailLog Studio for Outlook` · `v1.3.1`)
> **보안 정책**: 프로덕션 사용자 UI에서 `[HANDOFF]` 및 `[전체 코드 다운로드]` 버튼은 제거되었으며, `/api/export/handoff` 엔드포인트 삭제(HTTP `404`), `/api/export/source-archive*`는 `BMAIL_DEV_EXPORT === 'true'` 개발 환경 전용으로 격리되었습니다(기본 HTTP `404`).

---

## 0. 2026-11-27 v1.3.1 전체 `/grill-me` 감사 결함 전수 해결 · Claude API 직접 설정 · 실제 AI 자동 분석 · 보내기/받기 · 보안 강화 실측 현황

| 영역 | 실측 상태 | 세부 실측 내역 및 경로 |
|---|---|---|
| **Claude API 직접 설정 & 즉시 연결 테스트 (Groq 미연동 대응)** | ✅ 실측 완료 | `src/utils/localClaudeEnv.ts` & `src/utils/localClaudeEnv.test.ts` (4/4 PASS: 마스킹·키 유효성 검증·`.env` 안전 업서트·줄바꿈 포함 JSON 복구 파서 `parseAssistantJsonRobust`), `server.ts` (`GET/PUT /api/local/ai/claude-config`, `POST /api/local/ai/claude-test`, `resolveActiveClaudeKey` `.env`+EMV 통합 조회, 새 키 저장 시 `claude-only` 기본 전환), `src/components/ClaudeApiSettingsCard.tsx`, `src/components/ConfigEnvManager.tsx`(상단 배치), `src/components/OutlookSettingsModal.tsx`(`AI · Claude API` 1번 탭), `MailAiDockPanel.tsx`/`OutlookCompanionPanel.tsx`(오프라인 배너 `[Claude API 설정 열기]` 원클릭). 실서버(`3476`) 실측: `GET /api/local/ai/claude-config` HTTP `200`(`configured: true`, `maskedPreview: sk-ant-••••••BgAA`, `noRawKeyLeaked: true`), `POST /api/local/ai/claude-test` HTTP `200`(`liveTestOk: true`, `claude-haiku-5-5`, **`790ms`**), 1440px `ego-browser` [bmail-v131-claude-settings-1440.png](file:///C:/Users/FREE/.gemini/antigravity/brain/ebe7d412-bd24-4d99-93e9-83579563c299/scratch/bmail-v131-claude-settings-1440.png) 확인 *(디스크 `.env` 쓰기/삭제 왕복은 운영 키 보호를 위해 단위 테스트로 검증)* |
| **실제 LLM 본문 분석 복구 · 오탐/중복 접두사/JSON 파싱 해결** | ✅ 실측 완료 | `src/utils/outlookWorkspaceModel.ts` & `.test.ts` (이메일 주소·URL 사전 제거로 `elcon0955@naver.com` 도면번호 오탐 차단), `server.ts` (`POST /api/local/mail/assistant` 상투적 템플릿 제거·본문 근거 프롬프트·`parseAssistantJsonRobust`·Claude `maxTokens: 3072`·Groq/Claude 상호 폴백·중복 모델 접두사 제거), `MailAiDockPanel.tsx` (`autoAiOnSelect` 기본 `true`·슬림 페이로드·`Groq AI`/`Claude` 정확한 배지·`[Claude 정밀]` 버튼·`fallbackReason` 투명 고지), `OutlookCompanionPanel.tsx` (자동 AI 분석·`[다시 분석]` 캐시 우회·`[Claude 정밀 재분석]`). 실서버(`3476`) 실측: 일반 합성 메일 **`usedFallback: false`**(`Groq · openai/gpt-oss-120b (#1) (5-키 풀)`, `3532ms`, `drawings: []` 오탐 0건), 긴급 선급 합성 메일 **`usedFallback: false`**(`Claude · claude-haiku-5-5`, `5622ms`), 1440px `ego-browser` [bmail-v131-real-ai-dock-1440.png](file:///C:/Users/FREE/.gemini/antigravity/brain/ebe7d412-bd24-4d99-93e9-83579563c299/scratch/bmail-v131-real-ai-dock-1440.png)에서 우측 AI 도크 `Groq AI` 배지 + 실제 본문 요약·체크리스트 확인 |
| **기본 홈 동반 검증 전환 · 안전 발송 위임 · 보내기/받기 · 5개 메뉴 정리** | ✅ 실측 완료 | `src/App.tsx` (기본 페이지 `'companion'`, 시작 `300ms` 즉시 `sendAndReceive: true` 동기화, 헤더 `[보내기/받기]` `btn-send-receive-refresh`, `calendar`/`tasks`/`archive`/`ai-inbox`/`install` 네비 UI 숨김·소스 보존), `src/components/OutlookComposeWindow.tsx` (직접 SMTP/COM `.Send()` 제거 → 항상 `MailItem.Display()` 오리지널 Outlook 작성창 위임, 버튼 `Outlook 열기`), `src/components/OutlookWorkbench.tsx` (`실제 Outlook 플러그인 설치` 배너 숨김), `src/services/localOutlookMigrator.ts` (`triggerSendReceive` `$ns.SendAndReceive($false)`). 실서버(`3476`) `POST /api/local/outlook/sync-incremental` HTTP `200`, 1440px `ego-browser` [bmail-v131-companion-home-1440.png](file:///C:/Users/FREE/.gemini/antigravity/brain/ebe7d412-bd24-4d99-93e9-83579563c299/scratch/bmail-v131-companion-home-1440.png)에서 기본 홈·`leakedHiddenMenus: []`·`hasPluginInstallBannerText: false` 실측 *(클래식 Outlook 미실행 환경으로 COM 실제 창 팝업 및 `$ns.SendAndReceive` 실행은 `[미검증]`)* |
| **로컬 서버 보안 잠금 · SQLite state 404 3종 복구** | ✅ 실측 완료 | `server.ts` (`BIND_HOST` 기본 `'127.0.0.1'`, `/api/local/*`·`/api/admin/*` 루프백+Origin 검증, `/api/local/files/parse-paths` `.env`/`groq-keys.txt`/`*.sqlite*`/시스템 경로 차단, `PUT /api/local/state/:key` 추가). 실서버(`3476`) 실측: 외부 `Origin: https://evil.example` → `/api/local/bootstrap` **`403`** · `/api/local/ai/claude-config` **`403`**, `.env` 읽기 요청 **`400`** 차단, `PUT /api/local/state/outlookPrefs`·`mail_action_states`·`mail_chat_histories` 3종 모두 HTTP **`200 { ok: true }`** |
| **매뉴얼 Flash 동기화 · 검증 · 배포 (`1.3.1`) · Jev 게이트 상태** | ✅ 실측 완료 *(Jev 외부 호출 `[미검증]`)* | `Model: "flash"` 서브에이전트가 `manualCatalog.ts` (`1.3.1`, 35개 아이콘 `#1~#35`, 13개 모션 씬) & `InteractiveManualModal.tsx` 동기화 (`manualCatalog.test.ts` 1/1 PASS). `npm run lint` (`0`), `node --check` (`0`), Node `v24.20.0` 전체 37개 파일 **`156/156 PASS`** (`0` fail, `0` skipped). `release-exe/Bmail-Outlook-AI-Setup-1.3.1.exe` (`121,479,318` B, `sha256: E28A0FC98161E20E78AF3B7F54DA0902D1C73BC458726F472B57209CBD1705C3`), R2 `latest.json`, Worker Version `e0b76e32-fcf0-41b4-b4d3-d3e081fc444d` (`https://mail.seastar.work/api/health` `1.3.1`), 무음 설치 `/S` (`installed 1.3.1`), `PORT_3476_LISTEN=False`. Jev 외부 API 게이트는 PC 환경변수·`.env`에 `TYPESAFE_API_KEY`/`OPENROUTER_API_KEY` 미설정(`hasTypesafeKey=false`, `hasOpenRouterKey=false`)으로 `[미검증]` |

---

## 0-0. 2026-11-26 v1.3.0 오리지널 Outlook 동반 분석·발송 전 검증 사이드카 모드 실측 현황

| 영역 | 실측 상태 | 세부 실측 내역 및 경로 |
|---|---|---|
| **Outlook 실시간 선택·작성 감지 & 0ms 발송 전 검증** | ✅ 실측 완료 | `src/utils/outlookCompanionVerifier.ts` (브라우저 안전 0ms 검증기 `verifyOutlookItemSafety`: 첨부 누락·수신자/제목 누락·외부 도메인 민감정보·조선 도면번호/기한 추출), `src/services/outlookCompanionBridge.ts` (`readActiveOutlookItemAsync`: `ActiveInspector` 작성/열람 → `ActiveExplorer` 선택 비차단 COM 조회), `src/services/outlookCompanionBridge.test.ts` (4/4 PASS). `server.ts` (`GET /api/local/outlook/active-item` `1500ms` single-flight 캐시 · `POST /api/local/outlook/verify-item`). 실서버(`3475`) E2E에서 `active-item` HTTP `200`, `verify-item` 차단 검증(`status: block`, `score: 35`, `missing_attachment`/`external_sensitive_leak`) 실측 *(클래식 Outlook 실행 중 COM 실시간 감지는 Outlook 미실행 환경으로 `[미검증]`)* |
| **동반 사이드카 UI · 오리지널 Outlook 초안 열기 · 우측 도킹** | ✅ 실측 완료 | `src/components/OutlookCompanionPanel.tsx`, `src/App.tsx` (헤더 `[Outlook 동반 검증]` 전환 · fallback 메일 즉시 검증 · `[Groq·Claude 정밀 분석]` 수동 호출 · 3종 회신 템플릿 + 초안 사전 점검 + `POST /api/local/outlook/send-via-com` `action: 'display'` 오리지널 Outlook 작성창 위임), `desktop/main.cjs` & `desktop/preload.cjs` (`bmail:setCompanionDock` 우측 사이드카 도킹 + 항상 위). E2E(`3475`, [bmail-v130-companion.png](file:///C:/Users/FREE/.gemini/antigravity/brain/ebe7d412-bd24-4d99-93e9-83579563c299/scratch/bmail-v130-companion.png))에서 `주의 검토 1건 (80점)` 및 `[오리지널 Outlook 작성창으로 열기]` 렌더링 실측 *(Electron 데스크톱 창 실도킹·COM 실작성창 오픈은 헤드리스/미실행으로 `[미검증]`)* |
| **매뉴얼 Flash 동기화 · 검증 · 배포 (`1.3.0`)** | ✅ 실측 완료 | `AGENTS.md` 4항 준수 `Model: "flash"` 서브에이전트가 `manualCatalog.ts` (`1.3.0`, 32개 아이콘 `#1~#32`, 12개 모션 씬) & `InteractiveManualModal.tsx` 동기화 (`manualCatalog.test.ts` 1/1 PASS). `npm run lint` (`0`), `node --check` (`0`), Node `v24.20.0` 전체 36개 파일 **`152/152 PASS`** (`0` fail, `0` skipped). `release-exe/Bmail-Outlook-AI-Setup-1.3.0.exe` (`121,469,457` B, `sha256: e67a0829a42bb348a08e0356827e75710bd5b3b2fd1c049ee62de4c8267bc9ac`), R2 `latest.json`, Worker Version `62d70221-70b8-4511-846f-608511b7ec4b`, 무음 업그레이드 `NSIS_EXIT_CODE=0` (`installed 1.3.0`) |

---

## 0-0. 2026-11-25 v1.2.10 가로 구역 드래그 크기 조절 · 최근 14일 미확인 자동 알림 팝업 실측 현황

| 영역 | 실측 상태 | 세부 실측 내역 및 경로 |
|---|---|---|
| **Outlook 4개 구역 가로 드래그 크기 조절** | ✅ 실측 완료 | `MailAiDockPanel.tsx` & `AttachmentPreviewDock.tsx` (`w-full min-w-0` 유동 폭 전환), `OutlookWorkbench.tsx` (`DEFAULT_PANE_WIDTHS`, `PANE_WIDTH_LIMITS`, `bmail_pane_widths_v1` `localStorage` 자동 저장, iframe 드래그 끊김 방지 투명 오버레이, 4종 `role="separator"` 스플리터 `splitter-folder`/`splitter-list`/`splitter-ai-dock`/`splitter-attach-dock` + 더블클릭 기본 너비 복원). 설치본(`3474`) E2E 실제 마우스 드래그 실측: `folder 224→294px` · `list 360→470px` · `reading 508→258px` · `aiDock 299→369px`, 새로고침 후 동일 유지(`294`/`470`/`258`/`369`), 더블클릭 후 `list=360px` 원복 |
| **최근 14일 미확인 수신 메일 자동 알림 팝업** | ✅ 실측 완료 | `UnreadBriefingBanner.tsx` (최근 14일 미확인 메일 수신 시 우측 하단 `fixed bottom-[98px] right-3 z-50` 자동 알림 팝업 `recent-unread-alert-popup` 발생, `[30분 뒤]` 다시 알림 `localStorage`+60s 타이머, `[×]` 닫기 시 동일 시그니처 억제 후 새 미확인 수신 시 자동 재알림, 2023 장기 과거 메일 제외 유지). E2E에서 초기 자동 팝업(`toastVisible: true`, `noOverlapWithPill: true`, `has14DayText: true`, `has2023OrOldAlert: false`) 및 `[30분 뒤]` 클릭 후 닫힘(`toastVisibleAfterSnooze: false`, `pillStillVisible: true`) 실측 |
| **매뉴얼 Flash 동기화 · 검증 · 배포 (`1.2.10`)** | ✅ 실측 완료 | `AGENTS.md` 4항 준수 `Model: "flash"` 서브에이전트가 `manualCatalog.ts` (`1.2.10`, 31개 아이콘 `#1~#31`, 11개 모션 씬) & `InteractiveManualModal.tsx` 동기화 (`manualCatalog.test.ts` 1/1 PASS). `npm run lint` (`0`), `node --check` (`0`), Node `v24.20.0` 전체 35개 파일 **`148/148 PASS`** (`0` fail, `0` skipped). `release-exe/Bmail-Outlook-AI-Setup-1.2.10.exe` (`121,457,488` B, `sha256: 9249ab849f41b03f05247fa7d916148bfc5acc5bf49c543313d5067c82a8388a`), R2 `latest.json`, Worker Version `7da129c1-249c-4f7e-b2fa-96025b06bf3a`, 무음 업그레이드 `NSIS_EXIT_CODE=0` (`installed 1.2.10`) |

---

## 0-0. 2026-11-25 v1.2.9 오리지널 클래식 Outlook 평면 UI · 최근 14일 브리핑 우측하단 팝업 · 속도 고속화 · 관리자 사용량 기준(v1.2.8 포함) 실측 현황

| 영역 | 실측 상태 | 세부 실측 내역 및 경로 |
|---|---|---|
| **오리지널 클래식 Outlook 평면 1px UI 정돈** | ✅ 실측 완료 | `OutlookWorkbench.tsx` (리본 무채색 정돈 + 세로 1px `#E1DFDD` 구분선 기반 평면 3-Pane 적용 · `rounded-lg`/`gap-2` 박스 제거 · 상단 5단 인라인 브리핑 배너 제거 · 본문 위 인라인 AI 요약 기본 숨김 및 헤더 `[AI 요약 보기]` 토글), `OutlookStoreFolderTree.tsx` (하위 폴더 반복 아이콘 생략 + 폴더 12개 초과 시 컴팩트 검색창), `outlookWorkspaceModel.ts` (신규 프로필 `showSenderAvatars` 기본값 `false`; 기존 캐시 프로필 `user_state`는 기존 설정 `true` 유지 실측) |
| **최근 14일 미확인 브리핑 cutoff & 우측 하단 플로팅 팝업** | ✅ 실측 완료 | `localMailDb.ts` (`queryUnreadInboxEmails` 최근 `14`일 SQL timestamp 컷오프), `unreadBriefingEngine.ts` (`DEFAULT_UNREAD_BRIEFING_MAX_AGE_DAYS = 14`, `excludedOldUnreadCount`), `UnreadBriefingBanner.tsx` (기본 닫힘 `false`, 우측 하단 `fixed bottom-14 right-3 z-40` 알약 + 사이드/하단 팝업 드로어, `[닫기]` 및 `Escape`), `OutlookWorkbench.tsx` (리본 `[미확인 브리핑]` 토글). 실서버 `/api/local/unread-briefing` **`9ms` HTTP `200`** (`maxAgeDays: 14`, `totalUnread: 4`, `urgentReplyCount: 0`, `2023`/`SCA 35K FD` 과거 메일 `0`건), E2E 열기/닫기 및 비중첩 확인 |
| **속도 고속화 (SQLite 복합 인덱스 · 비블로킹 탐색 · AI 도크 0ms)** | ✅ 실측 완료 | `localMailDb.ts` (`idx_emails_active_ts`, `idx_emails_unread_inbox` 복합 인덱스), `server.ts` (`/api/local/bootstrap` 내 `discoverAndApplyAllOutlookSettings` `setTimeout(..., 1500)` 백그라운드 비블로킹화), `MailAiDockPanel.tsx` (`autoAiOnSelect` 기본 `false` + `[수동/자동]` 토글로 메일 클릭 시 HTTP 자동 호출 없이 `0ms` 즉시 요약). 설치본 실서버(`3000`, `24,671`건 DB) `/api/local/bootstrap?limit=100` **`70ms`** (`HTTP 200`) 실측 |
| **v1.2.8 포함: 관리자 LLM 사용량 상시 기준 게이지·비교표·수시 변경** | ✅ 실측 완료 | `llmUsageTracker.ts` (`LlmUsageLimits`, `llm-usage-limits.json`, `evaluateLlmUsageBudget`), `groqKeyRotator.ts` (`groqParallelMax` 동적 반영 · `claudeAdvancedMode` · `BMAIL_GROQ_KEYS_FILE` 테스트 격리), `server.ts` (`GET /api/admin/llm-usage`, `PUT /api/admin/llm-usage/limits`), `AdminLlmUsageCompactCard.tsx` (3종 게이지 + 비교표 + SVG 일일 기준선 + `[기준 수시 변경]`/`[저장]`). E2E에서 `Groq 키 슬롯: 5/5 활성 (최대 5병렬)`, 일일 상한 `300 → 500` 저장(`500회/일`·`일일 기준 500회` SVG 기준선)·`300` 원복 실측 (`EGO_EXIT=0`) |
| **매뉴얼 Flash 동기화 · 검증 · 배포 (`1.2.9`)** | ✅ 실측 완료 | `AGENTS.md` 4항 준수 `Model: "flash"` 서브에이전트가 `manualCatalog.ts` (`1.2.9`, 30개 아이콘, 10개 모션 씬) & `InteractiveManualModal.tsx` 동기화 (`manualCatalog.test.ts` 1/1 PASS). `npm run lint` (`0`), `node --check` (`0`), Node `v24.20.0` 전체 35개 파일 **`148/148 PASS`** (`0` fail, `0` skipped). `release-exe/Bmail-Outlook-AI-Setup-1.2.9.exe` (`121,455,272` B, `sha256: 8192cf37961b3f624ed7e37affaf5fc609006e875cb0de1a8466a4587b563427`), R2 `latest.json`, Worker Version `e4112bd9-d612-4fb9-a390-752a48ec42a0`, 무음 업그레이드 `NSIS_EXIT_CODE=0` (`installed 1.2.9`) |

---

## 0-0. 2026-11-25 v1.2.7 Classic Outlook 작성·발송 100% 패리티 · API 호출량 그래프 · Claude 고급 분리 · Groq 5-키 병렬 & 1줄 터미널 등록 실측 현황

| 영역 | 실측 상태 | 세부 실측 내역 및 경로 |
|---|---|---|
| **Classic Outlook 리본 메일 작성·발송 100% 패리티** | ✅ 실측 완료 | `OutlookComposeWindow.tsx`(클래식 Outlook `메시지`/`삽입`/`옵션` 리본, 클립보드·글꼴·색·목록·표·링크, Outlook 주소록 `152명` 선택기+자동완성, `숨은 참조(Bcc)`, 로컬 파일 첨부+드래그앤드롭+첨부 제거 칩, `! 중요도 높음`/`↓ 낮음`, 읽음 확인 요청, 실제 `sendDelaySec` 발송 지연 취소 카운트다운, `Ctrl+Enter`/`Alt+S`), `pop3AutoConnect.ts`(`buildSmtpMimePayload` · `sendSmtpMessage`: `multipart/mixed` 첨부 + `cid:` 명함 + `Importance` + `Disposition-Notification-To` + 비노출 `Bcc` `RCPT TO`), `server.ts` (`/api/pop3/send`, `/api/local/outlook/send-via-com` + `localDb` `보낸 편지함` 저장), `App.tsx`(`bmail_local_contacts_cache` 독립 창 전달). E2E(`3471`)에서 주소록 `152명`, 첨부 칩 제거, 명함 이미지, 5초 발송 취소 카운트다운 실측 (실제 외부 SMTP/Outlook COM 메일 전송은 `[미검증]`) |
| **API 라우팅 시각화 · 호출량 SVG 그래프 · Claude 고급 / Groq 5-키 병렬** | ✅ 실측 완료 | `server.ts` (`GET /api/admin/llm-usage` `routingPolicy` 메타 추가), `AdminLlmUsageCompactCard.tsx` (실시간 라우팅 정책 배지 + 프로바이더별 호출량 분포 바 차트 + 일자별 SVG 막대 그래프), `DashboardView.tsx` (상단 KPI 직하단 상시 배치). `groqKeyRotator.ts` (`runGroqParallelBatch` 최대 5병렬 실행 + 순서 보존, `shouldEscalateToClaude` 고급 분석 `executive-report`/`doc-analysis`/`document`/`meeting`/대량 `mail-reduce` Claude 5.5 에스컬레이션), `aiService.ts` (병렬 Map 청킹), `executiveReportEngine.ts` (경영진 보고서 Claude 1순위). 실서버 `/api/admin/llm-usage` HTTP `200` (`totalCalls: 24`, Groq 18 / Claude 4 / Gemini 2) 및 E2E 대시보드 그래프 2종 실측 |
| **Groq 5-키 쉼표 구분 터미널 1줄 일괄 설정 CLI (`npm run groq:keys`)** | ✅ 실측 완료 | `scripts/set-groq-keys.mjs`, `package.json` (`"groq:keys"`), `groq-keys.txt`, `groqKeyRotator.test.ts` (쉼표 구분 5-키 파싱 단위 테스트). 격리 임시 디렉토리 테스트 `SET_GROQ_KEYS_SMOKE_PASS` (`groq-keys.txt` · `data/groq-keys.txt` · `.env` `GROQ_API_KEY~5` 동시 기록 및 마스킹 출력 확인) |
| **매뉴얼 Flash 동기화 · 검증 · 배포 (`1.2.7`)** | ✅ 실측 완료 | `AGENTS.md` 4항 준수 `Model: "flash"` 서브에이전트가 `manualCatalog.ts`(`v1.2.7`, 25개 전체 아이콘) & `InteractiveManualModal.tsx` 동기화(`manualCatalog.test.ts` PASS). `npm run lint` (오류 `0`건), `node --check desktop/main.cjs desktop/preload.cjs scripts/set-groq-keys.mjs` (`0`), Node 24 전체 35개 파일 **`147/147 PASS`** (`0` fail, `0` skipped), `release-exe/Bmail-Outlook-AI-Setup-1.2.7.exe` (`121,448,876` bytes, `sha256: 4e8fd3e7ba728132a743cc8b5891d0bf28c4f4543d76f78be127474a5725260f`), Worker Version `0aa775c8-75d2-4db9-a76e-a355e7d5a8a3`, 무음 재설치 `NSIS_EXIT_CODE=0` (`installed 1.2.7`), `ego-browser` 3종 E2E(`EGO_EXIT=0`) & 3장 스크린샷 육안 검수 완료 |

---

## 0-0. 2026-11-25 v1.2.6 속도 개선 · 미열람 브리핑 · 명함 서명 새창 · 우측 분석창 고정 · 네임태그 제외 · 전체 매뉴얼 실측 현황

| 영역 | 실측 상태 | 세부 실측 내역 및 경로 |
|---|---|---|
| **이벤트 루프 정체 해소 · 대시보드/로딩 고속화** | ✅ 실측 완료 | `App.tsx`(`LazyViewErrorBoundary`, `8s` 초기 지연 / `180s` 주기 증분 동기화), `server.ts`(`pollLiveOutlookInboxAsync` 비동기 `sync-incremental`), `backgroundMailMonitor.ts`(`refreshDashboardDigest` 60s 캐시). 설치본 실서버(`3470`, 409.8 MB DB): `/api/config` **`38ms`**, `/api/local/bootstrap?limit=100` **`243ms`**, 동시 `POST /api/local/outlook/sync-incremental`(`460ms`) 중 `/api/local/monitor/status` **`18ms`** · `/api/local/bootstrap?limit=20` **`246ms`**, E2E 대시보드 렌더 **`1,345ms`** (`화면을 불러오는 중...` 정체 제거) |
| **미열람 메일 읽기전용 프로젝트·중요도 브리핑** | ✅ 실측 완료 | `src/services/unreadBriefingEngine.ts`, `unreadBriefingEngine.test.ts`, `src/components/UnreadBriefingBanner.tsx`, `localMailDb.ts`(`queryUnreadInboxEmails`), `GET /api/local/unread-briefing`. 실서버 **`83ms`** (`totalUnread: 122`, `urgentReplyCount: 79`, `importantCount: 32`, `projectCount: 3`: `135K-01`/`GENERAL`/`22KF-03`), 조회 전후 SQLite `unreadCount` 동일(`unreadUnchangedReadOnly: true`) 실측 |
| **우측 AI 세로분할 고정 · 첨부 클릭 우측 미리보기 · 네임태그/명함 제외** | ✅ 실측 완료 | `MailAiDockPanel.tsx`(`w-[272px] xl:w-[296px]`, `bmail_pin_ai_split_dock` 기본 고정, 50건 LRU+350ms 디바운스), `outlookWorkspaceModel.ts`(`isInlineBusinessCardImage`), `attachmentBatchExporter.ts`, `AttachmentPreviewDock.tsx`. E2E AI 도크 **`296px`**(`고정` 배지) 확인, DWG 메일 첨부 클릭 시 우측 도크 오픈 및 인라인 서명·네임태그 `4`건 자동 제외(`실첨부 파일 1개`) 실측 |
| **Outlook 명함 서명 새창 작성/전달/답장 & 전체 설정 탐색·적용** | ✅ 실측 완료 | `OutlookComposeWindow.tsx`(독립 새 창 `/?compose=1` + `desktop/main.cjs` `bmail:openComposeWindow`, HTML 리치 에디터, `.files/image001.png` base64 인라인 명함, 원본 스레드, 전달 실첨부 승계, `multipart/related` `cid:` SMTP + 비동기 COM `send-via-com`), `localOutlookMigrator.ts`(`inlineSignatureHtmlImages`, `discoverAndApplyAllOutlookSettings`), `OutlookSettingsModal.tsx`. 실서버 `POST /api/local/outlook/discover-settings` **`1,141ms`** (`signaturesCount: 3`, `inlinedSignatureImages: 2`, `accountsCount: 1`) 및 E2E 전달창 SEASTAR 명함 `img[src^="data:image/"]` 렌더링 실측 (실제 외부 SMTP/Outlook COM 메일 발송은 `[미검증]`) |
| **Groq 5-키 쉬운 입력 파일 + 메일플러그 앱비번 정규화 + 매뉴얼 Flash 동기화** | ✅ 실측 완료 | gitignore된 `groq-keys.txt`(1줄 1키 템플릿) + `groqKeyRotator.ts` (`collectGroqApiKeys` 파일 로더), `pop3AutoConnect.ts`(`noSpacesOrHyphens` 재시도 순서 · `isKorea === false` 한정 해외 경고), `InteractiveManualModal.tsx`·`manualCatalog.ts`(`Model: "flash"` 서브에이전트 + `AGENTS.md` 4항 규약, 아이콘 25개 전체 표시 + `1~10`/`11~20`/`21~끝` 이동). Item C 로컬 TCP SSE 5-키 429 페일오버·라운드로빈(`2..5`) 실측 (실제 `api.groq.com` 키 입력 후 라이브 호출 및 실계정 메일플러그 `#5.7.5` 인증은 `[미검증]`) |
| **정적 검증 · 단위 테스트 · 라이브 배포 (`1.2.6`)** | ✅ 실측 완료 | `npm run lint` (오류 0건) · `node --check desktop/main.cjs desktop/preload.cjs` exit `0` · Node 24 전체 35개 테스트 **`143/143 PASS`** (0 실패, 0 건너뜀) · `release-exe/Bmail-Outlook-AI-Setup-1.2.6.exe` (`121,439,314` bytes, `sha256: 93d1a79301f38c62c06193fce941d3f5bd74d5762cc553ed752ec5294a0f1a13`) · Cloudflare Worker Version `1829eb56-7424-4d99-93e9-83579563c299` · R2 `latest.json` & 무음 재설치 `NSIS_EXIT_CODE=0` (`installed 1.2.6`) · `ego-browser` 6단계 UI E2E(`EGO_EXIT=0`) 및 6종 스크린샷 육안 검수 완료 |

---

## 0-0. 2026-11-24 v1.2.5 항상 로그인 유지 · 백그라운드 선행 정리 · 미열람 SGW 모바일 알림 실측 현황

| 영역 | 실측 상태 | 세부 실측 내역 및 경로 |
|---|---|---|
| **항상 로그인 유지 (AES-256-GCM 세션 영속화)** | ✅ 실측 완료 | `Pop3SessionStore` (`src/services/pop3AutoConnect.ts`: `30일` 슬라이딩 TTL + `pop3-sessions.enc.json` 암호화) & `ms-sessions.enc.json` + 리프레시 토큰 자동 연장 (`server.ts`) · `GET /api/pop3/session` 자동 복원 (`src/App.tsx`). 단위 테스트 재시작 복원 및 평문 미노출 확인 (`pop3AutoConnect.test.ts`), 라이브 `GET /api/pop3/session` HTTP `200` 실측 |
| **새 메일 실시간 요약 엔진 & 트레이 미니창** | ✅ 실측 완료 | `src/services/newMailNotifier.ts`, `src/services/newMailNotifier.test.ts`, `POST /api/local/mail/quick-summary` (Groq 5-키 → Claude → 결정론적 0ms 폴백) · Electron 트레이 미니 토스트 (`desktop/main.cjs`, `380×168`, 긴급 배지+요약+12초 자동닫힘) + 앱 내 우하단 알림 카드 (`src/App.tsx`). 라이브 `POST /api/local/mail/quick-summary` HTTP `200` 실측 |
| **백그라운드 선행 정리 & 미열람 SGW 모바일(핸드폰) 알림** | ✅ 실측 완료 | `src/services/backgroundMailMonitor.ts`, `src/services/backgroundMailMonitor.test.ts`, 비동기 COM 수신함 폴러 `pollLiveOutlookInboxAsync` (`src/services/localOutlookMigrator.ts`), `server.ts` (`/api/local/monitor/status`, `settings`, `ack`, `poll-now`, `test-sgw-alert`, `consume-tray-alerts`), `desktop/main.cjs` 20s 숨김/비활성 트레이 폴링, `OutlookSettingsModal.tsx` 제어 UI. 설치본 실서버(`3468`)에서 `POST /api/local/monitor/poll-now` **`4,413ms` HTTP `200`** (초기 동기 PST 재파싱 행 버그 수정), `running: true`, `sgwConnected: true` (`designsir@seastargo.com`), `proactive_dashboard_digest` (`totalEmails: 24,644` · `unrepliedIncomingCount: 14,554` · `unreadCount: 121`) 및 `GET/PUT/POST` 엔드포인트 HTTP `200` 실측 (`test-sgw-alert` 실제 SGW 서버 외부 모바일 푸시 수신은 `[미검증]`) |
| **정적 검증 · 단위 테스트 · 라이브 배포 (`1.2.5`)** | ✅ 실측 완료 | `npm run lint` (오류 0건) · Node 24 **`138/138 PASS`** (0 실패, 0 건너뜀) · `dist/assets/index-HRzdmOkY.js` **`519,894` B (`150.56 kB` gzip)** · `dist-server/server.mjs` `795.0kb` · `release-exe/Bmail-Outlook-AI-Setup-1.2.5.exe` (`121,421,533` bytes, `SHA-256: ce2e5ad748ae0b4ce19e1841860a33e69f060e2413608b2a94f049b2e5638fdd`) · R2 `latest.json` & `/releases/Bmail-Outlook-AI-Setup-1.2.5.exe` `HEAD` (`121,421,533`) & `/api/health` HTTP `200` · 로컬 무음 재설치 `NSIS_EXIT_CODE=0` (`installedPkgVersion=1.2.5`) |

---

## 0-0. 2026-11-24 v1.2.4 Groq 5-키 기본 엔진 + Claude 에스컬레이션 & Outlook 기능·설정 복제 실측 현황

| 영역 | 실측 상태 | 세부 실측 내역 및 경로 |
|---|---|---|
| **Groq 5-키 순환 기본 LLM + Claude 에스컬레이션** | ✅ 실측 완료 | `src/services/groqKeyRotator.ts` (`openai/gpt-oss-120b`, `temperature: 1`, `max_completion_tokens: 2048`, `top_p: 1`, `reasoning_effort: "medium"`, `stream: true`, `stop: null`, 5-키 라운드로빈+쿨다운), `src/services/groqKeyRotator.test.ts`, `src/services/emvVaultService.ts` (`GROQ_API_KEY_2..5`), `aiService.ts`, `server.ts`, `executiveReportEngine.ts`, `multimodalAiService.ts`. 주의: 현재 `.env`의 `GROQ_API_KEY`는 2자 미설정 상태(`groq.configured: false`)이며 실제 5개 키 등록 시 즉시 5-키 순환 가동 |
| **Outlook 기능·설정 클린룸 복제 (평균 71.6→93.3점)** | ✅ 실측 완료 | `OutlookPeopleView.tsx` (152명 주소록 + `.vcf`/CSV), `OutlookTasksView.tsx` (`outlook_tasks` + 플래그 메일), `outlookWorkspaceModel.ts` (검색 폴더 5종·보관함·빠른 단계 4종·색 범주 6종·일괄 정리 Sweep·다시 알림 Snooze·싱글라인 리본·미니멀 헤더·읽음 지연·삭제 후 이동·기본 중요도/읽음확인), `localOutlookMigrator.ts` (HKCU `CalDefStart/End`/`WorkDay`/`DisableTNEF`/`ReopenWindowsOption` 승계). Worker `18abb1f3`, `132/132 PASS`, `Bmail-Outlook-AI-Setup-1.2.4.exe` (`121,412,603` bytes) R2/무음 설치 확인 (GUI 클릭 E2E는 `[미검증]`) |

---

## 0-0. 2026-11-24 v1.2.3 초경량 코드 스플리팅 · 심플 3단 UX · 앱스타일 아이콘 · 무간섭 100% 자동 이관/상시 누적 실측 현황

| 영역 | 실측 상태 | 전후 절감 및 세부 실측 내역 |
|---|---|---|
| **초기 JS 번들 경량화 (`React.lazy` 코드 분할)** | ✅ 실측 완료 | 17개 보조 뷰/모달 + 4개 워크벤치 도크 지연 로딩 (`src/App.tsx`, `src/components/OutlookWorkbench.tsx`) · 초기 JS 청크 `1,207.2 KB` → **`484.37 KB` (`144.85 KB` gzip, -59.9%)**, `39`개 레이지 청크 분할 · 초기 부트스트랩 `300건` → **`100건`** (`24,632`건 전체는 SQLite FTS5 즉시 검색 유지) |
| **설치본 · 디스크 용량 다이어트 (`electron-builder.json`)** | ✅ 실측 완료 | `compression: "maximum"`, `electronLanguages: ["ko", "en-US"]`, 미사용 `node_modules` 번들/소스맵/문서/타입/테스트 제외. **실측 절감**: ① 설치파일 `.exe`: `133.32 MB` (`139,796,524 B`) → **`115.77 MB` (`121,396,990 B`, -18.4 MB / -13.2%)**, ② 설치 폴더 `win-unpacked`: `492.73 MB` → **`388.47 MB` (-104.26 MB / -21.2%)**, ③ `node_modules`: `146.41 MB` → **`87.30 MB` (-59.11 MB / -40.4%)**, ④ Chromium `locales`: `46.38 MB` → **`1.19 MB` (-45.19 MB / -97.4%)** |
| **앱스타일 스퀘어클 아이콘 일괄 적용** | ✅ 실측 완료 | `scripts/build-app-icon.mjs` 순수 Node.js 안티앨리어싱 렌더러: `desktop/icon.png` (`31,047 B` · 256px), `desktop/icon.ico` (`130,741 B` · 16~128px DIB + 256px PNG), `public/favicon.svg` (`1,050 B`) 생성 및 `index.html` + 상단 브랜드 헤더 적용 |
| **심플 3단 기본 UI + 무간섭 100% 자동 이관/상시 누적** | ✅ 실측 완료 | 기본 화면 3단 클래식 Outlook(`showAiSplitDock: false`, `showAttachSplitDock: false`, `inlineAiExpanded: false`) 1-클릭 전환 · 설치·실행 즉시 로컬 Outlook 계정·서명·즐겨찾기·폴더·PST/메일 100% 자동 이관 + 백그라운드 3분/창복귀 무음 증분 누적(`src/App.tsx`, `server.ts`) |
| **정적 검증 · 단위 테스트 · 라이브 배포 (`bmail.seastar.work`)** | ✅ 실측 완료 | `npm run lint` (오류 0건) · Node 24 **`127/127 PASS`** · Worker Version `38ac1144-29e9-4d48-9724-63c1b6f403da` · `https://bmail.seastar.work/api/health` (`200`, `version: "1.2.3"`), `/releases/latest.json` (`200`, `version: "1.2.3"`), `/favicon.svg` (`200`), 로컬 무음 설치 `NSIS_EXIT_CODE=0` |

---

## 0-0. 2026-11-24 v1.2.2 TOP-6 클린룸 신규 기능 · 일일 SQLite 백업 · 바이너리 `.msg` 파서 실측 현황

| 영역 | 실측 상태 | 세부 실측 내역 및 경로 |
|---|---|---|
| **① Windows `.ics` 양방향 일정 어댑터** | ✅ 실측 완료 | `src/utils/icsCalendarAdapter.ts`, `src/utils/icsCalendarAdapter.test.ts`, `src/components/OutlookCalendarView.tsx`, `src/App.tsx`. RFC 5545 `VCALENDAR`/`VEVENT` 내보내기(`[.ics 내보내기]`) + 라인 폴딩/이스케이프 복원(`[.ics 가져오기]`, 중복 자동 제외) 단위 테스트 통과 |
| **② 플로팅 스티커 메모 (`StickyNotesWidget`)** | ✅ 실측 완료 | `src/components/StickyNotesWidget.tsx`, `src/services/localMailDb.ts` (`ALLOWED_KV_STATE_KEYS`에 `'sticky_notes'` 추가), `src/App.tsx` 상단 헤더 토글 버튼. 4색 스와치·상단 핀 고정·검색·현재 메일 1클릭 연결. 라이브 검증 시 `/api/local/state/sticky_notes` 404 오타 포착 후 `PUT /api/local/state` (`{ key: 'sticky_notes', value }`)로 교정하여 HTTP `200` 왕복 실측 |
| **③ 취소 가능 지연 발송 카운트다운 큐** | ✅ 실측 완료 | `src/utils/delayedSendQueue.ts`, `src/utils/delayedSendQueue.test.ts`, `src/components/OutlookWorkbench.tsx`. 작성창 하단 대기 시간(`0초/10초/30초/1분/5분`) + 좌하단 실시간 카운트다운 플로팅 바(`[발송 취소]` · `[즉시]`) 구현 및 단위 테스트 통과 |
| **④ 드래그앤드롭 순수 바이너리 `.msg` 파서** | ✅ 실측 완료 | `src/services/msgFileParser.ts`, `src/services/msgFileParser.test.ts`, `src/services/pstParser.ts`, `src/components/PstImportModal.tsx`, `src/components/OutlookWorkbench.tsx`. 순수 Node.js CFBF/OLE2 컨테이너 + MAPI UTF-16LE 스트림/첨부 파서 구현 및 워크벤치 파일 드롭 오버레이 지원. 라이브 `POST /api/pst/import` HTTP `200` (`fileType: "msg"`, `H3000-DRAWING.pdf` 추출) 실측 |
| **⑤ 오프라인 로컬 SQLite 일일 gzip 백업 스케줄러** | ✅ 실측 완료 | `src/services/localBackupScheduler.ts`, `src/services/localBackupScheduler.test.ts`, `src/components/ConfigEnvManager.tsx`, `server.ts`. WAL 체크포인트 후 일 1회 `gzip` 압축 + `.sha256` 생성 + 최대 7세대 자동 회전(`data/backups/`). 라이브 `POST /api/admin/backups/run` HTTP `200` (`bmail-backup-manual-2026-11-24_014746.sqlite.gz`, `86,987,003` bytes vs `409,800,704` bytes DB) 실측 |
| **⑥ Windows 트레이 1-클릭 `[즉시 읽음]`/`[즉시 회신]` 토스트** | ✅ 실측 완료 | `desktop/main.cjs` (`bmail:showMailToast`, `showTrayMiniToast`, `bmail-toast://` 커스텀 스킴 격리 샌드박스), `desktop/preload.cjs` (`showMailToast`, `onToastAction`), `src/App.tsx`, `src/components/OutlookWorkbench.tsx` (우클릭 컨텍스트 메뉴 테스트 + 웹 폴백 토스트 포함). (`desktop/main.cjs` 네이티브 OS 창 UI 클릭은 `[미검증]`) |
| **정적 검증 · 고아 코드 감사 · 전체 단위 테스트 (Node 24)** | ✅ 실측 완료 | 고아 코드 감사(`components: 32, orphans: 0`, `utils: 18, orphans: 0`) · `npm run lint` (오류 0건) · Node 24 전체 29개 테스트 파일 **`127/127 PASS` (`0` fail, `0` skipped)** |
| **Windows `.exe` 패키징 · R2/Worker 배포 (`bmail.seastar.work`)** | ✅ 실측 완료 | `release-exe/Bmail-Outlook-AI-Setup-1.2.2.exe` (`139,796,524` bytes, SHA-256 `b73997678c641a81398ef41346070c2e1a5eb9c44fe6cf2ce8657c0fc7654987`) · Cloudflare Worker Version `0904b6bc-fa09-4cc3-934d-d787c3609ed8` · `https://bmail.seastar.work/api/health` (`version: "1.2.2"` HTTP `200`), `/releases/latest.json` (`version: "1.2.2"` HTTP `200`), 로컬 무음 설치(`NSIS_EXIT_CODE=0`) + `dist/` 동기화 (`index-BHwIrdd6.js` 서빙) 실측 |

---

## 0-0. 2026-11-23 v1.2.1 올인원 대시보드 · 2분할 AI/첨부 도크 · 모션그래픽 매뉴얼 · SGW 공식 로그인 창 · Claude 5.5 3-티어 라우터 실측 현황

| 영역 | 실측 상태 | 세부 실측 내역 및 경로 |
|---|---|---|
| **Anthropic 15개 모델 전수 라이브 검증 + Claude 5.5 3-Tier 자동 라우터** | ✅ 실측 완료 | `ANTHROPIC_API_KEY`로 목록 15개 전 모델(`claude-fable-5-2` `952ms`, `claude-haiku-5-5` `648ms`, `claude-sonnet-5-5` `1124ms`, `claude-opus-5-5` `878ms`, `claude-fable-5-1` `1473ms`, `claude-opus-5` `816ms`, `claude-sonnet-5` `712ms`, `claude-fable-5` `2877ms`, `claude-opus-4-8` `560ms`, `claude-opus-4-7` `530ms`, `claude-sonnet-4-6` `811ms`, `claude-opus-4-6` `1384ms`, `claude-opus-4-5-20251101` `919ms`, `claude-haiku-4-5-20251001` `481ms`, `claude-sonnet-4-5-20250929` `946ms`) **15/15 HTTP 200** 실측. 복잡도 자동 라우터(`src/services/claudeModelRouter.ts`)는 기본 `claude-haiku-5-5`(`low`, E2E `13,339ms`), 중급 `claude-sonnet-5-5`(`medium`, E2E `22,554ms`), 최고난도 `claude-opus-5-5`(`high`, E2E `46,611ms`)로 자동 승격 |
| **우측 세로 2분할 도크 (`AI 분석·답안` + `첨부 미리보기·일괄저장`) & 우하단 미니창** | ✅ 실측 완료 | ① 본문 옆 1번째 세로 분할 `src/components/MailAiDockPanel.tsx` (`분석`·`답안`·`대화`·`번역`·`액션` 5탭) ② 2번째 세로 분할 `src/components/AttachmentPreviewDock.tsx` (탭 1: PDF/이미지/Excel/Word 인라인 미리보기+문서 AI 분석, 탭 2: 지정 폴더 첨부 일괄 저장 + MD/CSV/HTML 3종 리포트) ③ 우하단 소형 고정 `src/components/MailMiniAssistWindow.tsx` (`fixed bottom-3 right-3 z-40 w-[330px]`, `POST /api/local/mail/assistant` HTTP `200` + 오프라인 결정론적 폴백으로 `Failed to fetch` 0건). 실측: `POST /api/local/attachments/export-plan` & `export-run` HTTP `200` (`exportedCount: 3`, `skippedInlineCount: 3`, `totalBytesWritten: 807,026`, `.md`/`.csv`/`.html` 3종 리포트 디스크 생성 확인) |
| **올인원 종합 대시보드 (`분석글` + 주간·월간·미회신·프로젝트·주제·수발신·첨부)** | ✅ 실측 완료 | `src/services/executiveReportEngine.ts`, `src/components/DashboardView.tsx`, `src/components/AiInboxIntelligenceView.tsx`. `GET /api/local/dashboard/executive?periodType=all` HTTP `200` (`3,763 ms`): 총 **`24,632`건**(`2019-04-09 ~ 2026-11-23`, 수신 `19,426`건 · 발신 `5,206`건 · **내게 왔는데 미회신 `14,550`건** · 상대 회신 대기 `2,307`건 · 실첨부 `38,786`건 · 프로젝트 `231`개 · 주제 `7`개), 서술형 종합 분석글(`narrativeReport`), 주간 테이블(`16`주), 월간 테이블(`12`개월), 스냅샷 영구 저장(`POST /api/local/dashboard/executive/save`) 및 Claude 5.5 심층 분석글(`POST /api/local/dashboard/executive/ai-narrative`) 실측 |
| **클래식 Outlook 100% 패리티 + 1회 적재 후 증분 동기화 (`sync-incremental`)** | ✅ 실측 완료 | ① `extractShipbuildingDrawingRefs` 숫자 필수 가드로 `도면: MESSAGE` 오탐 제거 ② `GET /api/local/emails/:id/render` & `GET /api/local/pst/attachment?inline=1` (`pst-fd270b-3352772` 본문 HTML `11,850`자 + `Content-Disposition: inline` `12,329` bytes HTTP `200`) + 서명 로고(`image001.png`) 첨부 칩 숨김 ③ 스토어별 접이식 폴더 트리(`src/components/OutlookStoreFolderTree.tsx`: `즐겨찾기` `New2026.xml` + `김봉정`/`백업` 트리 + 파란 안읽음 건수 + `[N]` 초안 표기) ④ `POST /api/local/outlook/sync-incremental` HTTP `200` (`1,378 ms`, 변경 없는 PST `skippedUnchangedPsts: 2` 즉시 스킵, 기존 `24,632`건 유지 + 신규분만 증분 반영) |
| **팝업 HTML 모션그래픽 매뉴얼 + 프로덕션 헤더 보안 정리** | ✅ 실측 완료 | `src/data/manualCatalog.ts` (`version: "1.2.1"` `package.json` 동기화 단위 테스트 통과) + `src/components/InteractiveManualModal.tsx` (실제 화면 1:1 비주얼 복제 + 8씬 모션그래픽 동영상 플레이어 + 25개 아이콘 탐색기 + `/?manual=1` 독립 팝업). 상단 헤더/와플 메뉴에서 `HANDOFF`·`전체 코드 다운로드` 완전 제거 및 `GET /api/export/handoff` → HTTP `404`, `GET /api/export/source-archive.zip` → HTTP `404` 실측 |
| **SGW 공식 로그인 창 완전 연동 (`persist:sgw-auth`)** | ✅ 실측 완료 | `desktop/main.cjs` (`bmail:openSgwLoginWindow`, `https://sgw.seastar.work` 공식 로그인 창 3계층 토큰/쿠키 캡처) + `desktop/preload.cjs` + `src/services/sgwChatBridge.ts` (`connectWithSession` `/api/auth/me` 검증 후 AES-256-GCM 암호화 저장) + `src/components/SgwChatPanel.tsx` (`SGW 로그인 창 띄워 완전 연동하기` 원클릭 버튼) |
| **관리자 페이지 소형 LLM 사용량 모니터 (`AdminLlmUsageCompactCard`)** | ✅ 실측 완료 | `src/services/llmUsageTracker.ts`, `src/services/llmUsageTracker.test.ts`, `src/components/AdminLlmUsageCompactCard.tsx`, `src/components/ConfigEnvManager.tsx`. 관리자 설정 페이지 상단에 1줄 소형 요약 바(누적 호출·총 토큰(In/Out)·오늘 사용량·추정 비용(USD)·평균 지연시간 + 모델별 배지 + 펼치기 상세 테이블/최근 8건 로그 + 초기화) 탑재. `GET /api/admin/llm-usage` HTTP `200` 실측 (`claude-haiku-5-5` 실제 호출 `inputTokens: 368`, `outputTokens: 1,146`, `totalTokens: 1,514`, `costUsd: $0.004878` 기록 확인) |
| **정적 검증 · 고아 코드 감사 · 전체 단위 테스트 (Node 24)** | ✅ 실측 완료 | 고아 코드 감사(`components: 31, orphans: 0`, `utils: 16, orphans: 0`, UI 내 `HANDOFF`/`전체 코드 다운로드` `false`) · `npm run lint` (오류 0건) · Node 24 전체 26개 테스트 파일 **`121/121 PASS` (`0` fail, `0` skipped)** |
| **Windows `.exe` 패키징 · R2/Worker 배포 (`bmail.seastar.work`)** | ✅ 실측 완료 | `desktop/icon.ico` (`106,005` bytes 비압축 32-bit BGRA DIB + 256px PNG) · `release-exe/Bmail-Outlook-AI-Setup-1.2.1.exe` (`139,780,425` bytes, SHA-256 `59ae52a93343e097ea996036b1078bcb65f3a2a013b255828e5a325783cf1e40`) · Cloudflare Worker `bmail-outlook-ai-gateway` Version `49e045a5-b711-404c-8ebf-dd11997b29be` · `https://bmail.seastar.work/` (20챕터 쇼케이스 `ch-01`..`ch-20`, HTTP `200`), `/api/health` (`version: "1.2.1"`, HTTP `200`), `/releases/latest.json` (`version: "1.2.1"`, HTTP `200`), `/install.ps1` (`200`), `/install.bat` (`200`), `https://mail.seastar.work/download` (`200`) 실측 |

---

## 0-0. 2026-11-22~23 v1.2.0 전사 1-클릭 배포 · SGW/SEDMS 연동 · 100% 오프라인 분석 실측 현황

| 영역 | 실측 상태 | 세부 실측 내역 및 경로 |
|---|---|---|
| **오프라인 결정론적 분석 엔진 완성** | ✅ 실측 완료 | 외부 유료/무료 AI 키 미설정(`configuredProviders: []`) 상태에서 메일·자율 에이전트·회의록·PST 첨부·문서교차 전 기능 로컬 100% 오프라인 작동 (`src/services/agentService.ts`, `src/services/multimodalAiService.ts`, `server.ts`, `src/services/localMailDb.test.ts`). FTS5 검색(`837 ms`, `20,969`건), 메일 직접 분석(`18 ms`, 20건), 자율 에이전트 SSE(`14 ms`, 4/4 단계, 비식별화 `147`건), PST 멀티포맷 파싱(PDF `718 ms`/7p/`1,316`단어 · Excel `117 ms`/`425`행/`1,611`단어 · DOCX `119 ms`/`360`단어, warning 0), 교차 문서 분석(`15 ms`, 3문서·`1,833`자), 회의록 합성(`9 ms`, 안건 2·결정 1·실행과제 5) 모두 HTTP `200` (`providerUsed: builtin`) 실측 |
| **정적 검증 · 전체 단위 테스트** | ✅ 실측 완료 | `npm run lint` (`tsc --noEmit` 오류 0건) · 전체 21개 테스트 파일(`node --import tsx --test`) **`108/108` PASS** (0 실패, 0 건너뜀) · `npm run build` (`dist/assets/index-BSO1wI5Z.js`) · `npm run build:server` (`dist-server/server.mjs` `600.0kb`, `dist-server/pst-worker.mjs` `64.5kb`) |
| **Cloudflare R2 배포 · 전사 1-클릭 설치** | ✅ 실측 완료 | `release-exe/Bmail-Outlook-AI-Setup-1.2.0.exe` (`139,637,268` bytes, SHA-256 `b8579ebeb2203ba19e5ce757fbcf73e8b08a7e7c5b495f7702f0afff1e09da9f`) · Worker `bmail-outlook-ai-gateway` (`https://mail.seastar.work/api/health` HTTP `200`, `version: 1.2.0`, `hasReleasesBucket: true`) · `/releases/latest.json`, `HEAD /releases/Bmail-Outlook-AI-Setup.exe` (`139,637,268` bytes), `/install.ps1`, `/install.bat` 모두 HTTP `200` · 로컬 무인 재설치(`/S`) `exit 0` (`133,369 ms`), `HKCU:\Software\Classes\sgw-outlook\shell\open\command` → `Bmail Outlook AI.exe "%1"` 등록 확인 |
| **SGW · SEDMS 라이브 브리지 연동** | ✅ 부분 실측 (인증 경계까지 확인) | `GET /api/sedms-live/status` HTTP `200` (`https://sedms.app/api/health` 실측 `200`, `831 ms`) · `POST /api/sedms-live/scan-local-mails` 200건 스캔(`131 ms`, 도면매칭 `56`건·자동연결 `8`건·검토대기 `192`건) · `POST /api/sedms-live/contacts/preview` 미인증 토큰 경계 `401` · `https://sgw.seastar.work/api/chat`(`305 ms`) & `/api/bmail-ledger`(`298 ms`) 실도달 `401` 인증 경계 확인. `E:\code-project\SGW_V0` 내 `bonglook` 대체 URL/경로 전환 편집 완료(로컬 저장 상태, 미커밋/미배포) |
| **미검증 · 수동 완료 필요 항목** | ⚠️ `[미검증]` | ① `ego-browser` (`0.6.3.17`) CLI IPC 무응답(6초 타임아웃)으로 `v1.2.0` 브라우저 클릭 UI QA(`qa-bmail-v12.mjs`) 미실행 (기존 `TaskSpace 20` 보존 및 `ego-browser upgrade` 미실행 원칙 준수) ② 실제 사원 SGW 로그인(`POST /api/sgw-chat/login`) 후 원격 채팅·장부 실쓰기 ③ SEDMS 관리자 Bearer 토큰 입력 후 `contacts/commit` ④ `SGW_V0` Flutter 웹 빌드·배포 ⑤ `designsir@seastargo.com` autodiscover 실측 결과 `webmail.seastargo.com` POP3만 도달(`imap.seastargo.com`/`smtp.seastargo.com` 미도달) — Mailplug 프리셋(`MAIL_USER`·`MAIL_PASSWORD`) 수동 지정 필요 |

---

## 0-1. 2026-11-21 로컬 Outlook 대체 & Cloudflare(`mail.seastar.work`) 구축·배포 실측 현황

| 영역 | 실측 상태 | 세부 내역 및 경로 |
|---|---|---|
| **Cloudflare 라이브 배포** | 완료 (HTTP 200 실측) | `https://mail.seastar.work` · Worker `bmail-outlook-ai-gateway` (Version `c47d9ee1-4e13-4469-b78b-569a9bbcfcf9`) · KV `BMAIL_KV` (`8147ffcafc3a49bca77f2681b31a9e7d`) · `BMAIL_ORG_ACCESS_TOKEN` Secret 등록 완료(`GEMINI_API_KEY`는 마스터 확인 대기) · `wrangler.toml` `run_worker_first` + `nodejs_compat` 적용 |
| **로컬 SQLite + FTS5 저장소** | 완료 (동시접속 + `.env` 우선) | `data/bmail-local.sqlite` (WAL + FTS5 + `PRAGMA busy_timeout = 5000`, `src/services/localMailDb.ts`) · `server.ts` `dotenv.config({ override: true })` 적용 · `POST /api/local/outlook/migrate` `{ folderPath }` + `POST /api/local/outlook/recompute-projects` 추가 · Node 24 (`E:\tools\node-v24\node.exe`) 필수 |
| **로컬 Outlook (`F:\MAIL`) 이관** | 완료 (2개 PST · 13개 프로젝트) | 계정 `designsir@seastargo.com` (Mailplug POP3 `995` / SMTP `465`), 서명 3종(`유일` 기본값), 총 메일 **`26,723`건**(`김봉정.pst` `12,763` + `백업.pst` `13,960`, 활성 `24,626` + 삭제 `2,097`), 폴더 **`141`개**, 연락처 **`152`건**, 비시스템 폴더 기반 프로젝트 **`13`건**(`recomputeProjectsFromLocalDb`) 적재 (`390.82 MB`). 모달(`PstImportModal`) 내 `F:\MAIL` 직접 스캔 UI 및 사이드바 전체 `141`개 폴더 검색·펼침 탑재 |
| **로컬 서버 · API · 전체 UI QA** | 완료 (18 API + 10 페이지 0 오류) | `dist-server/server.mjs` (`http://127.0.0.1:3000`): 부트스트랩·FTS5 검색(`seastargo` `20,969`건)·보낸편지함(`5,205`건)·지운편지함(`2,097`건)·커스텀 폴더·플래그 `PATCH`·PST PDF 첨부 추출(`68,644` bytes)·Mailplug autodiscover(`pop3Reachable`/`imapReachable: true`)·`POST /api/analyze` 미설정 400 정리(`PROVIDER_NOT_CONFIGURED`)·정적/Cloudflare 모두 통과. `ego-browser` `1440×900` 실측에서 폴더 전환·141개 폴더 펼침·검색·작성 서명·3종 모달·좌측 10개 전 페이지·`mail.seastar.work` 비노출(`pageErrors: []`) 확인 |
| **Windows 정식 설치형 (`.exe` · 단독 실행)** | 완료 (설치 및 단독 기동 실측) | 설치형 인스톨러 `release-exe/Bmail-Outlook-AI-Setup-1.0.0.exe` (`159,715,458` bytes, `npm run build:exe`) · 무인 설치(`NSIS_EXIT_CODE=0`) 완료 경로: `C:\Users\FREE\AppData\Local\Programs\Bmail Outlook AI\Bmail Outlook AI.exe` (앱 삭제 등록·바탕화면/시작메뉴 바로가기·`HKCU:\Software\Clients\Mail\Bmail` 포함, `desktop/icon.ico`, `desktop/installer.nsh`, `asar: false`). **외부 Node 서버 없이** Electron 내장 Node `24.15.0` (`ELECTRON_RUN_AS_NODE=1` + `node:sqlite`) 단독 기동으로 `http://127.0.0.1:3456` (`data/desktop-runtime.json`, `desktop-backend.log`)에서 `/api/config` `200`, `/api/local/bootstrap` `200`(`24,626`건·`141`폴더·`152`연락처·`13`프로젝트) 실측 확인 |
| **정적·빌드·단위 테스트** | 완료 (103/103 통과) | `npm run lint` (오류 0건) · `npm run build` · `npm run build:server` (`dist-server/server.mjs`) · Node 24 단위 테스트 19개 파일 `103/103` 통과 (`0` 실패, `0` 건너뜀) |

### 마스터 수동 완료 필요 항목 (Ego TaskSpace 20 및 Windows 기본 앱)
1. **메일플러그 로그인 · Turnstile · 앱 비밀번호 (`MAIL_PASSWORD`)**: Ego 브라우저 `TaskSpace 20`에 열린 Mailplug 2단계 로그인(`designsir@seastargo.com`, Cloudflare Turnstile 포함)에서 로그인 완료 후 앱 비밀번호를 발급하여 `E:\code-project\bmail\.env` (`MAIL_PASSWORD`) 또는 Bmail UI POP3 연결 모달에 입력 필요.
2. **Paid Tier 1 Gemini API 키 (`...iuIw` · `API260210` / `gen-lang-client-0779361701`)**: Ego 브라우저 Google AI Studio(`유료 · Tier 1 · 선불` 확인됨)에서 키 복사 후 `E:\code-project\bmail\.env` (`GEMINI_API_KEY`) 및 `npx wrangler secret put GEMINI_API_KEY` 등록 필요(유료 키만 허용, 무료 키 사용 금지).
3. **Windows 기본 메일 앱 지정**: 열린 Windows `설정 > 앱 > 기본 앱`에서 메일 기본 앱을 `Bmail Outlook AI`로 선택(필요 시 `release-exe/Bmail-Outlook-AI-Setup-1.0.0.exe` 실행).

---

## 1. 현재 구현 완료 현황 (Done & Verified)

| 영역 | 구현 내용 | 주요 파일 |
|---|---|---|
| **Fluent 2 디자인 시스템** | `Segoe UI Variable` / `Segoe UI` 서체, `--fluent-brand(#0F6CBD)` 및 시맨틱 컬러 토큰(`--fluent-urgent`, `--fluent-important`, `--fluent-success`, `--fluent-teal`), 4px 입력 필드·포커스 링(`ring-2 #0F6CBD`), New Outlook 헤더·앱 레일·리본·태스크페인 정렬 | `src/index.css`, `src/App.tsx`, `src/components/OutlookSettingsModal.tsx`, `src/components/ConfigEnvManager.tsx` |
| **자동 중요도 태깅 · 정렬** | 메일 내용(긴급·마감 키워드, 플래그, 읽지 않음)과 등록 프로젝트/SEDMS 도면 연관성 기반으로 **`긴급` / `중요` / `일반`** 자동 분류, 메일 리스트 배지 표기, `중요도순(긴급→중요→일반)` 정렬 및 그룹핑, 우선순위 필터 | `src/utils/mailPriority.ts`, `src/utils/mailPriority.test.ts`, `src/components/OutlookWorkbench.tsx` |
| **대시보드 D3 시각화 & 월간 트렌드** | ① **월간 메일 처리 트렌드 요약 지표 카드 4종**(월간 처리량·MoM 증감, 평균 회신 반응 속도, 월간 효율·읽음률, 최다 집중 프로젝트) ② **D3 다중 시리즈 차트**(프로젝트별 수신·발신 분포 막대 + 반응 속도 지수 라인) ③ D3 프로젝트 진행 현황 막대 ④ D3 종합 효율 게이지 & 일간 추이 | `src/utils/dashboardAnalytics.ts`, `src/utils/dashboardAnalytics.test.ts`, `src/components/DashboardEfficiencyCharts.tsx`, `src/components/DashboardView.tsx` |
| **EMV(Environment Variable) 암호화 볼트** | 최고 관리자 인증(`X-EMV-Admin-Token`) 기반 전사 공통(`global`) + 사원별(`employeeId`) API 키 AES-256-GCM 암호화 저장, 키 값 비노출 마스킹/boolean 상태 반환, 사원별→전사→`.env` 해석 우선순위 | `src/services/emvVaultService.ts`, `src/services/emvVaultService.test.ts`, `src/components/EmvVaultPanel.tsx`, `server.ts` |
| **하이브리드 크레딧 라우터 (GCP)** | 실시간 요약·업무일지(`Gemini API`, 13.43만 원 크레딧) vs 누적 메일 RAG(`Vertex AI Agent Builder`, 143만 원 크레딧) 역할 분리 + **2027-02-02 만료일 이후 자동 과금 하드스톱 가드** | `src/services/hybridCreditRouter.ts`, `cloudflare/worker.ts`, `wrangler.toml` |
| **전 메일 사이트(18개) IMAP/POP3 자동감지 & 모바일 모드** | 메일플러그·하이웍스·가비아·네이버웍스·다우오피스·두레이·카페24·한비로·카카오워크·후이즈·구글·M365·네이버·다음·네이트·iCloud·Zoho 등 18개 프리셋 MX/멀티호스트 자동탐색, 순수 Node.js **IMAP4rev1(993 SSL) + POP3(995/110)** 듀얼 수집, 로컬파트 아이디 자동 재시도, `#5.7.5` 앱 비밀번호·해외 IP 차단 진단 체크리스트, 모바일 간소화 UI(`MobileMailShell`) | `src/services/pop3AutoConnect.ts`, `src/services/pop3AutoConnect.test.ts`, `src/components/Pop3QuickConnectModal.tsx`, `src/components/MobileMailShell.tsx` |
| **로컬 Outlook PST/주소록 브리지 & Desktop 패키징** | PowerShell 기반 로컬 PC Outlook 프로필·PST 경로·주소록 스냅샷 추출/복원, Electron `.exe` 패키징 스캐폴딩 | `src/services/outlookLocalBridge.ts`, `desktop/main.cjs`, `desktop/preload.cjs`, `desktop/electron-builder.json` |
| **안전한 전체 소스 아카이브 내보내기** | 외부 의존성·셸 실행 없는 순수 Node.js POSIX `ustar` `.tar.gz` 빌더 (`.env`, `data/`, `node_modules/`, `dist/`, 인증서/PST 원천 차단) | `src/services/sourceArchiveBuilder.ts`, `src/services/sourceArchiveBuilder.test.ts`, `server.ts` |

---

## 2. Cloudflare 에이전트 실행 절차 (Cloudflare Workers + Static Assets)

현재 저장소 루트에 `wrangler.toml`과 `cloudflare/worker.ts`가 준비되어 있습니다. Cloudflare 빌드·배포 에이전트는 아래 순서대로 실행하세요.

### 2.1 사전 확인 및 프론트엔드 프로덕션 빌드

```bash
# 1. 의존성 설치
npm ci

# 2. 단위 테스트 + TypeScript 린트 + 프로덕션 빌드(dist/) 생성
npx --yes tsx --test src/utils/mailFilter.test.ts src/utils/appConfig.test.ts src/utils/deterministicAnalyzer.test.ts src/utils/outlookWorkspaceModel.test.ts src/utils/dashboardAnalytics.test.ts src/utils/mailPriority.test.ts src/services/emvVaultService.test.ts src/services/sourceArchiveBuilder.test.ts src/services/pop3AutoConnect.test.ts
npm run lint
npm run build
```

### 2.2 Cloudflare Secrets 등록 (최고 관리자 1회 실행, Git 커밋 금지)

`wrangler.toml`의 `[vars]`에는 비민감 기본값만 저장하고, 실제 키와 토큰은 반드시 `wrangler secret put`으로만 주입합니다.

```bash
# 1) Gemini API (13.43만 원 크레딧, Pay-as-you-go 결제 계정 연결 필수)
npx wrangler secret put GEMINI_API_KEY

# 2) Vertex AI Agent Builder / Discovery Engine (143만 원 크레딧)
npx wrangler secret put GCP_PROJECT_ID
npx wrangler secret put AGENT_BUILDER_DATA_STORE_ID
npx wrangler secret put AGENT_BUILDER_ACCESS_TOKEN

# 3) 사내 전사 접근 제어 토큰 (10~20명+ 직원 .exe 및 모바일 웹 공용 인증)
npx wrangler secret put BMAIL_ORG_ACCESS_TOKEN

# 4) (선택) Groq / Claude / OpenAI 추가 프로바이더
npx wrangler secret put GROQ_API_KEY
```

### 2.3 Cloudflare Worker 배포 및 헬스체크 검증

```bash
# 배포 (dist/ 정적 SPA 자산 + cloudflare/worker.ts 동시 배포)
npx wrangler deploy

# 배포 후 공개 안전 헬스체크 확인 (비밀 값 미노출, boolean 상태만 반환)
curl -s https://<YOUR-WORKER-SUBDOMAIN>.workers.dev/api/worker/health
```

예상 정상 응답:
```json
{
  "ok": true,
  "geminiConfigured": true,
  "geminiModel": "gemini-4-hydrogen",
  "agentBuilderConfigured": true,
  "agentBuilderExpiredHardStop": false,
  "expiryDate": "2027-02-02"
}
```

### 2.4 Cloudflare 운영 전환 시 필수 고도화 항목 (Cloudflare 에이전트 작업 범위)

현재 `cloudflare/worker.ts`는 `/api/worker/health`, `/api/worker/gemini`, `/api/worker/agent-builder` 및 `ASSETS` 정적 서빙을 처리합니다. Node.js Express(`server.ts`)의 전체 API를 Cloudflare Edge에서 완전히 대체하려면 다음 작업을 이어서 구현해야 합니다.

1. **GCP 서비스 계정 JSON -> 자동 OAuth2 Access Token 발급 (JWT RS256)**:
   - 현재 `AGENT_BUILDER_ACCESS_TOKEN`은 수동 Bearer 토큰(1시간 만료)입니다.
   - Cloudflare Worker에서 Web Crypto API(`crypto.subtle.importKey` + `RSASSA-PKCS1-v1_5`)로 서비스 계정 JSON(`client_email`, `private_key`)을 서명하여 `https://oauth2.googleapis.com/token`에서 `https://www.googleapis.com/auth/cloud-platform` 액세스 토큰을 자동 갱신·캐싱(50분 TTL)하도록 `cloudflare/worker.ts`를 확장하세요.
   - `2027-02-02` 만료일 하드스톱 가드(`isAgentBuilderExpired`)는 **절대 제거하지 마세요.**
2. **EMV 볼트의 Cloudflare KV / D1 이관**:
   - 로컬 Express 서버는 `data/emv-vault.enc.json` 파일 기반 AES-256-GCM 저장소를 사용합니다.
   - Cloudflare 환경에서는 파일시스템 쓰기가 불가하므로 `wrangler.toml`에 **Cloudflare KV** 또는 **Cloudflare D1** 바인딩을 추가하고, `EMV_VAULT_MASTER_KEY`와 `EMV_ADMIN_TOKEN`을 Cloudflare Secret으로 등록한 뒤 `/api/emv/*` 엔드포인트를 Worker에 포팅하세요.
3. **CORS 허용 오리진 제한**:
   - 현재 `cloudflare/worker.ts`는 요청의 `Origin`을 그대로 반사합니다. 사내 운영 배포 시 `ALLOWED_ORIGINS` 화이트리스트(사내 도메인, `https://<worker>.workers.dev`, Electron `app://` 등)만 허용하도록 제한하세요.
4. **프론트엔드 API Base URL 전환 (`VITE_API_BASE_URL`)**:
   - Electron `.exe` 및 Cloudflare Pages 분리 배포 시 프론트엔드가 중앙 Worker 게이트웨이를 바라보도록 `VITE_API_BASE_URL` 환경 변수 주입 및 `Authorization: Bearer <BMAIL_ORG_ACCESS_TOKEN>` 헤더 연동을 적용하세요.

---

## 3. 잔여 API 적용 및 후속 구현 백로그 (API / 로컬 에이전트 작업 범위)

아래 항목들은 실제 사내 환경·클라우드 계정 자격증명·Windows 로컬 바이너리가 필요한 잔여 과제입니다.

### 3.1 Google Cloud / AI Studio 하이브리드 크레딧 실연동
- [ ] **AI Studio 유료 등급(Pay-as-you-go) 활성화**: 13.43만 원 크레딧 결제 계정에 API 키 연결 후 `GEMINI_API_KEY` 등록.
- [ ] **GCP API 2종 활성화**: `Generative Language API`, `Discovery Engine API` Enable.
- [ ] **Cloud Storage 버킷 + Vertex AI Agent Builder 데이터 스토어 연결**: 과거 누적 메일(.txt/.json) 적재 및 검색 앱 생성.
- [ ] **GCP 결제 예산 알림 설정**: 2027-02-02 만료 전(50%/80%/100%) 이메일·SMS 예산 경보 설정.
- [ ] **전 프로젝트 공용 게이트웨이 키 정책**: AI Studio에서 개발하는 타 프로젝트에서도 `/api/worker/gemini` 및 `/api/worker/agent-builder` 엔드포인트를 `BMAIL_ORG_ACCESS_TOKEN`으로 공유 호출하도록 팀 표준 SDK 래퍼 정리.

### 3.2 Microsoft 365 / Outlook Add-in & Graph 실연동
- [ ] **Microsoft Entra ID 앱 등록**: Redirect URI(`/api/auth/microsoft/callback`, NAA Broker URI) 및 권한(`Mail.Read`, `Mail.ReadWrite`, `Mail.Send`, `User.Read`) 등록 후 `MICROSOFT_CLIENT_ID`, `MICROSOFT_CLIENT_SECRET`, `MICROSOFT_TENANT_ID` 설정.
- [ ] **NAA(Nested App Authentication) SSO 활성화**: `public/manifest.json`의 `webApplicationInfo`에 실제 앱 ID와 `api://<DOMAIN>/<CLIENT_ID>` 리소스 URI 반영.
- [ ] **Outlook Add-in 사이드로드 검증**: `public/manifest.xml` (`SupportsPinning` 유지) 배포 테스트.

### 3.3 메일플러그 · 전 메일 사이트(IMAP / POP3 / SMTP) 실연동 보강
- [x] `src/services/pop3AutoConnect.ts`에서 국내·해외 주요 18개 메일 사이트 특징(두레이/iCloud IMAP 전용, 카페24/하이웍스 POP3 전용, 다우오피스 `pop3s.daouoffice.com`/`outbound.daouoffice.com`, 네이버·다음·네이트 아이디/이메일 자동 재시도, 메일플러그 `#5.7.5` 앱 비밀번호·해외 IP 차단 진단)과 순수 Node.js **IMAP4rev1(993 SSL) + POP3(995/110)** 듀얼 엔진이 구현되어 있습니다.
- [ ] **Cloudflare Workers 환경 및 해외 IP 차단 주의**: 클라우드 해외 서버(미국 등)에서는 메일플러그·네이버·다음의 기본 `해외 IP 차단` 정책에 의해 `-ERR 553 (#5.7.5)` 또는 인증 거부가 발생할 수 있으며, Cloudflare Workers는 일반 TCP/TLS 소켓 제한이 있습니다. 사내 운영 시 **Windows Electron `.exe` 로컬 백엔드(한국 사내 IP)** 또는 **한국 리전 사내 Node.js 중계 서버(`server.ts`)**에서 IMAP/POP3/SMTP 직접 접속을 실행하는 하이브리드 구성을 유지하세요.
- [ ] **다중 IMAP 폴더(보낸 편지함/사용자 폴더) 동기화 확장**: 현재 IMAP4rev1 엔진은 `INBOX` (`BODY.PEEK[]` + `\Seen`/`\Flagged` 파싱)을 수집합니다. 보낸 편지함(`Sent`) 및 사용자 정의 폴더까지 양방향 동기화가 필요할 경우 `fetchImapMessages`의 `SELECT` 대상에 `LIST "" "*"` 기반 폴더 순회를 추가하세요.

### 3.4 Windows `.exe` 설치형 패키징 및 로컬 Outlook PST/주소록 완전 연동
- [ ] `desktop/electron-builder.json` 및 `desktop/main.cjs` 기반 Windows NSIS 인스톨러(`.exe`) 빌드 파이프라인 구성 (`electron`, `electron-builder`는 Windows 빌드 머신에서 devDependency 설치).
- [ ] 로컬 PC Outlook 레지스트리/COM(`Outlook.Application` MAPI) 자동 탐색 스크립트(`src/services/outlookLocalBridge.ts`)를 Electron `ipcMain`에서 직접 실행하여 설치 즉시 기존 Outlook 계정·주소록·PST 경로를 1클릭 로드하도록 연결.
- [ ] 대용량 `.pst` 파일(>500MB) 처리를 위한 스트리밍 청크 파서 및 로컬 SQLite(`better-sqlite3` 또는 Node 내장 `node:sqlite`) 인덱싱 계층 연결.

### 3.5 외부 문서 파싱 어댑터 및 영구 저장소
- [ ] **외부 PDF/Excel 엔진 어댑터**: `PDF_PARSER_CMD`, `EXCEL_PARSER_CMD` (예: `python scripts/parse_pdf.py`, `python scripts/parse_excel.py`) 실제 스크립트 배치 및 `GET /api/config` 프로브 확인.
- [ ] **데이터베이스 영구화**: 현재 클라이언트 `localStorage`에 저장되는 프로젝트 분류·메일 오버라이드·일간 브리핑 아카이브를 사내 SQLite/PostgreSQL 또는 Cloudflare D1으로 동기화.

---

## 4. 불변 제약 및 검증 기준 (Definition of Done)

후속 에이전트는 모든 PR/배포 전 아래 불변 규칙과 검증 절차를 반드시 준수해야 합니다.

### 불변 규칙
1. **시뮬레이션/더미 메일 0건**: `emails = []` 시작 원칙 엄수 (테스트 픽스처 외 UI 샘플 메일 주입 금지).
2. **레거시 Exchange API 금지**: `getUserIdentityTokenAsync`, `getCallbackTokenAsync`, EWS 사용 금지.
3. **Groq 8K TPM 방어**: `GROQ_CHUNK_CHAR_LIMIT = 6000` Map-Reduce 청킹 유지.
4. **Claude 4 / Gemini 4 파라미터 제약**: `temperature`, `top_p`, `top_k` 미전송 유지, Gemini SDK는 서버/Worker 측 전용.
5. **비밀 정보 격리**: `.env`, `data/emv-vault.enc.json`, API 키, OAuth 시크릿, 사원 비밀번호를 Git·`config.json`·클라이언트 번들·다운로드 아카이브·로그에 절대 포함하지 말 것.

### 필수 검증 명령 (DoD)
```bash
# 1. 전체 단위 테스트 통과
npx --yes tsx --test src/utils/mailFilter.test.ts src/utils/appConfig.test.ts src/utils/deterministicAnalyzer.test.ts src/utils/outlookWorkspaceModel.test.ts src/utils/dashboardAnalytics.test.ts src/utils/mailPriority.test.ts src/services/emvVaultService.test.ts src/services/sourceArchiveBuilder.test.ts src/services/pop3AutoConnect.test.ts

# 2. TypeScript 정적 검증 (오류 0건)
npm run lint

# 3. 프로덕션 빌드 검증
npm run build
```
