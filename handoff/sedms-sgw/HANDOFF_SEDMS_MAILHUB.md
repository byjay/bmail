지시서 20261120-MAILHUB 읽음

# SEDMS 메일 허브(MAILHUB) 실행 핸드오프 패키지 — Bmail → SEDMS(`E:/code-project/SEDMS_G`) 이관서

- **작성일**: 2026-11-20
- **발신**: Bmail (Outlook 모방 메일 AI 앱) 담당 에이전트
- **수신**: SEDMS (`E:/code-project/SEDMS_G`, 기준 커밋 `561e994`, 브랜치 `test/int-1114` 기반 `feat/mailhub-m<n>`) 담당 에이전트
- **보안 전제**: 본 문서·SQL·테스트에 특정 호선·회사·조선소·직원 실명·비밀번호·키 평문을 일절 포함하지 않음(S4, S5 준수).

---

## 1. Bmail 저장소에서 이미 구현·검증 완료된 모듈 (직접 이식 가능)

아래 순수 TypeScript 모듈은 외부 의존성 없이 Node/Workers(Hono) 호환으로 구현 및 단위·뮤테이션 테스트 검증 완료되어 있으며, `E:/code-project/SEDMS_G`로 그대로 이식 가능합니다.

| 모듈(Bmail 경로) | SEDMS 이식 대상 | 검증 완료 내용 |
|---|---|---|
| `src/utils/sedmsMailHub.ts` | `src/server/services/mailHubMatcher.ts` | ① `normalizeAlphaNumeric`·`canonDrawingNo`·`createFlexibleDrawingPattern` (`drawingNoKey.ts` 호환)<br>② 결정적 5순위 매칭 엔진(규칙 1~5, 규칙 2·3 충돌 시 자동 대기열, 회사 `tenant_id` 완전 격리, DB/설정 기반 토큰 전용)<br>③ 테넌트별 서버 간 HMAC-SHA256 서명/검증(`X-SEDMS-Tenant`, `X-SEDMS-Timestamp`, `X-SEDMS-Signature`, `timingSafeEqual`)<br>④ 멱등 수집(`UNIQUE(tenant_id, message_id)`) + 첨부 `sha256` 중복 방지<br>⑤ NAS 추가 저장 경로 생성기(`_mail/YYYYMM/<safe-message-id>/<sha12>_<filename>`, S7 준수)<br>⑥ 출도 발신 헤더 생성기(`Message-ID`, `X-SEDMS-Tenant/Project/Distribution`)<br>⑦ M5 봉인 도면 「앞으로만」 가드(역행/동일 개정 409 반환, S6 준수) |
| `src/utils/sedmsMailHub.test.ts` | `tests/mailhub/` | 합성 메일 30건 표준 세트(규칙 1~5 + 충돌 케이스), 타 테넌트 후보 0건, 3회/동시 수집 멱등성, `sha256` 중복 0, 봉인 도면 역행 409, HMAC 인증 401/403/200, 뮤테이션 검증 |
| `handoff/sedms-sgw/migrations/0106_mailhub_core.sql` | `migrations/0106_mailhub_core.sql` | `mail_accounts`, `mail_messages`, `mail_attachments` (D1 호환) |
| `handoff/sedms-sgw/migrations/0107_mailhub_links_queue.sql` | `migrations/0107_mailhub_links_queue.sql` | `mail_links`, `mail_review_queue` |
| `handoff/sedms-sgw/migrations/0108_mailhub_outbound_tracking.sql` | `migrations/0108_mailhub_outbound_tracking.sql` | 기존 배포 발급 테이블 `outbound_message_id` 컬럼 추가 |
| `handoff/sedms-sgw/formal/mailhub-ingest/MailHubIngest.tla` (+ `.cfg`) | `formal/mailhub-ingest/` | TLA+ 동시 수집·재시도·커서 진행 형식 명세 (`NoDuplicateMessages`, `NoDataLossOnCursorAdvance` 불변식) |

---

## 2. SEDMS 저장소(`E:/code-project/SEDMS_G`)에서만 수행 가능한 이관 작업 (M0~M7)

이 워크스페이스는 Bmail 저장소이므로, 아래 SEDMS 본체 작업은 SEDMS 전용 worktree(`feat/mailhub-m0` ~ `m7`)에서 순서대로 실행하세요.

### M0: 계정·시크릿·마이그레이션 0106·서버 간 HMAC 인증
1. `handoff/sedms-sgw/migrations/0106_mailhub_core.sql`을 SEDMS `migrations/`에 복사 후 실 D1 miniflare 테스트 적용.
2. Hono 라우터에 `POST /api/mail/ingest` 추가 (`verifySedmsIngestSignature` + 테넌트별 시크릿 `env.SEDMS_INGEST_HMAC_SECRETS_JSON` 또는 Secrets Binding).
3. 실 D1 테스트: 서명 없음(401) / 서명 불일치(403) / 다른 회사 키(403) / 정상 서명(200).

### M1: 사내 메일 수집기 (`workers/mail-collector` 또는 `scripts/mail-collector`) + NAS/R2 어댑터
1. 읽기 전용 어댑터 2종 구현:
   - M365 Graph (`GET /v1.0/users/{sharedMailbox}/mailFolders/{folderId}/messages/delta`) — **기본값: 공용 메일함 + 지정 폴더만(Q2)**
   - IMAP (`imapflow` 등 MIT/BSD 라이브러리, `readOnly: true` 플래그 강제)
2. 원문(`.eml`)·본문(`.txt`)·첨부파일을 **NAS 저장소 어댑터(내부) / R2(외부 SaaS)**에 **추가 저장 전용(S7)**으로 먼저 기록(`buildNasMailStorageKey` 경로 규칙 `_mail/YYYYMM/<msg-id>/<sha12>_<filename>`). 원본 이동·삭제 절대 금지.
3. 저장 완료 후 메타데이터 + `storage_key`만 HMAC 서명하여 `POST /api/mail/ingest` 전송.
4. 메일 서버 장애·저장소 미연결 시 `ok: true`로 삼키지 말고 5xx + 사유 코드 기록, 커서 미진행.
5. `formal/mailhub-ingest/MailHubIngest.tla`를 TLC로 실행하여 상태 폭발 없이 `NoDuplicateMessages`, `NoDataLossOnCursorAdvance` 통과 로그 확보.

### M2: 결정적 5순위 매칭 엔진 + 마이그레이션 0107
1. `migrations/0107_mailhub_links_queue.sql` 적용.
2. `src/utils/sedmsMailHub.ts`의 `matchMailDeterministically`를 SEDMS 기존 `drawingNoKey.ts`, `companies`(0092), `recipient_groups`(0094), 프로젝트 `vessel_aliases` DB 쿼리에 연결(S4: 하드코딩 리터럴 0건).
3. `GET /api/projects/:id/mail`, `GET /api/mail/:msgId`, `GET /api/mail/review`, `POST /api/mail/:msgId/link` 라우트 구현(`tenantOf` + `projectScopeGuard` + `requireServerRole` 전부 적용).
4. 격리 3종 테스트(다른 회사 404 · 승인 대기 403 · 권한 부족 403) + 30건 표준 세트 100% 일치 + 뮤테이션 4종(서명 제거/message_id 유일 제거/tenant 조건 제거/충돌 대기열 우회) 검증.

### M3: SEDMS 화면 연동 (M-41 분할창 · M-44 그리드 · M-48 아이콘 · M-52 글 최소)
1. 도면 선택 시 오른쪽 분할 창에 「메일」 탭 추가(연결 스레드, 수신/발신 아이콘, 날짜, 업체, 첨부 수).
2. 프로젝트 화면에 「메일」 탭(전체 스레드 + 미분류 배지 수) 추가.
3. 확인 대기열 그리드 화면(모달 금지): 후보 프로젝트·규칙·점수 표시 → 인라인 확정/변경/거부.
4. 라이트·다크 두 테마 대비 4.5:1 검증 및 헤드리스 캡처.

### M4: 출도 메일 강화 + 마이그레이션 0108
1. `migrations/0108_mailhub_outbound_tracking.sql`의 대상 테이블명을 SEDMS 실제 배포 발급 테이블명과 대조 후 적용.
2. `src/server/services/emailDispatchService.ts`에서 발송 시 `buildSedmsOutboundDispatchHeaders` 호출 → `Message-ID`, `X-SEDMS-Tenant`, `X-SEDMS-Project`, `X-SEDMS-Distribution` 헤더 부여 및 `outbound_message_id` DB 저장.
3. 출도 발신 → 회신(`In-Reply-To`) 합성 왕복 테스트로 규칙 1 자동 연결 확인.

### M5: 첨부 → 도면 대장 등록 (`POST /api/mail/:msgId/attach-to-drawing`)
1. 첨부 행의 「도면에 붙이기」 버튼 → 정규화 도면번호 자동 제안 → 설계자 이상 권한자 확인 후 기존 SEDMS 도면 업로드·개정·봉인 파이프라인 호출(S6: 자동 변경 금지).
2. 봉인 도면 역행 개정 첨부 시 기존 규칙대로 HTTP `409` 반환 확인.

### M6: Bmail 기능 대응표(빈칸 0) 및 퇴역 절차
| Bmail 기존 기능 | SEDMS Mail Hub 대응 위치 | 상태 |
|---|---|---|
| Outlook 3열 메일함·스레드·검색 | M3 프로젝트/도면 분할창 「메일」 탭 + SGW Bmail 메인메뉴 | Bmail 구현 완료 / SEDMS M3 이식 |
| PST/EML/MBOX 수집 및 첨부 추출 | M1 `mail-collector` (M365 Graph + IMAP 읽기 전용) | Bmail 파서 완료 / SEDMS M1 워커 이식 |
| 프로젝트 코드·도면번호 자동 분류 | M2 `matchMailDeterministically` (규칙 1~5 + `drawingNoKey.ts`) | 로직·30건 테스트 완료 / SEDMS D1 연결 |
| 첨부파일 분석·원본 보관 | M1 NAS/R2 추가 저장 + M5 `attach-to-drawing` | 경로 규칙·봉인 가드 완료 / SEDMS NAS 어댑터 연결 |
| 출도 메일 발송 및 회신 추적 | M4 `emailDispatchService.ts` + `outbound_message_id` | 헤더 생성기·규칙 1 완료 / SEDMS M4 연결 |
| AI 요약·초안·용어사전 | SEDMS `aiEgress` 관문 + 회사별 기본 꺼짐 (`sgwAiService.ts` 패턴) | Bmail 비공개 AI 가드 완료 / SEDMS `aiEgress` 연결 |
| 확인 대기열(오매칭 방어) | M2/M3 `mail_review_queue` 그리드 | Bmail 구현 완료 / SEDMS M3 이식 |

*퇴역(Q4): M6 대조표 실측 검증 통과 및 마스터 승인 후 진행.*

### M7: 매뉴얼 및 가드 최종 통과
- SEDMS 전체 가드 통과 확인: `tsc 0`, `build 0`, `project-literals`, `color-literals`, `ui-verbosity`, `no-emoji`, `scan-orphans`, `schema-isolation` 기준선 상향 금지.
- 운영 배포·운영 D1 마이그레이션·실제 메일 계정 연결은 마스터 승인 후 실행.
