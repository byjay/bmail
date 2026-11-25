# SGW(콘트롤타워·관문) 연동 요청서 및 계약 핸드오프 — Bmail 메인메뉴 통합

- **작성일**: 2026-11-20
- **발신**: Bmail (사내메일 AI 전용앱) 담당 에이전트
- **수신**: SGW(콘트롤타워·관문) 프로젝트 팀장 세션
- **원칙**: SGW 저장소 코드는 Bmail 담당이 직접 수정하지 않음(요청서 §6.4). 아래 인터페이스 계약을 SGW 측에 반영해 주시면 즉시 연동됩니다.

---

## 1. Bmail 측 준비 완료 항목 (즉시 연동 가능)

1. **SGW 메인메뉴 / 확장앱 사이드패널 진입 경로**:
   - 전체 화면(SGW 메인메뉴 탭/iframe): `https://<BMAIL_HOST>/sgw/mail?embed=sgw` (또는 `/?embed=sgw`)
   - 확장앱 사이드패널(좁은 폭 `<= 480px`): 동일 URL에서 반응형 1열/2열 자동 전환
   - `?embed=sgw` 진입 시 Outlook 외부 크롬·플러그인 설치 배너·PST 안내를 자동 숨김하고 SGW 사내메일 본체로 표시
2. **SGW SSO HMAC-SHA256 토큰 교환 API**:
   - `POST /api/sgw/auth/exchange` (Body: `{ "token": "<base64url(payloadJson)>.<hexHmacSha256>" }`)
   - 또는 진입 URL 쿼리 `?embed=sgw&sso_token=<token>` → 서버가 `SGW_SSO_SECRET`으로 상수 시간 검증(`crypto.timingSafeEqual`) + 만료(`exp`) + 재사용(`nonce`) 차단 후 `HttpOnly; Secure; SameSite=None` 세션 쿠키(`sgw_session_id`) 발급
   - 페이로드 규격:
     ```json
     {
       "sub": "user_unique_id",
       "name": "표시이름",
       "department": "부서명(선택)",
       "email": "user01@example.test",
       "role": "user | manager | admin",
       "exp": 1795160000,
       "nonce": "uuid-v4"
     }
     ```
3. **세션·신원 통일 (`sub` 기준)**:
   - 모든 메일·수신자·권한·첨부 서명 URL·마이그레이션(`senderId`/`recipientId`)은 SGW `sub`로 통일.
   - 자체 비밀번호 DB 없음.
   - `GET /api/sgw/auth/me`는 Bmail 세션 확인 및 `SGW_BASE_URL/auth/me` 대조 지원.
4. **직원 목록 프록시 (`GET /api/sgw/employees`)**:
   - `${SGW_BASE_URL}/api/employees` 실시간 호출 + 60초 메모리 TTL 캐시(DB 복제 금지).
5. **권한·보안 6대 시나리오 자동 테스트 통과**:
   - ① 타인 메일 조회 불가(403) ② 타인 명의 발송 불가(서버 `session.sub` 강제) ③ 수신자만 읽음 변경(403) ④ 당사자만 삭제 + 휴지통→영구삭제 2단계 ⑤ 프롬프트 주입에도 자동 발송/삭제/전달 0건 ⑥ AI 기본값 꺼짐 + 꺼짐 시 AI 호출 0회 + 외부 퍼블릭 LLM 호스트 원천 차단.
6. **`internal_mail` 이전 도구**:
   - `POST /api/sgw/migration/dry-run` (건수·중복·누락·UTF-8 인코딩·자가발송·ID 매핑 미리보기)
   - `POST /api/sgw/migration/execute` (`approved: true` 필수, ID 기준 대조 리포트 반환 — 실명 미포함)
   - `POST /api/sgw/migration/rollback` (이전 데이터만 안전 롤백)

---

## 2. SGW 측 반영 요청 사항

| # | 항목 | 요청 내용 |
|---|---|---|
| 1 | **메인메뉴 "메일" 연결** | SGW 메뉴의 "메일" 클릭 시 단기 HMAC SSO 토큰을 발급하여 `https://<BMAIL_HOST>/sgw/mail?embed=sgw`로 진입(iframe 또는 새 탭). |
| 2 | **CSP / iframe 허용 헤더** | SGW 호스트 도메인을 Bmail `.env`의 `SGW_FRAME_ANCESTORS`에 등록할 수 있도록 운영·스테이징 SGW 오리진 전달. |
| 3 | **공유 시크릿 전달** | `SGW_SSO_SECRET`(32바이트 이상 랜덤 hex)을 `.env`/시크릿 저장소로 공유(문서·커밋 평문 금지). |
| 4 | **`/auth/me` & `/api/employees` 확인** | 응답 필드에 `sub`(또는 `id`), `name` 포함 여부 및 인증 전달 방식(Cookie / Bearer) 확인. |
| 5 | **비공개 AI 엔드포인트 & 용어 목록** | `SGW_PRIVATE_AI_BASE_URL`(사내 로컬/비공개 LLM), 초기 조선·선박 전장 용어 JSON/CSV, 첨부 크기(`SGW_ATTACHMENT_MAX_MB`) 및 허용 확장자(`SGW_ATTACHMENT_ALLOWED_EXT`) 정책 전달. |
| 6 | **기존 `/api/mail` 단계적 전환 (§4.4)** | ① `internal_mail` D1 덤프(JSON) → Bmail 드라이런 검수 → ② 승인 후 이전 실행 및 건수·샘플(ID 기준) 대조 → ③ SGW `POST/PUT/DELETE /api/mail` 읽기 전용 전환 → ④ 일정 기간 병행 운영 후 마스터 승인으로 제거. |
