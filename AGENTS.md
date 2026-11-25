# AGENTS.md — 전역 코딩 및 엔지니어링 지침 (MailLog Studio for Outlook)

본 저장소에서 작업하는 모든 AI 에이전트(로컬 에이전트 및 코드 어시스턴트)와 개발자는 아래 7대 전역 원칙을 최우선 규범으로 준수한다.

## 1. 전역 코딩 및 엔지니어링 7대 원칙

1. **과잉 강조 배제 및 간결성**: 불필요한 수식어, 과장된 칭찬, 코드에 이미 드러난 중복 설명을 생략하고 핵심 사실과 변경점만 제시한다.
2. **최소 의존성**: 동일 기능 구현 시 외부 라이브러리 추가를 최소화하고, Node.js/웹 표준 API와 간결한 구조를 우선한다. 신규 패키지 도입 시 라이선스(MIT/BSD/Apache-2.0 우선, GPL/AGPL 배제)와 유지보수 상태를 검증한다.
3. **방어적 설계**: 정상 경로보다 실패 경로(경계값, `null`/`undefined`, 동시성, 타임아웃, 임시 파일/메모리 누수)를 먼저 설계하고, 예외를 빈 `catch`로 삼키지 않고 명확한 컨텍스트와 함께 로깅·전파한다.
4. **숨은 의존성 보호**: 미사용 코드처럼 보여도 레거시·외부(Office.js, 외부 PDF/Excel 어댑터, VSTO 브리지 등) 참조 가능성을 먼저 경고하고 임의 삭제를 제안하지 않는다.
5. **3단계 자체 검증**: 코드 변경 출력 전 반드시 아래 순으로 교차 검증한다.
   - 1단계(구문): TypeScript 타입 및 구문 오류 없음 (`npm run lint`)
   - 2단계(로직·엣지 케이스): 경계값, 빈 입력, 잘못된 JSON/파일, 타임아웃 단위 테스트 통과
   - 3단계(부수 효과): 기존 API 계약, `config.json`/`.env` 분리, `SupportsPinning` 등 하위 호환성 유지
6. **보안 원칙**: `.env`, API 키, OAuth 시크릿, 메일 비밀번호, 인증서 등 민감 정보는 절대 하드코딩·`config.json` 저장·클라이언트 노출·로그 출력하지 않는다. 서버 환경 변수(`.env`)와 boolean 상태 노출 원칙을 유지한다.
7. **작업 분할 및 검증 기준 (DoD)**: 복잡한 구현은 완료 기준(Definition of Done)과 함께 원자적 단위로 나누어 수행하며, 매 변경 종료 시 아래 검증 절차를 통과해야 한다.

## 2. 저장소 필수 검증 절차 (DoD)

```bash
# 1. 단위 테스트 통과
npx --yes tsx --test src/utils/mailFilter.test.ts src/utils/appConfig.test.ts src/utils/deterministicAnalyzer.test.ts

# 2. TypeScript 정적 검증 (오류 0건)
npm run lint

# 3. 프로덕션 빌드 검증
npm run build
```

## 3. 프로젝트 불변 제약

- **시뮬레이션 데이터 0건**: 샘플/더미 메일 생성 금지 (`emails = []` 시작).
- **레거시 Exchange API 금지**: `getUserIdentityTokenAsync`, `getCallbackTokenAsync`, EWS 사용 금지.
- **Groq 8K TPM 방어**: `GROQ_CHUNK_CHAR_LIMIT = 6000` Map-Reduce 청킹 유지.
- **Claude 4 / Gemini 4 파라미터 제약**: `temperature`/`top_p`/`top_k` 미전송 유지, Gemini SDK는 서버 측 전용.

## 4. 기능 업데이트 시 인터랙티브 매뉴얼 자동 동기화 규약 (`Model: "flash"`)

- 새로운 기능 추가, 버전 범프(`package.json`), 단축키/UI/API 변경이 발생할 때마다 메인 에이전트는 **반드시 `Model: "flash"` 서브에이전트**를 호출하여 아래 두 파일을 최신 상태로 동기화해야 한다:
  1. `src/data/manualCatalog.ts` — `MANUAL_APP_VERSION`을 `package.json`과 일치시키고 `MANUAL_FEATURE_CHANGELOG`, `MANUAL_MOTION_SCENES`, `MANUAL_UI_ICONS`(전체 아이콘 번호·위치·엔드포인트·팁)를 누락 없이 갱신한다.
  2. `src/components/InteractiveManualModal.tsx` — 모든 아이콘 목록(#1~전체)이 잘리거나 8번까지만 표시되지 않도록 전체 목록을 즉시 탐색 가능하게 유지하고, 모션 스텝·목업 프리뷰와 카탈로그가 항상 1:1 일치하도록 검증한다 (`npx --yes tsx --test src/data/manualCatalog.test.ts`).
