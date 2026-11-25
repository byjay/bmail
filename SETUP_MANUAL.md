# Bmail — Outlook AI 워크스페이스 완벽 설정 · 설치 · 빌드 매뉴얼

> **문서 대상**: 사내 IT 최고 관리자, 개발자, Cloudflare 배포 담당자 및 일반 사원(10~20명+ 전사 배포)
> **프로젝트명**: Bmail (`MailLog Studio for Outlook`)
> **포함 범위**: 로컬 실행 · 18개 메일 서비스(POP3/IMAP/SMTP) 연결 · EMV 암호화 볼트 · AI/GCP 크레딧 · Windows `.exe` 패키징 · Cloudflare Workers 배포 · 장애 해결 FAQ

---

## 목차

1. [사전 준비 및 로컬 빌드·실행 (Windows / macOS / Linux)](#1-사전-준비-및-로컬-빌드실행)
2. [국내·해외 18개 메일 서비스별 POP3 · IMAP · SMTP 설정 매뉴얼](#2-국내해외-18개-메일-서비스별-pop3--imap--smtp-설정-매뉴얼)
3. [메일플러그(Mailplug) `#5.7.5` 완벽 해결 가이드 (필독)](#3-메일플러그mailplug-575-완벽-해결-가이드-필독)
4. [사내 전사 EMV(Environment Variable) 볼트 및 사원별 API 키 배포](#4-사내-전사-emv-볼트-및-사원별-api-키-배포)
5. [AI 엔진 · GCP 하이브리드 크레딧(2027-02-02 자동 하드스톱) 설정](#5-ai-엔진--gcp-하이브리드-크레딧-설정)
6. [Microsoft 365 OAuth 및 Outlook Add-in(manifest.xml) 설치](#6-microsoft-365-oauth-및-outlook-add-in-설치)
7. [로컬 PC Outlook PST · 주소록 자동 인수 및 Windows `.exe` 설치형 빌드](#7-로컬-pc-outlook-pst--주소록-자동-인수-및-windows-exe-설치형-빌드)
8. [Cloudflare Workers + Static Assets 배포 절차](#8-cloudflare-workers--static-assets-배포-절차)
9. [검증 체크리스트(DoD) 및 장애 대응 FAQ](#9-검증-체크리스트dod-및-장애-대응-faq)

---

## 1. 사전 준비 및 로컬 빌드·실행

> **중요**: 클라우드 미리보기 서버는 해외 IP(대만 `TW` 등)에서 실행되므로 국내 기업메일(메일플러그·네이버·다음 등)의 `[해외 IP 차단]` 정책에 걸릴 수 있습니다. **한국 사내 PC에서 로컬 실행 또는 Windows `.exe`로 실행하면 해외 IP 차단 없이 즉시 연결됩니다.**

### 1.1 필수 소프트웨어
- **Node.js**: `v20.18+` 또는 `v22 LTS` (권장) — [https://nodejs.org](https://nodejs.org)
- **npm**: `v10+` (Node.js 설치 시 자동 포함)
- *(선택)* **Git** 및 **Cloudflare Wrangler CLI** (`npx wrangler`)

### 1.2 압축 해제 및 환경 변수(`.env`) 준비

```bash
# 1) 다운로드한 소스 압축 해제 후 폴더 이동
unzip bmail-outlook-full-source-YYYY-MM-DD.zip
cd bmail-outlook-source

# 2) 환경 변수 템플릿 복사 (.env 파일은 절대 Git 커밋 금지)
# Windows PowerShell:
Copy-Item .env.example .env
# macOS / Linux:
cp .env.example .env
```

`.env` 파일에 최소 필수 값을 입력합니다 (상세 키는 4·5장 참조):

```env
PORT=3000
APP_URL=http://localhost:3000

# EMV 최고 관리자 비밀번호 및 암호화 마스터 키 (32바이트 이상 무작위 문자열 권장)
EMV_ADMIN_TOKEN=change-this-to-strong-admin-password
EMV_VAULT_MASTER_KEY=change-this-to-32-plus-char-random-secret-key

# 기본 AI 엔진 (Gemini API 또는 Groq/Claude/OpenAI 중 택1 이상)
GEMINI_API_KEY=your-gemini-api-key
GROQ_API_KEY=
ANTHROPIC_API_KEY=
OPENAI_API_KEY=
```

### 1.3 의존성 설치 · 검증 · 실행

```bash
# 1) 의존성 설치
npm install

# 2) TypeScript 정적 검사 (오류 0건 확인)
npm run lint

# 3) 프로덕션 프론트엔드 빌드 (dist/ 생성)
npm run build

# 4) 개발 서버 실행 (포트 3000)
npm run dev
```

브라우저에서 **`http://localhost:3000`** 을 열면 Bmail Outlook 워크스페이스가 즉시 실행됩니다.

---

## 2. 국내·해외 18개 메일 서비스별 POP3 · IMAP · SMTP 설정 매뉴얼

앱 상단의 **플러그 아이콘(`회사 POP3 메일 연결`)**을 누르고 **서비스 형태 + 이메일 + 비밀번호(또는 앱 비밀번호)**를 입력하면 아래 18개 서비스의 서버 주소·포트·TLS 정책을 자동 분석하고, 실패 시 자동 해결 엔진(최대 4단계: 권장 프로토콜 → 앱 비밀번호 공백/하이픈 자동 제거 → 반대 프로토콜 전환 → 도메인 전용 서버)이 동작합니다.

| 메일 서비스 | 지원 프로토콜 | 권장 | 수신 서버 : 포트 (SSL) | 발신 SMTP : 포트 | 비밀번호 정책 | 필수 주의사항 |
|---|---|---|---|---|---|---|
| **메일플러그 (Mailplug)** | POP3 · IMAP | **POP3** | `pop3.mailplug.co.kr:995` (SSL)<br>`imap.mailplug.co.kr:993` (SSL) | `smtp.mailplug.co.kr:465` (SSL) | **앱 비밀번호 필수** | **2025.04.01 이후 개통 계정 IMAP 제한** · 30일 미사용 시 자동 OFF · 해외 IP 차단 확인 |
| **하이웍스 (Hiworks)** | **POP3 전용** | **POP3** | `pop3s.hiworks.com:995` (SSL) | `smtps.hiworks.com:465` (SSL) | 전용 비밀번호 권장 | IMAP 미지원 · 환경설정에서 POP3 사용함 체크 필수 |
| **가비아 기업메일** | **POP3 전용** | **POP3** | `pop3.gabia.com:995` (SSL)<br>또는 `pop3.{도메인}:995` | `smtp.gabia.com:465` (SSL) | 일반/전용 비밀번호 | 웹메일 환경설정 > POP3 사용 허용 필요 |
| **네이버 웍스 (NAVER WORKS)** | IMAP · POP3 | IMAP | `imap.worksmobile.com:993` (SSL)<br>`pop.worksmobile.com:995` (SSL) | `smtp.worksmobile.com:465` (SSL) | **앱 비밀번호 필수** | 관리자 콘솔에서 외부 IMAP/POP3 허용 + 외부 앱 비밀번호 발급 |
| **다우오피스 (DaouOffice)** | IMAP · POP3 | IMAP | `imap.daouoffice.com:993` (SSL)<br>`pop3s.daouoffice.com:995` (SSL)<br>*(구축형 `mail.{도메인}`)* | `outbound.daouoffice.com:465` (SSL) | 일반/전용 비밀번호 | 클라우드형은 `pop3s.daouoffice.com` / 발신 `outbound.daouoffice.com` 사용 |
| **두레이 (NHN Dooray)** | **IMAP 전용** | **IMAP** | `imap.dooray.com:993` (SSL) | `smtp.dooray.com:465` (SSL) | **전용 암호 필수** | POP3 미지원 · `[메일 > 설정 > 외부 메일 연동]` 전용 암호 발급 *(공공기관용 제한)* |
| **카페24 웹메일 (Cafe24)** | **POP3 전용** | **POP3** | `webmail.{도메인}:110` (또는 `995` SSL) | `smtp.cafe24.com:587` (STARTTLS) | 일반 비밀번호 | IMAP 미지원 · 반드시 전체 이메일 주소로 로그인 |
| **한비로 그룹웨어** | IMAP · POP3 | IMAP | `imap.{도메인}:993` / `mail.{도메인}:993` | `smtp.{도메인}:465` | 일반 비밀번호 | 아이디/전체 이메일 자동 재시도 지원 |
| **카카오워크 메일** | IMAP · POP3 | IMAP | `imap.kakaowork.com:993` (SSL)<br>`pop3.kakaowork.com:995` (SSL) | `smtp.kakaowork.com:465` (SSL) | **앱 비밀번호 필수** | 카카오워크 웹메일 설정 > 외부 메일 연동 허용 |
| **후이즈 메일 (Whois)** | IMAP · POP3 | IMAP | `imap.whois.co.kr:993` / `pop.whois.co.kr:995` | `smtp.whois.co.kr:465` | 일반 비밀번호 | 환경설정 > POP3/IMAP 허용 |
| **Google Workspace / Gmail** | IMAP · POP3 | IMAP | `imap.gmail.com:993` (SSL)<br>`pop.gmail.com:995` (SSL) | `smtp.gmail.com:465` (SSL) | **16자리 앱 비밀번호 필수** | Google 계정 2단계 인증 ON → `[보안 > 앱 비밀번호]` 16자리 발급 |
| **Microsoft 365 / Outlook** | IMAP · POP3 | **OAuth** | `outlook.office365.com:993` / `995` | `smtp.office365.com:587` | **OAuth 권장** | 기업 계정은 Basic Auth(비밀번호) 차단됨 → 6장 Microsoft OAuth 사용 |
| **네이버 메일 (Naver)** | IMAP · POP3 | IMAP | `imap.naver.com:993` (SSL)<br>`pop.naver.com:995` (SSL) | `smtp.naver.com:465` (SSL) | **2단계 인증 시 앱비번** | 환경설정 > POP3/IMAP 사용함 + 해외 로그인 차단 확인 (아이디 자동 재시도) |
| **다음 / 카카오 메일 (Daum)** | IMAP · POP3 | IMAP | `imap.daum.net:993` (SSL)<br>`pop.daum.net:995` (SSL) | `smtp.daum.net:465` (SSL) | **2단계 인증 시 앱비번** | 환경설정 > IMAP/SMTP 사용함 + 카카오 계정 앱 비밀번호 발급 |
| **네이트 메일 (Nate)** | IMAP · POP3 | IMAP | `imap.mail.nate.com:993` (SSL)<br>`pop3.mail.nate.com:995` (SSL) | `smtp.mail.nate.com:465` (SSL) | 일반 비밀번호 | 발신 서버가 `smtp.mail.nate.com`임에 주의 (아이디 자동 재시도) |
| **Apple iCloud 메일** | **IMAP 전용** | **IMAP** | `imap.mail.me.com:993` (SSL) | `smtp.mail.me.com:587` (STARTTLS) | **앱 암호 필수** | POP3 미지원 · `appleid.apple.com`에서 앱 암호 생성 |
| **Zoho Mail** | IMAP · POP3 | IMAP | `imap.zoho.com:993` / `imappro.zoho.com:993` | `smtp.zoho.com:465` (SSL) | 2FA 시 앱 비밀번호 | Zoho Mail 설정 > Mail Accounts > IMAP Access 활성화 |
| **자사 메일 서버 (직접 입력)** | IMAP · POP3 | 자동 탐색 | `imap.{도메인}`, `pop3.{도메인}`, `mail.{도메인}` | `smtp.{도메인}:465` | 서버 정책 따름 | 고급 설정을 열어 호스트·포트·SSL 직접 지정 가능 |

---

## 3. 메일플러그(Mailplug) `#5.7.5` 완벽 해결 가이드 (필독)

메일플러그 연동 시 `-ERR 553 sorry, that password allowed relay (#5.7.5)` 또는 `A001 NO 553 ... (#5.7.5)`가 발생한다면 아래 **4가지를 순서대로 점검**하세요.

### 1단계: 해외 IP 차단 확인 (클라우드 미리보기 사용 시 가장 흔한 원인)
- 클라우드 프리뷰 환경은 해외 IP(예: 대만 `TW`)에서 메일플러그에 접속합니다.
- **조치 A (웹메일 허용)**:
  `메일플러그 웹메일 로그인` → `우측 상단 프로필/설정` → `[보안 설정]` → `[해외 로그인 차단 / 국가·IP 접근 설정]` → 앱에 표시된 **서버 접속 IP를 복사해 예외 IP에 등록**하거나 접속 국가 허용.
- **조치 B (권장 — 사내 로컬 실행)**:
  다운로드한 소스를 **한국 사내 PC에서 `npm run dev` 또는 Windows `.exe`**로 실행하면 한국 IP로 직접 연결되므로 해외 IP 차단이 발생하지 않습니다.

### 2단계: 웹메일 `POP3 사용함` 켜기 (30일 미사용 시 자동 OFF)
1. `메일플러그 웹메일` → 좌측 하단 **`[환경설정]`** → **`[POP3/IMAP 설정]`** 이동
2. **`POP3 사용함`** 체크 후 반드시 **`[저장]`** 클릭
3. **2025년 4월 1일 이후 개통된 그룹웨어/기업메일**은 메일플러그 공식 정책으로 **IMAP 이용이 제한**되므로 반드시 **POP3 (`pop3.mailplug.co.kr:995`, SSL)**를 선택하세요.

### 3단계: 일반 로그인 비밀번호가 아닌 `[앱 비밀번호]` 입력
1. `메일플러그 웹메일` → 우측 상단 **`프로필`** → **`[로그인 보안 설정]`** → **`[앱 비밀번호]`**
2. 새 앱 비밀번호를 생성해 Bmail 비밀번호 입력란에 붙여넣습니다. (하이픈·공백이 포함되어도 Bmail 자동 해결 엔진이 자동 보정합니다.)
3. 여전히 차단된다면 회사 메일플러그 **관리자 페이지**(`[관리자 > 메일 관리 > POP3/IMAP 사용 설정]` 및 `[보안 > 접근 IP 관리]`)에서 사원의 외부 메일 프로그램 연동 자체가 제한되어 있는지 확인하세요.

### 4단계: 사내 정책상 POP3/IMAP 차단 시 즉시 대안
- 기존 PC에 설치된 Outlook에서 **`.pst` / `.ost` / `.eml` 파일 불러오기** 기능을 사용하면 메일 서버 보안 정책과 무관하게 전체 메일·폴더·첨부파일·일정·프로젝트 분류를 100% 동일하게 분석할 수 있습니다.

---

## 4. 사내 전사 EMV 볼트 및 사원별 API 키 배포

Bmail은 **최고 관리자가 API 키를 한 번 등록하면 사내 10~20명+ 직원 모두가 별도 키 입력 없이 즉시 AI 기능을 사용**할 수 있도록 AES-256-GCM 암호화 볼트(`data/emv-vault.enc.json`)를 내장하고 있습니다.

### 4.1 키 해석 우선순위
사원이 로그인(또는 사원 이메일/ID 식별)되면 서버가 자동으로 아래 순서대로 키를 해석합니다.
1. **사원 개별 키 (`employee:<email>`)** — 특정 사원에게 별도 할당된 키가 있으면 최우선 적용
2. **전사 공통 키 (`global`)** — 최고 관리자가 등록한 회사 공용 키 적용
3. **서버 환경 변수 (`.env`)** — 위 두 항목이 없을 때 서버 `.env` 기본값 적용

### 4.2 최고 관리자 설정 방법
1. 앱 상단 우측 **방패 아이콘 (`EMV 및 서버 설정`)** 클릭
2. `.env`에 설정한 **`EMV_ADMIN_TOKEN`** 입력 후 관리자 잠금 해제
3. **`전사 공통 (모든 사원 자동 적용)`** 탭에서 `GEMINI_API_KEY`, `GROQ_API_KEY`, `ANTHROPIC_API_KEY`, `OPENAI_API_KEY` 등을 저장
4. 개별 과금/할당량이 필요한 사원이 있으면 **`사원 개별 설정`** 탭에서 사원 이메일(`hong@company.co.kr`)을 지정해 별도 키 등록
5. **보안 보장**: 저장된 API 키 원문은 브라우저·`config.json`·로그·소스 다운로드 파일 어디에도 노출되지 않으며, 화면에는 마스킹 값과 설정 여부(`true/false`)만 표시됩니다.

---

## 5. AI 엔진 · GCP 하이브리드 크레딧 설정

| 용도 | 권장 엔진 | 환경 변수 | 안전장치 |
|---|---|---|---|
| 실시간 메일 요약 · 3옵션 회신 · 일일 업무일지 | **Gemini API** (13.43만 원 크레딧) | `GEMINI_API_KEY`, `GEMINI_MODEL` | 서버 측 전용 호출, temperature/top_p 미전송 |
| 과거 누적 메일 대용량 RAG 검색 | **Vertex AI Agent Builder** (143만 원 크레딧) | `GCP_PROJECT_ID`, `AGENT_BUILDER_DATA_STORE_ID`, `AGENT_BUILDER_ACCESS_TOKEN` | **`2027-02-02` 이후 자동 과금 차단 하드스톱** |
| 무료 초고속 일지·요약 보조 | **Groq Cloud** | `GROQ_API_KEY` | `GROQ_CHUNK_CHAR_LIMIT = 6000` Map-Reduce 자동 청킹 (8K TPM 413 방지) |
| 한국어 정밀 보고서·블로그 | **Claude / OpenAI** | `ANTHROPIC_API_KEY`, `OPENAI_API_KEY` | EMV 볼트 또는 `.env` 선택 연동 |

---

## 6. Microsoft 365 OAuth 및 Outlook Add-in 설치

### 6.1 Microsoft Entra ID 앱 등록
1. [Microsoft Entra 관리 센터](https://entra.microsoft.com) → **앱 등록** → 새 등록
2. **리디렉션 URI (웹)**: `<APP_URL>/auth/microsoft/callback` (로컬: `http://localhost:3000/auth/microsoft/callback`)
3. **API 권한 추가 (Microsoft Graph 위임 권한)**: `User.Read`, `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`, `offline_access`
4. **인증서 및 비밀**에서 클라이언트 시크릿 발급 후 `.env` 또는 EMV 볼트에 등록:
   - `MICROSOFT_CLIENT_ID`
   - `MICROSOFT_CLIENT_SECRET`
   - `MICROSOFT_TENANT_ID` (단일 테넌트 ID 또는 `common`)

### 6.2 실제 Outlook에 Add-in 사이드로드
1. 앱 좌측 하단 **`manifest.xml 다운로드`** 또는 **`Outlook 플러그인 설치 센터`** 클릭
2. Outlook 웹/데스크톱 → **[추가 기능 가져오기] > [내 추가 기능] > [사용자 지정 추가 기능 추가] > [파일에서 추가]** → `manifest.xml` 업로드
3. `SupportsPinning`이 포함되어 있어 메일을 전환해도 우측 Bmail AI 작업창을 고정 유지할 수 있습니다.

---

## 7. 로컬 PC Outlook PST · 주소록 자동 인수 및 Windows `.exe` 설치형 빌드

### 7.1 기존 PC Outlook 설정·주소록·PST 1클릭 자동 파악
1. 앱 상단 **`EMV 및 서버 설정`** 또는 **`원본 PST · EML 불러오기`** 창에서 **로컬 PC Outlook 설정 자동 탐색 스크립트**(`.ps1`)를 다운로드/실행합니다.
2. 생성된 프로파일 JSON을 업로드하면 현재 PC에 설치된 Outlook의 **계정 목록, PST 파일 경로, 주소록, 서명/폴더 구조**가 자동 반영됩니다.

### 7.2 Windows 설치형 인스톨러(`.exe`) 빌드 절차 (Windows PC에서 실행)

`desktop/` 폴더에 `main.cjs`, `preload.cjs`, `electron-builder.json`이 이미 준비되어 있습니다.

```powershell
# 1) 프론트엔드 빌드 확인
npm install
npm run build

# 2) Electron 패키징 도구 설치 (Windows 빌드 머신 전용)
npm install -D electron electron-builder

# 3) Windows NSIS 설치형 .exe 생성 (release/ 폴더에 출력)
npx electron-builder --config desktop/electron-builder.json --win nsis
```

> **사내 10~20명 배포 팁**: 빌드 시 중앙 서버 주소(`VITE_API_BASE_URL=https://<사내-또는-Cloudflare-도메인>`)를 연결하면, 사원들은 설치형 `.exe`만 실행해 본인 이메일/PST로 로그인하고 AI API 키는 최고 관리자의 EMV 볼트를 자동 공유합니다.

---

## 8. Cloudflare Workers + Static Assets 배포 절차

상세 아키텍처와 후속 확장 작업(GCP 서비스 계정 JWT 자동 토큰 갱신, EMV 볼트 KV/D1 이관, CORS 화이트리스트)은 **`HANDOFF.md`** 에 별도 정리되어 있습니다.

```bash
# 1) 프론트엔드 dist/ 빌드
npm ci
npm run lint
npm run build

# 2) Cloudflare Secrets 1회 등록 (wrangler.toml에 비밀 키 저장 금지)
npx wrangler secret put GEMINI_API_KEY
npx wrangler secret put GCP_PROJECT_ID
npx wrangler secret put AGENT_BUILDER_DATA_STORE_ID
npx wrangler secret put AGENT_BUILDER_ACCESS_TOKEN
npx wrangler secret put BMAIL_ORG_ACCESS_TOKEN

# 3) 배포 (dist/ 정적 SPA + cloudflare/worker.ts 동시 배포)
npx wrangler deploy

# 4) 헬스체크 확인
curl -s https://<YOUR-SUBDOMAIN>.workers.dev/api/worker/health
```

> **주의**: Cloudflare Workers는 일반 TCP/TLS 소켓 제약이 있으므로 사내 메일 서버 POP3/IMAP 직접 접속은 **Windows `.exe` 로컬 백엔드** 또는 **한국 리전 Node.js 서버(`server.ts`)**에서 실행하는 하이브리드 구조를 권장합니다.

---

## 9. 검증 체크리스트(DoD) 및 장애 대응 FAQ

### 9.1 빌드 전 필수 검증 명령

```bash
# 1. 전체 단위 테스트
npx --yes tsx --test \
  src/utils/mailFilter.test.ts \
  src/utils/appConfig.test.ts \
  src/utils/deterministicAnalyzer.test.ts \
  src/utils/outlookWorkspaceModel.test.ts \
  src/utils/dashboardAnalytics.test.ts \
  src/utils/mailPriority.test.ts \
  src/utils/headerSyncState.test.ts \
  src/services/emvVaultService.test.ts \
  src/services/sourceArchiveBuilder.test.ts \
  src/services/pop3AutoConnect.test.ts

# 2. TypeScript 정적 검증 (오류 0건)
npm run lint

# 3. 프로덕션 빌드 검증
npm run build
```

### 9.2 자주 묻는 문제 해결 (FAQ)

- **Q1. `EADDRINUSE: address already in use :::3000` 오류가 납니다.**
  - 다른 프로그램이 3000 포트를 사용 중입니다. `.env`에서 `PORT=3001`로 변경하거나 기존 프로세스를 종료하세요.
- **Q2. 메일플러그 앱 비밀번호를 정확히 넣었는데도 `#5.7.5`가 뜹니다.**
  - 90% 이상이 **① 해외 IP 차단**(클라우드 서버 접속 시) 또는 **② 웹메일 `[POP3/IMAP 설정]` OFF(30일 미사용 자동 해제)** 또는 **③ 회사 관리자 외부 POP3 차단**입니다. 한국 PC에서 로컬 실행(`npm run dev`) 후 테스트하면 해외 IP 차단을 즉시 배제할 수 있습니다.
- **Q3. 대용량 `.pst` 파일을 올리면 시간이 오래 걸립니다.**
  - 파일 크기가 300MB 이상인 경우 Outlook에서 최근 연도별/폴더별로 `.pst`를 분할 내보내기하거나 `.eml`/`.mbox` 단위로 나누어 불러오면 훨씬 빠르게 인덱싱됩니다.
- **Q4. 다운로드한 소스 압축 파일에 내 비밀번호나 `.env`가 들어있나요?**
  - 아니요. `sourceArchiveBuilder`가 `.env`, 암호화 볼트(`data/`), 인증서(`.pem/.key/.pfx`), 메일 원본(`.pst/.ost`), 로그(`.log`), `node_modules/`, `dist/`를 원천 제외합니다.
