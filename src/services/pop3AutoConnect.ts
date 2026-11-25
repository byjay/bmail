import crypto from 'node:crypto';
import dns from 'node:dns/promises';
import fs from 'node:fs';
import net from 'node:net';
import os from 'node:os';
import path from 'node:path';
import tls from 'node:tls';
import type { MailItem } from '../types/mail';
import { parseEmlMessage } from '../utils/mailFileParsers';

export type MailReceiveProtocol = 'pop3' | 'imap';

export type MailServicePresetId =
  | 'auto'
  | 'mailplug'
  | 'hiworks'
  | 'gabia'
  | 'naverworks'
  | 'daouoffice'
  | 'dooray'
  | 'cafe24'
  | 'hanbiro'
  | 'kakaowork'
  | 'whois'
  | 'google'
  | 'microsoft'
  | 'naver'
  | 'daum'
  | 'nate'
  | 'icloud'
  | 'zoho'
  | 'custom';

export type UsernameFormatPolicy = 'full-email' | 'local-or-full';
export type AppPasswordPolicy = 'required' | 'recommended' | 'optional';

export interface MailServicePresetInfo {
  id: MailServicePresetId;
  label: string;
  vendorHint: string;
  supportedProtocols: MailReceiveProtocol[];
  defaultProtocol: MailReceiveProtocol;
  pop3HostTemplate: string;
  pop3Port: number;
  pop3Tls: boolean;
  pop3AltTemplates?: Array<{ hostTemplate: string; port: number; tls: boolean }>;
  imapHostTemplate: string;
  imapPort: number;
  imapTls: boolean;
  imapAltTemplates?: Array<{ hostTemplate: string; port: number; tls: boolean }>;
  smtpHostTemplate: string;
  smtpPort: number;
  smtpTls: boolean;
  mxPatterns: string[];
  usernameFormat: UsernameFormatPolicy;
  appPasswordPolicy: AppPasswordPolicy;
  overseasIpWarning: boolean;
  appPasswordGuide: string;
  protocolSetupGuide: string;
  note: string;
}

export const MAIL_SERVICE_PRESETS: MailServicePresetInfo[] = [
  {
    id: 'auto',
    label: '자동 감지 (회사 도메인 MX · POP3/IMAP 자동 조회)',
    vendorHint: 'Auto-Discovery',
    supportedProtocols: ['imap', 'pop3'],
    defaultProtocol: 'imap',
    pop3HostTemplate: 'pop3.{domain}',
    pop3Port: 995,
    pop3Tls: true,
    imapHostTemplate: 'imap.{domain}',
    imapPort: 993,
    imapTls: true,
    smtpHostTemplate: 'smtp.{domain}',
    smtpPort: 465,
    smtpTls: true,
    mxPatterns: [],
    usernameFormat: 'local-or-full',
    appPasswordPolicy: 'optional',
    overseasIpWarning: false,
    appPasswordGuide: '2단계 인증이나 앱 비밀번호 정책이 있는 서비스는 웹메일 보안 설정에서 앱 비밀번호를 발급해 입력하세요.',
    protocolSetupGuide: '도메인의 MX 레코드와 POP3(995/110)·IMAP(993) 포트를 자동 탐색해 최적 프로토콜을 선택합니다.',
    note: '이메일 도메인의 MX 레코드와 POP3/IMAP 포트를 자동 조회해 회사 메일 서버 설정을 완성합니다.',
  },
  {
    id: 'mailplug',
    label: '메일플러그 (Mailplug)',
    vendorHint: 'mailplug.co.kr',
    supportedProtocols: ['pop3', 'imap'],
    defaultProtocol: 'pop3',
    pop3HostTemplate: 'pop3.mailplug.co.kr',
    pop3Port: 995,
    pop3Tls: true,
    pop3AltTemplates: [{ hostTemplate: 'pop3.{domain}', port: 995, tls: true }],
    imapHostTemplate: 'imap.mailplug.co.kr',
    imapPort: 993,
    imapTls: true,
    imapAltTemplates: [{ hostTemplate: 'imap.{domain}', port: 993, tls: true }],
    smtpHostTemplate: 'smtp.mailplug.co.kr',
    smtpPort: 465,
    smtpTls: true,
    mxPatterns: ['mailplug.co.kr', 'mailplug.com', 'mfilter.mailplug'],
    usernameFormat: 'full-email',
    appPasswordPolicy: 'required',
    overseasIpWarning: true,
    appPasswordGuide:
      '웹메일/그룹웨어 로그인 비밀번호는 외부 연동 시 #5.7.5 오류로 차단됩니다. [프로필 > 로그인 보안 설정 > 앱 비밀번호]에서 전용 비밀번호를 생성해 입력하세요.',
    protocolSetupGuide:
      '[웹메일 환경설정 > POP3/IMAP 설정]에서 사용함을 켜세요(30일 미사용 시 자동 해제). 2025년 4월 1일 이후 개통된 그룹웨어/기업메일은 IMAP 이용이 제한되므로 POP3(995 SSL) 사용을 권장하며, 클라우드 접속 시 [해외 IP 차단] 예외 또는 로컬 .exe 실행이 필요할 수 있습니다.',
    note: '메일플러그는 POP3(995 SSL) 권장(2025년 4월 1일 이후 개통 서비스는 IMAP 제한), [앱 비밀번호] 발급, 30일 미사용 해제 확인, 해외 IP 차단 해제(클라우드 접속 시)가 필수입니다.',
  },
  {
    id: 'hiworks',
    label: '하이웍스 (Hiworks)',
    vendorHint: 'hiworks.com',
    supportedProtocols: ['pop3'],
    defaultProtocol: 'pop3',
    pop3HostTemplate: 'pop3s.hiworks.com',
    pop3Port: 995,
    pop3Tls: true,
    pop3AltTemplates: [{ hostTemplate: 'mail.{domain}', port: 995, tls: true }],
    imapHostTemplate: '',
    imapPort: 993,
    imapTls: true,
    smtpHostTemplate: 'smtps.hiworks.com',
    smtpPort: 465,
    smtpTls: true,
    mxPatterns: ['hiworks.co.kr', 'hiworks.com'],
    usernameFormat: 'full-email',
    appPasswordPolicy: 'recommended',
    overseasIpWarning: true,
    appPasswordGuide:
      '하이웍스 보안 정책에 따라 오피스 로그인 비밀번호가 아닌 [메일 전용 비밀번호]를 생성해 입력해야 할 수 있습니다.',
    protocolSetupGuide:
      '하이웍스는 IMAP을 미지원하며 POP3(995 SSL)만 지원합니다. [메일 > 환경설정 > POP3/SMTP 설정]에서 사용함을 켜세요.',
    note: '하이웍스는 POP3 전용(IMAP 미지원)이며, 외부 메일 허용 및 메일 전용 비밀번호 확인이 필요합니다.',
  },
  {
    id: 'gabia',
    label: '가비아 기업메일 / 하이메일 (Gabia)',
    vendorHint: 'gabia.com',
    supportedProtocols: ['pop3'],
    defaultProtocol: 'pop3',
    pop3HostTemplate: 'pop3s.hiworks.com',
    pop3Port: 995,
    pop3Tls: true,
    pop3AltTemplates: [
      { hostTemplate: 'mail.{domain}', port: 995, tls: true },
      { hostTemplate: 'mail.{domain}', port: 110, tls: false },
    ],
    imapHostTemplate: '',
    imapPort: 993,
    imapTls: true,
    smtpHostTemplate: 'smtps.hiworks.com',
    smtpPort: 465,
    smtpTls: true,
    mxPatterns: ['gabia.com', 'gabia.co.kr'],
    usernameFormat: 'full-email',
    appPasswordPolicy: 'recommended',
    overseasIpWarning: true,
    appPasswordGuide: '가비아/하이웍스 메일 전용 비밀번호 또는 2단계 앱 비밀번호를 확인하세요.',
    protocolSetupGuide: '하이웍스 연동형(pop3s.hiworks.com:995) 또는 도메인형(mail.회사도메인:995/110)을 자동 탐지합니다.',
    note: '가비아 기업메일은 POP3S(995) 게이트웨이 및 mail.도메인 주소를 자동 점검합니다.',
  },
  {
    id: 'naverworks',
    label: '네이버 웍스 (NAVER WORKS)',
    vendorHint: 'worksmobile.com',
    supportedProtocols: ['imap', 'pop3'],
    defaultProtocol: 'imap',
    pop3HostTemplate: 'pop.worksmobile.com',
    pop3Port: 995,
    pop3Tls: true,
    imapHostTemplate: 'imap.worksmobile.com',
    imapPort: 993,
    imapTls: true,
    smtpHostTemplate: 'smtp.worksmobile.com',
    smtpPort: 465,
    smtpTls: true,
    mxPatterns: ['worksmobile.com', 'navercloud'],
    usernameFormat: 'full-email',
    appPasswordPolicy: 'required',
    overseasIpWarning: true,
    appPasswordGuide: '네이버 웍스 [내 정보 > 보안 > 외부 앱 비밀번호]에서 전용 비밀번호를 발급해 입력하세요.',
    protocolSetupGuide: '관리자(Admin) [보안 > 서비스 권한]과 개인 메일 환경설정에서 IMAP/POP3 사용이 허용되어야 합니다.',
    note: '네이버 웍스는 IMAP(993)·POP3(995) 모두 지원하며 외부 앱 비밀번호 발급이 필요합니다.',
  },
  {
    id: 'daouoffice',
    label: '다우오피스 (DaouOffice)',
    vendorHint: 'daouoffice.com',
    supportedProtocols: ['imap', 'pop3'],
    defaultProtocol: 'imap',
    pop3HostTemplate: 'pop3s.daouoffice.com',
    pop3Port: 995,
    pop3Tls: true,
    pop3AltTemplates: [
      { hostTemplate: 'mail.{domain}', port: 995, tls: true },
      { hostTemplate: 'pop3.daouoffice.com', port: 995, tls: true },
    ],
    imapHostTemplate: 'imap.daouoffice.com',
    imapPort: 993,
    imapTls: true,
    imapAltTemplates: [{ hostTemplate: 'mail.{domain}', port: 993, tls: true }],
    smtpHostTemplate: 'outbound.daouoffice.com',
    smtpPort: 465,
    smtpTls: true,
    mxPatterns: ['daouoffice', 'terracetech', 'daou.co.kr'],
    usernameFormat: 'full-email',
    appPasswordPolicy: 'optional',
    overseasIpWarning: false,
    appPasswordGuide: 'OTP/2단계 인증 사용 계정은 다우오피스 보안 설정의 앱 비밀번호를 입력하세요.',
    protocolSetupGuide: '클라우드형(pop3s.daouoffice.com / outbound.daouoffice.com)과 사내 구축형(mail.도메인)을 자동 탐색합니다.',
    note: '다우오피스는 IMAP(993)·POP3(995) 및 outbound.daouoffice.com(465 SSL) 발송을 지원합니다.',
  },
  {
    id: 'dooray',
    label: '두레이 (NHN Dooray)',
    vendorHint: 'dooray.com',
    supportedProtocols: ['imap'],
    defaultProtocol: 'imap',
    pop3HostTemplate: '',
    pop3Port: 995,
    pop3Tls: true,
    imapHostTemplate: 'imap.dooray.com',
    imapPort: 993,
    imapTls: true,
    imapAltTemplates: [{ hostTemplate: 'imap.{domain}', port: 993, tls: true }],
    smtpHostTemplate: 'smtp.dooray.com',
    smtpPort: 465,
    smtpTls: true,
    mxPatterns: ['dooray.com', 'toast.com'],
    usernameFormat: 'full-email',
    appPasswordPolicy: 'required',
    overseasIpWarning: false,
    appPasswordGuide: '두레이 [메일 > 설정 > 외부 메일 연동(IMAP/SMTP)]에서 발급한 전용 암호를 입력해야 합니다.',
    protocolSetupGuide: '두레이는 POP3를 미지원하며 IMAP(993 SSL) 전용입니다. (공공기관용 두레이는 외부 IMAP 연동 제한)',
    note: '두레이는 IMAP(993 SSL) 전용이며 웹메일에서 발급한 IMAP/SMTP 전용 암호가 필수입니다.',
  },
  {
    id: 'cafe24',
    label: '카페24 웹메일 (Cafe24)',
    vendorHint: 'cafe24.com',
    supportedProtocols: ['pop3'],
    defaultProtocol: 'pop3',
    pop3HostTemplate: 'webmail.{domain}',
    pop3Port: 110,
    pop3Tls: false,
    pop3AltTemplates: [
      { hostTemplate: 'webmail.{domain}', port: 995, tls: true },
      { hostTemplate: 'pop3.{domain}', port: 995, tls: true },
      { hostTemplate: 'pop3.{domain}', port: 110, tls: false },
      { hostTemplate: 'mail.{domain}', port: 110, tls: false },
    ],
    imapHostTemplate: '',
    imapPort: 993,
    imapTls: true,
    smtpHostTemplate: 'smtp.cafe24.com',
    smtpPort: 587,
    smtpTls: false,
    mxPatterns: ['cafe24.com', 'cafe24corp', 'mw-0'],
    usernameFormat: 'full-email',
    appPasswordPolicy: 'optional',
    overseasIpWarning: false,
    appPasswordGuide: '카페24 웹메일 계정 비밀번호와 전체 이메일 주소(id@도메인)를 사용합니다.',
    protocolSetupGuide:
      '카페24는 IMAP을 미지원하며 POP3(webmail.회사도메인:110 또는 995) + SMTP(smtp.cafe24.com:587)를 사용합니다. 웹메일 환경설정에서 POP3/SMTP 사용함을 켜세요.',
    note: '카페24는 POP3 전용(webmail.도메인:110 / smtp.cafe24.com:587)으로 자동 탐색합니다.',
  },
  {
    id: 'hanbiro',
    label: '한비로 그룹웨어 (Hanbiro)',
    vendorHint: 'hanbiro.com',
    supportedProtocols: ['imap', 'pop3'],
    defaultProtocol: 'imap',
    pop3HostTemplate: 'mail.{domain}',
    pop3Port: 995,
    pop3Tls: true,
    pop3AltTemplates: [{ hostTemplate: 'pop3.{domain}', port: 995, tls: true }],
    imapHostTemplate: 'mail.{domain}',
    imapPort: 993,
    imapTls: true,
    imapAltTemplates: [{ hostTemplate: 'imap.{domain}', port: 993, tls: true }],
    smtpHostTemplate: 'mail.{domain}',
    smtpPort: 587,
    smtpTls: false,
    mxPatterns: ['hanbiro.com', 'hanbiro.net'],
    usernameFormat: 'local-or-full',
    appPasswordPolicy: 'optional',
    overseasIpWarning: false,
    appPasswordGuide: '한비로 그룹웨어 계정 비밀번호를 사용합니다. 전체 이메일 실패 시 아이디만 자동 재시도합니다.',
    protocolSetupGuide: 'mail.회사도메인 기준 IMAP(993 SSL), POP3(995 SSL), SMTP(587 STARTTLS)를 사용합니다.',
    note: '한비로 그룹웨어는 mail.도메인 IMAP(993)·POP3(995) 및 전체 이메일/아이디 자동 재시도를 지원합니다.',
  },
  {
    id: 'kakaowork',
    label: '카카오워크 메일 (Kakao Work)',
    vendorHint: 'kakao.com',
    supportedProtocols: ['imap', 'pop3'],
    defaultProtocol: 'imap',
    pop3HostTemplate: 'pop.kakao.com',
    pop3Port: 995,
    pop3Tls: true,
    imapHostTemplate: 'imap.kakao.com',
    imapPort: 993,
    imapTls: true,
    smtpHostTemplate: 'smtp.kakao.com',
    smtpPort: 465,
    smtpTls: true,
    mxPatterns: ['kakaowork.com', 'kakaoenterprise'],
    usernameFormat: 'full-email',
    appPasswordPolicy: 'recommended',
    overseasIpWarning: false,
    appPasswordGuide: '2단계 인증 적용 계정은 카카오워크 메일 환경설정에서 앱 비밀번호를 발급해 입력하세요.',
    protocolSetupGuide: '카카오워크 메일 환경설정에서 IMAP/POP3 외부 연동을 활성화하세요.',
    note: '카카오워크 메일은 imap.kakao.com(993) / pop.kakao.com(995) SSL 연결을 지원합니다.',
  },
  {
    id: 'whois',
    label: '후이즈 메일 (Whois)',
    vendorHint: 'whoisworks.com',
    supportedProtocols: ['imap', 'pop3'],
    defaultProtocol: 'pop3',
    pop3HostTemplate: 'pop3.whoisworks.com',
    pop3Port: 995,
    pop3Tls: true,
    pop3AltTemplates: [{ hostTemplate: 'mail.{domain}', port: 995, tls: true }],
    imapHostTemplate: 'imap.whoisworks.com',
    imapPort: 993,
    imapTls: true,
    smtpHostTemplate: 'smtp.whoisworks.com',
    smtpPort: 465,
    smtpTls: true,
    mxPatterns: ['whoisworks.com', 'whois.co.kr'],
    usernameFormat: 'full-email',
    appPasswordPolicy: 'optional',
    overseasIpWarning: false,
    appPasswordGuide: '후이즈 메일 로그인 비밀번호를 사용합니다.',
    protocolSetupGuide: '후이즈 메일 환경설정에서 POP3/IMAP/SMTP 사용함을 활성화하세요.',
    note: '후이즈 메일은 pop3.whoisworks.com(995) 및 imap.whoisworks.com(993)을 지원합니다.',
  },
  {
    id: 'google',
    label: 'Google Workspace / Gmail',
    vendorHint: 'gmail.com',
    supportedProtocols: ['imap', 'pop3'],
    defaultProtocol: 'imap',
    pop3HostTemplate: 'pop.gmail.com',
    pop3Port: 995,
    pop3Tls: true,
    imapHostTemplate: 'imap.gmail.com',
    imapPort: 993,
    imapTls: true,
    smtpHostTemplate: 'smtp.gmail.com',
    smtpPort: 465,
    smtpTls: true,
    mxPatterns: ['google.com', 'googlemail.com', 'aspmx.l.google.com', 'smtp.google.com'],
    usernameFormat: 'full-email',
    appPasswordPolicy: 'required',
    overseasIpWarning: false,
    appPasswordGuide: 'Google 계정 [보안 > 2단계 인증 > 앱 비밀번호]에서 16자리 앱 비밀번호를 생성해 공백 없이 입력하세요.',
    protocolSetupGuide: 'Gmail은 POP3가 점진적으로 축소되므로 IMAP(imap.gmail.com:993 SSL) 연결을 권장합니다.',
    note: 'Google Workspace/Gmail은 16자리 앱 비밀번호가 필수이며 IMAP(993 SSL) 연결을 기본 권장합니다.',
  },
  {
    id: 'microsoft',
    label: 'Microsoft 365 / Outlook',
    vendorHint: 'office365.com',
    supportedProtocols: ['imap', 'pop3'],
    defaultProtocol: 'imap',
    pop3HostTemplate: 'outlook.office365.com',
    pop3Port: 995,
    pop3Tls: true,
    imapHostTemplate: 'outlook.office365.com',
    imapPort: 993,
    imapTls: true,
    smtpHostTemplate: 'smtp.office365.com',
    smtpPort: 587,
    smtpTls: false,
    mxPatterns: ['protection.outlook.com', 'outlook.com', 'hotmail.com'],
    usernameFormat: 'full-email',
    appPasswordPolicy: 'required',
    overseasIpWarning: false,
    appPasswordGuide:
      'Microsoft 365는 기본 인증(Basic Auth)이 대부분 차단되어 있으므로 상단 Microsoft 365 OAuth 로그인 또는 앱 비밀번호 사용을 권장합니다.',
    protocolSetupGuide: 'Exchange Online 기본 인증이 비활성화된 사내 계정은 Graph OAuth2 연동이 필요합니다.',
    note: 'Microsoft 365는 기본 인증 제한 정책이 있어 OAuth 연동 또는 앱 비밀번호 허용 테넌트에서만 IMAP/POP3 접속이 가능합니다.',
  },
  {
    id: 'naver',
    label: '네이버 메일 (Naver)',
    vendorHint: 'naver.com',
    supportedProtocols: ['imap', 'pop3'],
    defaultProtocol: 'imap',
    pop3HostTemplate: 'pop.naver.com',
    pop3Port: 995,
    pop3Tls: true,
    imapHostTemplate: 'imap.naver.com',
    imapPort: 993,
    imapTls: true,
    smtpHostTemplate: 'smtp.naver.com',
    smtpPort: 587,
    smtpTls: false,
    mxPatterns: ['naver.com'],
    usernameFormat: 'local-or-full',
    appPasswordPolicy: 'required',
    overseasIpWarning: true,
    appPasswordGuide:
      '네이버 2단계 인증 사용 시 [네이버ID > 보안설정 > 애플리케이션 비밀번호]를 발급해 입력하세요. 아이디 형식(id / id@naver.com)은 자동 재시도합니다.',
    protocolSetupGuide:
      '[네이버 메일 > 환경설정 > POP3/IMAP 설정]에서 사용함을 켜고, 클라우드 접속 시 [해외 로그인 차단] 여부를 확인하세요.',
    note: '네이버 메일은 2FA 앱 비밀번호, POP3/IMAP 허용, 해외 로그인 차단 해제(클라우드 접속 시)가 필요합니다.',
  },
  {
    id: 'daum',
    label: '다음 / 카카오 메일 (Daum)',
    vendorHint: 'daum.net',
    supportedProtocols: ['imap', 'pop3'],
    defaultProtocol: 'imap',
    pop3HostTemplate: 'pop.daum.net',
    pop3Port: 995,
    pop3Tls: true,
    imapHostTemplate: 'imap.daum.net',
    imapPort: 993,
    imapTls: true,
    smtpHostTemplate: 'smtp.daum.net',
    smtpPort: 465,
    smtpTls: true,
    mxPatterns: ['daum.net', 'hanmail.net', 'kakao.com'],
    usernameFormat: 'local-or-full',
    appPasswordPolicy: 'required',
    overseasIpWarning: true,
    appPasswordGuide:
      '다음/카카오 계정은 [계정 보안 > 앱 비밀번호] 발급이 필수입니다. Daum 아이디(도메인 제외) 자동 재시도를 지원합니다.',
    protocolSetupGuide: '[다음 메일 > 환경설정 > IMAP/POP3]에서 사용함을 활성화하세요.',
    note: '다음/카카오 메일은 앱 비밀번호 발급과 IMAP/POP3 사용 설정이 필수입니다.',
  },
  {
    id: 'nate',
    label: '네이트 메일 (Nate)',
    vendorHint: 'nate.com',
    supportedProtocols: ['imap', 'pop3'],
    defaultProtocol: 'imap',
    pop3HostTemplate: 'pop3.nate.com',
    pop3Port: 995,
    pop3Tls: true,
    imapHostTemplate: 'imap.nate.com',
    imapPort: 993,
    imapTls: true,
    smtpHostTemplate: 'smtp.mail.nate.com',
    smtpPort: 465,
    smtpTls: true,
    mxPatterns: ['nate.com', 'empas.com'],
    usernameFormat: 'local-or-full',
    appPasswordPolicy: 'optional',
    overseasIpWarning: false,
    appPasswordGuide: '네이트 로그인 비밀번호를 사용하며, 아이디(도메인 제외)와 전체 이메일을 자동 시도합니다.',
    protocolSetupGuide: '[네이트 메일 > 환경설정 > POP3/IMAP 설정]에서 사용함을 켜세요. SMTP는 smtp.mail.nate.com:465(SSL)를 사용합니다.',
    note: '네이트 메일은 imap.nate.com(993), pop3.nate.com(995), smtp.mail.nate.com(465)을 사용합니다.',
  },
  {
    id: 'icloud',
    label: 'Apple iCloud 메일',
    vendorHint: 'icloud.com',
    supportedProtocols: ['imap'],
    defaultProtocol: 'imap',
    pop3HostTemplate: '',
    pop3Port: 995,
    pop3Tls: true,
    imapHostTemplate: 'imap.mail.me.com',
    imapPort: 993,
    imapTls: true,
    smtpHostTemplate: 'smtp.mail.me.com',
    smtpPort: 587,
    smtpTls: false,
    mxPatterns: ['icloud.com', 'me.com', 'mac.com'],
    usernameFormat: 'full-email',
    appPasswordPolicy: 'required',
    overseasIpWarning: false,
    appPasswordGuide: 'Apple 계정 페이지(appleid.apple.com) [로그인 및 보안 > 앱 전용 암호]에서 발급한 비밀번호를 입력하세요.',
    protocolSetupGuide: 'iCloud 메일은 POP3를 지원하지 않으며 IMAP(imap.mail.me.com:993 SSL) 전용입니다.',
    note: 'iCloud 메일은 IMAP(993 SSL) 전용이며 Apple 앱 전용 암호가 필수입니다.',
  },
  {
    id: 'zoho',
    label: 'Zoho Mail',
    vendorHint: 'zoho.com',
    supportedProtocols: ['imap', 'pop3'],
    defaultProtocol: 'imap',
    pop3HostTemplate: 'pop.zoho.com',
    pop3Port: 995,
    pop3Tls: true,
    pop3AltTemplates: [{ hostTemplate: 'poppro.zoho.com', port: 995, tls: true }],
    imapHostTemplate: 'imap.zoho.com',
    imapPort: 993,
    imapTls: true,
    imapAltTemplates: [{ hostTemplate: 'imappro.zoho.com', port: 993, tls: true }],
    smtpHostTemplate: 'smtp.zoho.com',
    smtpPort: 465,
    smtpTls: true,
    mxPatterns: ['zoho.com', 'zohomail'],
    usernameFormat: 'full-email',
    appPasswordPolicy: 'required',
    overseasIpWarning: false,
    appPasswordGuide: 'Zoho Accounts [Security > App Passwords]에서 전용 비밀번호를 생성하세요.',
    protocolSetupGuide: 'Zoho Mail 설정에서 IMAP/POP 액세스를 활성화하세요. 조직 도메인은 imappro.zoho.com을 자동 탐색합니다.',
    note: 'Zoho Mail은 IMAP(993)·POP3(995) 및 앱 비밀번호 인증을 지원합니다.',
  },
  {
    id: 'custom',
    label: '자사 메일 서버 (직접 / 자동 탐색)',
    vendorHint: 'Custom Domain',
    supportedProtocols: ['imap', 'pop3'],
    defaultProtocol: 'imap',
    pop3HostTemplate: 'pop3.{domain}',
    pop3Port: 995,
    pop3Tls: true,
    imapHostTemplate: 'imap.{domain}',
    imapPort: 993,
    imapTls: true,
    smtpHostTemplate: 'smtp.{domain}',
    smtpPort: 465,
    smtpTls: true,
    mxPatterns: [],
    usernameFormat: 'local-or-full',
    appPasswordPolicy: 'optional',
    overseasIpWarning: false,
    appPasswordGuide: '사내 메일 서버 정책에 따라 계정 비밀번호 또는 앱 비밀번호를 입력하세요.',
    protocolSetupGuide: 'imap.<도메인>, pop3.<도메인>, mail.<도메인>, webmail.<도메인> 포트를 자동 점검합니다.',
    note: '자사 호스팅 도메인의 IMAP(993)·POP3(995/110) 서버를 자동 점검합니다.',
  },
];

export interface DiscoveredPop3Config {
  email: string;
  domain: string;
  detectedProviderId: MailServicePresetId;
  detectedProviderName: string;
  confidence: 'high' | 'medium' | 'fallback';
  detectionMethod: 'preset' | 'mx-record' | 'domain-known' | 'port-probe' | 'heuristic';
  mxRecords: string[];
  supportedProtocols: MailReceiveProtocol[];
  recommendedProtocol: MailReceiveProtocol;
  pop3Host: string;
  pop3Port: number;
  pop3Tls: boolean;
  pop3Reachable: boolean;
  pop3Banner?: string;
  imapHost: string;
  imapPort: number;
  imapTls: boolean;
  imapReachable: boolean;
  imapBanner?: string;
  smtpHost: string;
  smtpPort: number;
  smtpTls: boolean;
  reachable: boolean;
  usernameFormat: UsernameFormatPolicy;
  appPasswordPolicy: AppPasswordPolicy;
  overseasIpWarning: boolean;
  appPasswordGuide: string;
  protocolSetupGuide: string;
  guidanceNote: string;
}

export interface MailConnectionDiagnosis {
  code:
    | 'app-password-required'
    | 'overseas-ip-or-policy'
    | 'protocol-disabled'
    | 'oauth-required'
    | 'auth-failed'
    | 'connection-failed';
  headline: string;
  checklist: string[];
  suggestProtocolSwitch?: MailReceiveProtocol;
  rawServerError: string;
  serverEgress?: ServerEgressInfo;
  autoHealAttempts?: AutoHealAttempt[];
  resolutionSteps?: ResolutionGuideStep[];
}

export interface ServerEgressInfo {
  ipv4: string | null;
  countryCode: string | null;
  countryName: string | null;
  city: string | null;
  isKorea: boolean | null;
}

export interface AutoHealAttempt {
  step: number;
  protocol: MailReceiveProtocol;
  host: string;
  port: number;
  strategyLabel: string;
  outcome: 'success' | 'failed';
  serverCode?: string;
}

export interface ResolutionGuideStep {
  id: string;
  priority: 'critical' | 'recommended' | 'alternative';
  title: string;
  menuPath: string;
  detail: string;
  copyValue?: string;
  copyLabel?: string;
  helpUrl?: string;
  helpLabel?: string;
}

export class MailAuthError extends Error {
  readonly diagnosis: MailConnectionDiagnosis;

  constructor(message: string, diagnosis: MailConnectionDiagnosis) {
    super(message);
    this.name = 'MailAuthError';
    this.diagnosis = diagnosis;
  }
}

export interface Pop3ConnectionParams {
  email: string;
  password: string;
  pop3Host: string;
  pop3Port: number;
  pop3Tls: boolean;
  providerId?: MailServicePresetId;
  usernameFormat?: UsernameFormatPolicy;
  maxMessages?: number;
  timeoutMs?: number;
  /** Test-only flag for loopback mock POP3 servers; NEVER set true from HTTP handlers */
  allowLoopbackForTest?: boolean;
}

export interface ImapConnectionParams {
  email: string;
  password: string;
  imapHost: string;
  imapPort: number;
  imapTls: boolean;
  providerId?: MailServicePresetId;
  usernameFormat?: UsernameFormatPolicy;
  maxMessages?: number;
  timeoutMs?: number;
  allowLoopbackForTest?: boolean;
}

export interface Pop3FetchResult {
  protocolUsed: MailReceiveProtocol;
  loginUsernameUsed: string;
  totalMailboxCount: number;
  totalOctets: number;
  fetchedCount: number;
  emails: MailItem[];
  banner: string;
}

export interface SmtpSendParams {
  email: string;
  password: string;
  smtpHost: string;
  smtpPort: number;
  smtpTls: boolean;
  to: string[];
  cc?: string[];
  subject: string;
  body: string;
  htmlBody?: string;
  timeoutMs?: number;
  allowLoopbackForTest?: boolean;
}

const ALLOWED_POP3_PORTS = new Set([995, 110]);
const ALLOWED_IMAP_PORTS = new Set([993, 143]);
const ALLOWED_SMTP_PORTS = new Set([465, 587]);
const DEFAULT_PROBE_TIMEOUT_MS = 3_500;
const DEFAULT_MAIL_IO_TIMEOUT_MS = 15_000;

export function sanitizeMailError(message: string, secret?: string): string {
  let out = String(message || '메일 서버 연결 오류가 발생했습니다.');
  if (secret && secret.length > 0) {
    out = out.split(secret).join('[REDACTED]');
  }
  out = out.replace(/PASS\s+[^\r\n]+/gi, 'PASS [REDACTED]');
  out = out.replace(/LOGIN\s+"[^"]*"\s+"[^"]*"/gi, 'LOGIN [REDACTED]');
  return out;
}

export function extractDomainFromEmail(email: string): string {
  const trimmed = String(email || '').trim().toLowerCase();
  if (!trimmed || /[\r\n\0\s]/.test(trimmed)) return '';
  const parts = trimmed.split('@');
  if (parts.length !== 2) return '';
  const [local, domain] = parts;
  if (!local || !domain || !/^[a-z0-9.-]+\.[a-z]{2,}$/i.test(domain)) return '';
  return domain;
}

export function extractLocalPartFromEmail(email: string): string {
  const trimmed = String(email || '').trim();
  const idx = trimmed.indexOf('@');
  return idx > 0 ? trimmed.slice(0, idx) : trimmed;
}

/**
 * Checks whether an IPv4 or IPv6 address belongs to loopback, private RFC1918,
 * link-local / cloud metadata, CGNAT, or unspecified space.
 */
export function isPrivateOrBlockedIp(rawIp: string): boolean {
  const ip = rawIp.trim().toLowerCase();
  if (!ip) return true;

  if (ip.startsWith('::ffff:')) {
    return isPrivateOrBlockedIp(ip.slice(7));
  }

  if (net.isIPv4(ip)) {
    const parts = ip.split('.').map((n) => Number.parseInt(n, 10));
    if (parts.length !== 4 || parts.some((n) => Number.isNaN(n) || n < 0 || n > 255)) return true;
    const [a, b] = parts;
    if (a === 0) return true;
    if (a === 10) return true;
    if (a === 127) return true;
    if (a === 169 && b === 254) return true;
    if (a === 172 && b >= 16 && b <= 31) return true;
    if (a === 192 && b === 168) return true;
    if (a === 100 && b >= 64 && b <= 127) return true;
    if (a >= 224) return true;
    return false;
  }

  if (net.isIPv6(ip)) {
    if (ip === '::' || ip === '::1') return true;
    if (ip.startsWith('fc') || ip.startsWith('fd')) return true;
    if (ip.startsWith('fe8') || ip.startsWith('fe9') || ip.startsWith('fea') || ip.startsWith('feb')) {
      return true;
    }
    if (ip.startsWith('ff')) return true;
    return false;
  }

  return true;
}

export function isValidMailHostname(host: string): boolean {
  const clean = host.trim().toLowerCase();
  if (!clean || clean.length > 253) return false;
  if (/[\r\n\0\s/\\:@]/.test(clean)) return false;
  if (clean === 'localhost' || clean.endsWith('.local') || clean.endsWith('.internal')) return false;
  if (net.isIP(clean)) return !isPrivateOrBlockedIp(clean);
  return /^[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?(?:\.[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?)+$/i.test(clean);
}

/**
 * Resolves host DNS records and ensures every target IP is a public routable address
 * and the target port is an allowed mail port.
 */
export async function assertSafeMailEndpoint(
  host: string,
  port: number,
  protocol: 'pop3' | 'imap' | 'smtp',
  allowLoopbackForTest = false
): Promise<void> {
  const allowedPorts =
    protocol === 'pop3'
      ? ALLOWED_POP3_PORTS
      : protocol === 'imap'
        ? ALLOWED_IMAP_PORTS
        : ALLOWED_SMTP_PORTS;
  if (!allowLoopbackForTest && !allowedPorts.has(port)) {
    throw new Error(`허용되지 않은 ${protocol.toUpperCase()} 포트입니다 (${port}).`);
  }
  if (allowLoopbackForTest && (host === '127.0.0.1' || host === '::1' || host === 'localhost')) {
    return;
  }
  if (!isValidMailHostname(host)) {
    throw new Error(`유효하지 않거나 내부망에 해당하는 메일 서버 호스트입니다 (${host}).`);
  }
  if (net.isIP(host)) {
    if (isPrivateOrBlockedIp(host)) {
      throw new Error('내부 사설 IP 또는 루프백 주소로의 메일 연결은 차단됩니다.');
    }
    return;
  }
  const records = await dns.lookup(host, { all: true });
  if (!records || records.length === 0) {
    throw new Error(`메일 서버 호스트(${host})의 DNS 주소를 찾을 수 없습니다.`);
  }
  for (const rec of records) {
    if (isPrivateOrBlockedIp(rec.address)) {
      throw new Error(`보안 정책상 내부망 주소로 해석되는 메일 호스트(${host})는 연결할 수 없습니다.`);
    }
  }
}

function resolveTemplate(template: string, domain: string): string {
  if (!template) return '';
  return template.replace(/\{domain\}/g, domain);
}

export function getPresetById(id: MailServicePresetId): MailServicePresetInfo {
  return MAIL_SERVICE_PRESETS.find((p) => p.id === id) || MAIL_SERVICE_PRESETS[0];
}

function matchPresetFromMx(mxHosts: string[]): MailServicePresetInfo | null {
  const joined = mxHosts.map((h) => h.toLowerCase());
  for (const preset of MAIL_SERVICE_PRESETS) {
    if (preset.id === 'auto' || preset.id === 'custom' || preset.mxPatterns.length === 0) continue;
    for (const pattern of preset.mxPatterns) {
      if (joined.some((mx) => mx.includes(pattern))) {
        return preset;
      }
    }
  }
  return null;
}

function matchPresetFromKnownDomain(domain: string): MailServicePresetInfo | null {
  const d = domain.toLowerCase();
  if (d === 'gmail.com' || d === 'googlemail.com') return getPresetById('google');
  if (d === 'naver.com') return getPresetById('naver');
  if (d === 'daum.net' || d === 'hanmail.net' || d === 'kakao.com') return getPresetById('daum');
  if (d === 'nate.com' || d === 'empas.com') return getPresetById('nate');
  if (d === 'icloud.com' || d === 'me.com' || d === 'mac.com') return getPresetById('icloud');
  if (d === 'outlook.com' || d === 'hotmail.com' || d === 'live.com') return getPresetById('microsoft');
  if (d === 'zoho.com' || d === 'zohomail.com') return getPresetById('zoho');
  if (d.endsWith('mailplug.co.kr') || d.endsWith('mailplug.com')) return getPresetById('mailplug');
  if (d.endsWith('hiworks.com') || d.endsWith('hiworks.co.kr')) return getPresetById('hiworks');
  if (d.endsWith('worksmobile.com')) return getPresetById('naverworks');
  if (d.endsWith('dooray.com')) return getPresetById('dooray');
  if (d.endsWith('kakaowork.com')) return getPresetById('kakaowork');
  if (d.endsWith('hanbiro.com') || d.endsWith('hanbiro.net')) return getPresetById('hanbiro');
  return null;
}

/**
 * Diagnoses real-world POP3/IMAP authentication errors across all Korean & global mail providers
 * and produces concrete, provider-specific resolution steps.
 */
export function diagnoseMailAuthError(params: {
  providerId?: MailServicePresetId;
  protocol: MailReceiveProtocol;
  rawServerError: string;
}): MailConnectionDiagnosis {
  const preset = getPresetById(params.providerId || 'auto');
  const raw = params.rawServerError || '';
  const lower = raw.toLowerCase();
  const otherProtocol: MailReceiveProtocol | undefined =
    params.protocol === 'pop3' && preset.supportedProtocols.includes('imap')
      ? 'imap'
      : params.protocol === 'imap' && preset.supportedProtocols.includes('pop3')
        ? 'pop3'
        : undefined;

  // 1. Mailplug & qmail/vpopmail #5.7.5 / "password allowed relay" / post-2025-04-01 IMAP restriction
  if (
    params.providerId === 'mailplug' ||
    lower.includes('#5.7.5') ||
    lower.includes('allowed relay')
  ) {
    const checklist = [
      '① 앱 비밀번호 필수: 웹메일 로그인 비밀번호는 외부 연동 시 #5.7.5 오류로 거부됩니다. 메일플러그 웹메일 [우측 상단 프로필 > 로그인 보안 설정 > 앱 비밀번호]에서 전용 비밀번호를 새로 발급해 입력하세요.',
      '② 해외 IP 차단 확인: 현재 클라우드 미리보기 서버는 해외 IP에서 접속합니다. 메일플러그 [로그인 보안 설정 > 해외 IP 차단]이 켜져 있으면 비밀번호가 정확해도 거부됩니다. (일시 해제/국가 허용 또는 사내 Windows .exe 로컬 실행 필요)',
      '③ POP3/IMAP 허용 확인: [웹메일 환경설정 > POP3/IMAP 설정]에서 사용함을 확인하세요. (30일 미사용 시 보안상 자동 비활성화됩니다)',
      params.protocol === 'imap'
        ? '④ 2025년 4월 1일 이후 개통 계정 IMAP 제한: 메일플러그 공식 정책상 2025.04.01 이후 개통된 그룹웨어/기업메일은 IMAP 이용이 제한됩니다. POP3(pop3.mailplug.co.kr:995 SSL)로 전환해 연결하세요.'
        : '④ 프로토콜 참고: 2025년 4월 1일 이후 개통된 메일플러그 그룹웨어/기업메일은 IMAP이 제한되므로 현재 POP3(995 SSL) 사용이 적합합니다.',
    ];
    return {
      code: 'overseas-ip-or-policy',
      headline:
        params.protocol === 'imap'
          ? '메일플러그 IMAP 인증 실패: 2025.04.01 이후 개통 계정의 IMAP 제한 또는 [앱 비밀번호]·해외 IP 차단을 확인하세요.'
          : '메일플러그 보안 정책 차단 (#5.7.5): 일반 비밀번호 대신 [앱 비밀번호] 발급 및 해외 IP 차단 설정 확인이 필요합니다.',
      checklist,
      suggestProtocolSwitch: otherProtocol,
      rawServerError: raw,
    };
  }

  // 2. Microsoft 365 Basic Auth blocked
  if (
    params.providerId === 'microsoft' ||
    lower.includes('basicauth') ||
    lower.includes('basic authentication') ||
    lower.includes('oauth')
  ) {
    return {
      code: 'oauth-required',
      headline: 'Microsoft 365 기본 비밀번호(Basic Auth) 인증이 차단되어 있습니다.',
      checklist: [
        '① Microsoft 365 기업 계정은 보안 정책상 기본 비밀번호 POP3/IMAP 접속이 비활성화되어 있습니다.',
        '② 앱 상단의 Microsoft 365 OAuth 연동 또는 Outlook .pst 파일 불러오기를 이용해 주세요.',
      ],
      rawServerError: raw,
    };
  }

  // 3. Explicit App Password / Application-Specific Password required (Gmail, Naver, Daum, Dooray, iCloud, Naver Works)
  if (
    lower.includes('application-specific') ||
    lower.includes('app password') ||
    lower.includes('web login required') ||
    lower.includes('webloginrequired') ||
    preset.appPasswordPolicy === 'required'
  ) {
    const checklist = [
      `① 앱 비밀번호 확인: ${preset.appPasswordGuide}`,
      `② 프로토콜 허용 확인: ${preset.protocolSetupGuide}`,
    ];
    if (preset.overseasIpWarning) {
      checklist.push(
        '③ 해외 IP 로그인 차단 확인: 현재 웹 서버가 해외 IP로 접속하므로 웹메일 보안 설정에서 해외 로그인 차단이 켜져 있으면 인증이 거부됩니다.'
      );
    }
    return {
      code: 'app-password-required',
      headline: `${preset.label} 인증 실패: 전용 앱 비밀번호 또는 외부 연동 허용 설정이 필요합니다.`,
      checklist,
      suggestProtocolSwitch: otherProtocol,
      rawServerError: raw,
    };
  }

  // 4. Protocol explicitly disabled on webmail
  if (
    lower.includes('not enabled') ||
    lower.includes('disabled') ||
    lower.includes('pop3 is not') ||
    lower.includes('imap is not') ||
    lower.includes('허용')
  ) {
    return {
      code: 'protocol-disabled',
      headline: `${preset.label}에서 ${params.protocol.toUpperCase()} 외부 연동이 꺼져 있습니다.`,
      checklist: [
        `① ${preset.protocolSetupGuide}`,
        otherProtocol
          ? `② 현재 ${params.protocol.toUpperCase()} 대신 ${otherProtocol.toUpperCase()} 프로토콜로 전환하여 접속해 보세요.`
          : '② 웹메일 환경설정 저장 후 1~2분 뒤 다시 시도해 주세요.',
      ],
      suggestProtocolSwitch: otherProtocol,
      rawServerError: raw,
    };
  }

  // 5. Generic authentication failure with provider-aware checklist
  const checklist = [
    `① 비밀번호 및 계정 형식 확인: ${preset.appPasswordGuide}`,
    `② 서비스별 연동 조건: ${preset.protocolSetupGuide}`,
  ];
  if (preset.overseasIpWarning) {
    checklist.push(
      '③ 해외 IP 차단 정책: 국내 기업메일은 클라우드(해외 IP) 접속을 차단할 수 있으므로 해외 IP 허용 또는 사내 로컬 .exe 실행을 확인하세요.'
    );
  }
  return {
    code: 'auth-failed',
    headline: `${preset.label} ${params.protocol.toUpperCase()} 인증에 실패했습니다.`,
    checklist,
    suggestProtocolSwitch: otherProtocol,
    rawServerError: raw,
  };
}

export interface MailPortProbeResult {
  reachable: boolean;
  banner?: string;
}

export type Pop3PortProbeResult = MailPortProbeResult;

/**
 * Fast, safe POP3 banner probe (reads first +OK greeting line and closes immediately).
 */
export async function probePop3Endpoint(
  host: string,
  port: number,
  useTls: boolean,
  timeoutMs = DEFAULT_PROBE_TIMEOUT_MS,
  allowLoopbackForTest = false
): Promise<MailPortProbeResult> {
  if (!host) return { reachable: false };
  try {
    await assertSafeMailEndpoint(host, port, 'pop3', allowLoopbackForTest);
  } catch {
    return { reachable: false };
  }

  return new Promise<MailPortProbeResult>((resolve) => {
    let settled = false;
    const finish = (result: MailPortProbeResult, sock?: net.Socket) => {
      if (settled) return;
      settled = true;
      try {
        sock?.destroy();
      } catch {
        // ignore cleanup errors
      }
      resolve(result);
    };

    const socket: net.Socket = useTls
      ? tls.connect({ host, port, servername: host, rejectUnauthorized: true })
      : net.connect({ host, port });

    const timer = setTimeout(() => finish({ reachable: false }, socket), timeoutMs);
    let buffer = '';

    socket.on('data', (chunk) => {
      buffer += chunk.toString('utf8');
      const lineEnd = buffer.indexOf('\n');
      if (lineEnd !== -1 || buffer.length > 256) {
        clearTimeout(timer);
        const firstLine = (lineEnd !== -1 ? buffer.slice(0, lineEnd) : buffer).trim();
        const ok = firstLine.toUpperCase().startsWith('+OK');
        finish({ reachable: ok, banner: firstLine.slice(0, 140) }, socket);
      }
    });

    socket.on('error', () => {
      clearTimeout(timer);
      finish({ reachable: false }, socket);
    });
  });
}

/**
 * Fast, safe IMAP4rev1 banner probe (reads first * OK greeting line and closes immediately).
 */
export async function probeImapEndpoint(
  host: string,
  port: number,
  useTls: boolean,
  timeoutMs = DEFAULT_PROBE_TIMEOUT_MS,
  allowLoopbackForTest = false
): Promise<MailPortProbeResult> {
  if (!host) return { reachable: false };
  try {
    await assertSafeMailEndpoint(host, port, 'imap', allowLoopbackForTest);
  } catch {
    return { reachable: false };
  }

  return new Promise<MailPortProbeResult>((resolve) => {
    let settled = false;
    const finish = (result: MailPortProbeResult, sock?: net.Socket) => {
      if (settled) return;
      settled = true;
      try {
        sock?.destroy();
      } catch {
        // ignore cleanup errors
      }
      resolve(result);
    };

    const socket: net.Socket = useTls
      ? tls.connect({ host, port, servername: host, rejectUnauthorized: true })
      : net.connect({ host, port });

    const timer = setTimeout(() => finish({ reachable: false }, socket), timeoutMs);
    let buffer = '';

    socket.on('data', (chunk) => {
      buffer += chunk.toString('utf8');
      const lineEnd = buffer.indexOf('\n');
      if (lineEnd !== -1 || buffer.length > 256) {
        clearTimeout(timer);
        const firstLine = (lineEnd !== -1 ? buffer.slice(0, lineEnd) : buffer).trim();
        const ok = /^\*\s+(OK|PREAUTH)\b/i.test(firstLine);
        finish({ reachable: ok, banner: firstLine.slice(0, 140) }, socket);
      }
    });

    socket.on('error', () => {
      clearTimeout(timer);
      finish({ reachable: false }, socket);
    });
  });
}

async function probePresetEndpoints(params: {
  preset: MailServicePresetInfo;
  domain: string;
  probePop3: (host: string, port: number, useTls: boolean) => Promise<MailPortProbeResult>;
  probeImap: (host: string, port: number, useTls: boolean) => Promise<MailPortProbeResult>;
}) {
  const { preset, domain, probePop3, probeImap } = params;

  // Resolve primary & alternate POP3 candidates
  let pop3Host = resolveTemplate(preset.pop3HostTemplate, domain);
  let pop3Port = preset.pop3Port;
  let pop3Tls = preset.pop3Tls;
  let pop3Probe: MailPortProbeResult = { reachable: false };

  if (preset.supportedProtocols.includes('pop3') && pop3Host) {
    const pop3Candidates = [
      { host: pop3Host, port: pop3Port, tls: pop3Tls },
      ...(preset.pop3AltTemplates ?? []).map((t) => ({
        host: resolveTemplate(t.hostTemplate, domain),
        port: t.port,
        tls: t.tls,
      })),
    ].filter((c) => Boolean(c.host));

    for (const cand of pop3Candidates) {
      const res = await probePop3(cand.host, cand.port, cand.tls);
      if (res.reachable) {
        pop3Host = cand.host;
        pop3Port = cand.port;
        pop3Tls = cand.tls;
        pop3Probe = res;
        break;
      }
    }
  }

  // Resolve primary & alternate IMAP candidates
  let imapHost = resolveTemplate(preset.imapHostTemplate, domain);
  let imapPort = preset.imapPort;
  let imapTls = preset.imapTls;
  let imapProbe: MailPortProbeResult = { reachable: false };

  if (preset.supportedProtocols.includes('imap') && imapHost) {
    const imapCandidates = [
      { host: imapHost, port: imapPort, tls: imapTls },
      ...(preset.imapAltTemplates ?? []).map((t) => ({
        host: resolveTemplate(t.hostTemplate, domain),
        port: t.port,
        tls: t.tls,
      })),
    ].filter((c) => Boolean(c.host));

    for (const cand of imapCandidates) {
      const res = await probeImap(cand.host, cand.port, cand.tls);
      if (res.reachable) {
        imapHost = cand.host;
        imapPort = cand.port;
        imapTls = cand.tls;
        imapProbe = res;
        break;
      }
    }
  }

  let recommendedProtocol: MailReceiveProtocol = preset.defaultProtocol;
  if (recommendedProtocol === 'imap' && !imapProbe.reachable && pop3Probe.reachable) {
    recommendedProtocol = 'pop3';
  } else if (recommendedProtocol === 'pop3' && !pop3Probe.reachable && imapProbe.reachable) {
    recommendedProtocol = 'imap';
  } else if (!preset.supportedProtocols.includes(recommendedProtocol)) {
    recommendedProtocol = preset.supportedProtocols[0] || 'pop3';
  }

  const smtpHost = resolveTemplate(preset.smtpHostTemplate, domain);

  return {
    pop3Host,
    pop3Port,
    pop3Tls,
    pop3Reachable: pop3Probe.reachable,
    pop3Banner: pop3Probe.banner,
    imapHost,
    imapPort,
    imapTls,
    imapReachable: imapProbe.reachable,
    imapBanner: imapProbe.banner,
    smtpHost,
    smtpPort: preset.smtpPort,
    smtpTls: preset.smtpTls,
    recommendedProtocol,
    reachable: pop3Probe.reachable || imapProbe.reachable,
  };
}

/**
 * Automatically detects company POP3/IMAP/SMTP configuration from email address & optional service preset:
 * 1) Explicit service preset (Mailplug, Hiworks, Cafe24, Dooray, Naver Works, DaouOffice, Hanbiro, etc.)
 * 2) Well-known email domain match
 * 3) Live DNS MX record lookup
 * 4) Multi-host POP3 + IMAP banner probing (imap.<d>, pop3.<d>, mail.<d>, webmail.<d>)
 */
export async function detectCompanyPop3Config(params: {
  email: string;
  providerPreset?: MailServicePresetId;
  resolveMxImpl?: (domain: string) => Promise<Array<{ exchange: string; priority: number }>>;
  probePortImpl?: (host: string, port: number, useTls: boolean) => Promise<MailPortProbeResult>;
  probeImapImpl?: (host: string, port: number, useTls: boolean) => Promise<MailPortProbeResult>;
}): Promise<DiscoveredPop3Config> {
  const email = String(params.email || '').trim().toLowerCase();
  const domain = extractDomainFromEmail(email);
  if (!domain) {
    throw new Error('올바른 이메일 주소를 입력해 주세요. (예: name@company.co.kr)');
  }

  const resolveMx = params.resolveMxImpl ?? ((d: string) => dns.resolveMx(d));
  const probePop3 =
    params.probePortImpl ?? ((h: string, p: number, t: boolean) => probePop3Endpoint(h, p, t));
  const probeImap =
    params.probeImapImpl ?? ((h: string, p: number, t: boolean) => probeImapEndpoint(h, p, t));

  const requestedPreset =
    params.providerPreset && params.providerPreset !== 'auto' ? params.providerPreset : 'auto';

  let mxRecords: string[] = [];
  try {
    const rawMx = await resolveMx(domain);
    mxRecords = rawMx
      .slice()
      .sort((a, b) => a.priority - b.priority)
      .map((r) => r.exchange.replace(/\.$/, '').toLowerCase())
      .filter(Boolean);
  } catch {
    mxRecords = [];
  }

  const buildResultFromPreset = async (
    preset: MailServicePresetInfo,
    detectionMethod: DiscoveredPop3Config['detectionMethod'],
    guidancePrefix?: string
  ): Promise<DiscoveredPop3Config> => {
    const probed = await probePresetEndpoints({ preset, domain, probePop3, probeImap });
    return {
      email,
      domain,
      detectedProviderId: preset.id,
      detectedProviderName: preset.label,
      confidence: 'high',
      detectionMethod,
      mxRecords,
      supportedProtocols: preset.supportedProtocols,
      recommendedProtocol: probed.recommendedProtocol,
      pop3Host: probed.pop3Host,
      pop3Port: probed.pop3Port,
      pop3Tls: probed.pop3Tls,
      pop3Reachable: probed.pop3Reachable,
      pop3Banner: probed.pop3Banner,
      imapHost: probed.imapHost,
      imapPort: probed.imapPort,
      imapTls: probed.imapTls,
      imapReachable: probed.imapReachable,
      imapBanner: probed.imapBanner,
      smtpHost: probed.smtpHost,
      smtpPort: probed.smtpPort,
      smtpTls: probed.smtpTls,
      reachable: probed.reachable,
      usernameFormat: preset.usernameFormat,
      appPasswordPolicy: preset.appPasswordPolicy,
      overseasIpWarning: preset.overseasIpWarning,
      appPasswordGuide: preset.appPasswordGuide,
      protocolSetupGuide: preset.protocolSetupGuide,
      guidanceNote: guidancePrefix ? `${guidancePrefix} ${preset.note}` : preset.note,
    };
  };

  // Case 1: Explicit user-selected preset
  if (requestedPreset !== 'auto' && requestedPreset !== 'custom') {
    return buildResultFromPreset(getPresetById(requestedPreset), 'preset');
  }

  // Case 2: Known consumer/enterprise domain match
  const knownPreset = matchPresetFromKnownDomain(domain);
  if (knownPreset && requestedPreset === 'auto') {
    return buildResultFromPreset(knownPreset, 'domain-known');
  }

  // Case 3: Company custom domain via MX record
  const mxPreset = requestedPreset === 'auto' ? matchPresetFromMx(mxRecords) : null;
  if (mxPreset) {
    return buildResultFromPreset(
      mxPreset,
      'mx-record',
      `MX 레코드(${mxRecords[0] || mxPreset.vendorHint})에서 ${mxPreset.label} 환경을 자동 감지했습니다.`
    );
  }

  // Case 4: Multi-host POP3 + IMAP probing for custom/self-hosted company domains
  const customPreset = getPresetById('custom');
  const imapCandidates = [
    { host: `imap.${domain}`, port: 993, tls: true },
    { host: `mail.${domain}`, port: 993, tls: true },
    ...(mxRecords[0] ? [{ host: mxRecords[0], port: 993, tls: true }] : []),
  ];
  const pop3Candidates = [
    { host: `pop3.${domain}`, port: 995, tls: true },
    { host: `mail.${domain}`, port: 995, tls: true },
    { host: `pop.${domain}`, port: 995, tls: true },
    { host: `webmail.${domain}`, port: 995, tls: true },
    { host: `webmail.${domain}`, port: 110, tls: false },
    { host: `pop3.${domain}`, port: 110, tls: false },
    ...(mxRecords[0] ? [{ host: mxRecords[0], port: 995, tls: true }] : []),
  ];

  let foundImap: { host: string; port: number; tls: boolean; banner?: string } | null = null;
  for (const cand of imapCandidates) {
    const probe = await probeImap(cand.host, cand.port, cand.tls);
    if (probe.reachable) {
      foundImap = { ...cand, banner: probe.banner };
      break;
    }
  }

  let foundPop3: { host: string; port: number; tls: boolean; banner?: string } | null = null;
  for (const cand of pop3Candidates) {
    const probe = await probePop3(cand.host, cand.port, cand.tls);
    if (probe.reachable) {
      foundPop3 = { ...cand, banner: probe.banner };
      break;
    }
  }

  if (foundImap || foundPop3) {
    const primaryHost = foundImap?.host || foundPop3!.host;
    const smtpHost = primaryHost.startsWith('mail.') ? primaryHost : `smtp.${domain}`;
    const supportedProtocols: MailReceiveProtocol[] = [];
    if (foundImap) supportedProtocols.push('imap');
    if (foundPop3) supportedProtocols.push('pop3');
    return {
      email,
      domain,
      detectedProviderId: 'custom',
      detectedProviderName: `자사 메일 서버 (${primaryHost})`,
      confidence: 'medium',
      detectionMethod: 'port-probe',
      mxRecords,
      supportedProtocols: supportedProtocols.length > 0 ? supportedProtocols : ['imap', 'pop3'],
      recommendedProtocol: foundImap ? 'imap' : 'pop3',
      pop3Host: foundPop3?.host ?? `pop3.${domain}`,
      pop3Port: foundPop3?.port ?? 995,
      pop3Tls: foundPop3?.tls ?? true,
      pop3Reachable: Boolean(foundPop3),
      pop3Banner: foundPop3?.banner,
      imapHost: foundImap?.host ?? `imap.${domain}`,
      imapPort: foundImap?.port ?? 993,
      imapTls: foundImap?.tls ?? true,
      imapReachable: Boolean(foundImap),
      imapBanner: foundImap?.banner,
      smtpHost,
      smtpPort: 465,
      smtpTls: true,
      reachable: true,
      usernameFormat: customPreset.usernameFormat,
      appPasswordPolicy: customPreset.appPasswordPolicy,
      overseasIpWarning: false,
      appPasswordGuide: customPreset.appPasswordGuide,
      protocolSetupGuide: customPreset.protocolSetupGuide,
      guidanceNote: `회사 메일 서버(${primaryHost})의 ${foundImap ? 'IMAP' : 'POP3'} 응답을 자동으로 확인했습니다.`,
    };
  }

  // Case 5: Heuristic fallback
  return {
    email,
    domain,
    detectedProviderId: 'custom',
    detectedProviderName: `회사 도메인 기본값 (${domain})`,
    confidence: 'fallback',
    detectionMethod: 'heuristic',
    mxRecords,
    supportedProtocols: ['imap', 'pop3'],
    recommendedProtocol: 'imap',
    pop3Host: `pop3.${domain}`,
    pop3Port: 995,
    pop3Tls: true,
    pop3Reachable: false,
    imapHost: `imap.${domain}`,
    imapPort: 993,
    imapTls: true,
    imapReachable: false,
    smtpHost: `smtp.${domain}`,
    smtpPort: 465,
    smtpTls: true,
    reachable: false,
    usernameFormat: customPreset.usernameFormat,
    appPasswordPolicy: customPreset.appPasswordPolicy,
    overseasIpWarning: false,
    appPasswordGuide: customPreset.appPasswordGuide,
    protocolSetupGuide: customPreset.protocolSetupGuide,
    guidanceNote:
      '자동 포트 응답이 확인되지 않았습니다. 이용 중인 메일 서비스 형태를 선택하거나 고급 설정에서 서버 주소를 직접 입력해 주세요.',
  };
}

/**
 * Low-level binary-safe line & literal socket reader for POP3, IMAP, and SMTP.
 */
class MailLineSocket {
  private buffer = Buffer.alloc(0);
  private readonly waiters: Array<() => void> = [];
  private closed = false;
  private socketError: Error | null = null;

  constructor(private readonly socket: net.Socket, private readonly timeoutMs: number) {
    socket.on('data', (chunk: Buffer) => {
      this.buffer = Buffer.concat([this.buffer, chunk]);
      this.flushWaiters();
    });
    socket.on('error', (err) => {
      this.socketError = err;
      this.closed = true;
      this.flushWaiters();
    });
    socket.on('close', () => {
      this.closed = true;
      this.flushWaiters();
    });
  }

  private flushWaiters(): void {
    while (this.waiters.length > 0) {
      const fn = this.waiters.shift();
      fn?.();
    }
  }

  private waitForData(): Promise<void> {
    if (this.socketError) return Promise.reject(this.socketError);
    if (this.closed) return Promise.reject(new Error('메일 서버 연결이 예기치 않게 종료되었습니다.'));
    return new Promise<void>((resolve, reject) => {
      const timer = setTimeout(() => {
        reject(
          new Error(`메일 서버 응답 시간이 초과되었습니다 (${Math.round(this.timeoutMs / 1000)}초).`)
        );
      }, this.timeoutMs);
      this.waiters.push(() => {
        clearTimeout(timer);
        if (this.socketError) reject(this.socketError);
        else resolve();
      });
    });
  }

  async readLine(): Promise<string> {
    while (true) {
      const idx = this.buffer.indexOf(0x0a); // '\n'
      if (idx !== -1) {
        const line = this.buffer.subarray(0, idx + 1).toString('utf8').replace(/\r?\n$/, '');
        this.buffer = this.buffer.subarray(idx + 1);
        return line;
      }
      if (this.closed) {
        if (this.buffer.length > 0) {
          const rest = this.buffer.toString('utf8').trim();
          this.buffer = Buffer.alloc(0);
          return rest;
        }
        throw this.socketError || new Error('서버 응답을 읽기 전에 연결이 닫혔습니다.');
      }
      await this.waitForData();
    }
  }

  async readExactBytes(byteLength: number): Promise<Buffer> {
    const safeLen = Math.max(0, Math.floor(byteLength));
    while (this.buffer.length < safeLen) {
      if (this.closed) {
        const partial = this.buffer;
        this.buffer = Buffer.alloc(0);
        return partial;
      }
      await this.waitForData();
    }
    const slice = this.buffer.subarray(0, safeLen);
    this.buffer = this.buffer.subarray(safeLen);
    return slice;
  }

  async readMultilineResponse(): Promise<{ statusLine: string; body: string }> {
    const statusLine = await this.readLine();
    if (!statusLine.toUpperCase().startsWith('+OK')) {
      return { statusLine, body: '' };
    }
    const lines: string[] = [];
    let totalBytes = 0;
    const MAX_MESSAGE_BYTES = 4 * 1024 * 1024;
    while (true) {
      const line = await this.readLine();
      if (line === '.') break;
      const unstuffed = line.startsWith('..') ? line.slice(1) : line;
      totalBytes += unstuffed.length + 2;
      if (totalBytes <= MAX_MESSAGE_BYTES) {
        lines.push(unstuffed);
      }
    }
    return { statusLine, body: lines.join('\r\n') };
  }

  send(command: string): void {
    this.socket.write(`${command}\r\n`, 'utf8');
  }

  destroy(): void {
    try {
      this.socket.destroy();
    } catch {
      // ignore
    }
  }
}

async function openMailSocket(params: {
  host: string;
  port: number;
  useTls: boolean;
  timeoutMs: number;
  allowLoopbackForTest: boolean;
  passwordForSanitize: string;
  protocolLabel: string;
}): Promise<net.Socket> {
  return new Promise<net.Socket>((resolve, reject) => {
    const sock: net.Socket = params.useTls
      ? tls.connect({
          host: params.host,
          port: params.port,
          servername: params.host,
          rejectUnauthorized: !params.allowLoopbackForTest,
        })
      : net.connect({ host: params.host, port: params.port });

    const timer = setTimeout(() => {
      sock.destroy();
      reject(
        new Error(
          `${params.protocolLabel} 서버(${params.host}:${params.port}) 연결 시간이 초과되었습니다.`
        )
      );
    }, params.timeoutMs);

    sock.once(params.useTls ? 'secureConnect' : 'connect', () => {
      clearTimeout(timer);
      resolve(sock);
    });
    sock.once('error', (err) => {
      clearTimeout(timer);
      sock.destroy();
      reject(
        new Error(
          `${params.protocolLabel} 서버(${params.host}:${params.port}) 연결 실패: ${sanitizeMailError(
            err.message,
            params.passwordForSanitize
          )}`
        )
      );
    });
  });
}

function buildLoginUsernameCandidates(
  email: string,
  usernameFormat?: UsernameFormatPolicy
): string[] {
  const full = email.trim();
  const local = extractLocalPartFromEmail(full);
  if (usernameFormat === 'local-or-full' && local && local !== full) {
    return [full, local];
  }
  return [full];
}

/**
 * Connects to a real POP3/POP3S server, authenticates with USER/PASS
 * (automatically trying local-part fallback for Naver/Daum/Nate/Custom when applicable),
 * fetches the newest messages via STAT + RETR, and parses them into MailItem[].
 */
export async function fetchPop3Messages(params: Pop3ConnectionParams): Promise<Pop3FetchResult> {
  const email = String(params.email || '').trim();
  const password = String(params.password || '');
  const host = String(params.pop3Host || '').trim();
  const port = Number(params.pop3Port) || (params.pop3Tls ? 995 : 110);
  const timeoutMs = params.timeoutMs ?? DEFAULT_MAIL_IO_TIMEOUT_MS;
  const maxMessages = Math.max(1, Math.min(80, Number(params.maxMessages) || 30));

  if (!email || /[\r\n\0]/.test(email)) {
    throw new Error('유효한 이메일 주소를 입력해 주세요.');
  }
  if (!password || /[\r\n\0]/.test(password)) {
    throw new Error('유효한 비밀번호를 입력해 주세요.');
  }

  await assertSafeMailEndpoint(host, port, 'pop3', Boolean(params.allowLoopbackForTest));

  const usernameCandidates = buildLoginUsernameCandidates(email, params.usernameFormat);
  let lastAuthRawError = '';

  for (let attemptIdx = 0; attemptIdx < usernameCandidates.length; attemptIdx += 1) {
    const loginUser = usernameCandidates[attemptIdx];
    const socket = await openMailSocket({
      host,
      port,
      useTls: params.pop3Tls,
      timeoutMs,
      allowLoopbackForTest: Boolean(params.allowLoopbackForTest),
      passwordForSanitize: password,
      protocolLabel: 'POP3',
    });

    const reader = new MailLineSocket(socket, timeoutMs);
    try {
      const banner = await reader.readLine();
      if (!banner.toUpperCase().startsWith('+OK')) {
        throw new Error(`POP3 서버 인사 응답 오류: ${banner}`);
      }

      reader.send(`USER ${loginUser}`);
      const userResp = await reader.readLine();
      if (!userResp.toUpperCase().startsWith('+OK')) {
        lastAuthRawError = sanitizeMailError(userResp, password);
        if (attemptIdx + 1 < usernameCandidates.length) {
          continue;
        }
        const diagnosis = diagnoseMailAuthError({
          providerId: params.providerId,
          protocol: 'pop3',
          rawServerError: lastAuthRawError,
        });
        throw new MailAuthError(`${diagnosis.headline} (${lastAuthRawError})`, diagnosis);
      }

      reader.send(`PASS ${password}`);
      const passResp = await reader.readLine();
      if (!passResp.toUpperCase().startsWith('+OK')) {
        lastAuthRawError = sanitizeMailError(passResp, password);
        if (attemptIdx + 1 < usernameCandidates.length) {
          continue;
        }
        const diagnosis = diagnoseMailAuthError({
          providerId: params.providerId,
          protocol: 'pop3',
          rawServerError: lastAuthRawError,
        });
        throw new MailAuthError(`${diagnosis.headline} (${lastAuthRawError})`, diagnosis);
      }

      reader.send('STAT');
      const statResp = await reader.readLine();
      const statMatch = statResp.match(/^\+OK\s+(\d+)\s+(\d+)/i);
      const totalMailboxCount = statMatch ? Number.parseInt(statMatch[1], 10) : 0;
      const totalOctets = statMatch ? Number.parseInt(statMatch[2], 10) : 0;

      const emails: MailItem[] = [];
      if (totalMailboxCount > 0) {
        const startIndex = Math.max(1, totalMailboxCount - maxMessages + 1);
        for (let msgNum = totalMailboxCount; msgNum >= startIndex; msgNum -= 1) {
          reader.send(`RETR ${msgNum}`);
          const { statusLine, body } = await reader.readMultilineResponse();
          if (!statusLine.toUpperCase().startsWith('+OK') || !body) {
            continue;
          }
          const parsed = parseEmlMessage(body, {
            idPrefix: `pop3-${host.replace(/[^a-z0-9]/gi, '')}-${msgNum}`,
            folderHint: 'inbox',
            category: 'POP3 수신',
          });
          emails.push({
            ...parsed,
            folderName: `POP3 (${email})`,
            source: 'outlook-live',
          });
        }
      }

      try {
        reader.send('QUIT');
        await reader.readLine();
      } catch {
        // ignore QUIT close race
      }

      return {
        protocolUsed: 'pop3',
        loginUsernameUsed: loginUser,
        totalMailboxCount,
        totalOctets,
        fetchedCount: emails.length,
        emails,
        banner,
      };
    } catch (err) {
      if (err instanceof MailAuthError) throw err;
      throw new Error(sanitizeMailError(err instanceof Error ? err.message : String(err), password));
    } finally {
      reader.destroy();
    }
  }

  const fallbackDiag = diagnoseMailAuthError({
    providerId: params.providerId,
    protocol: 'pop3',
    rawServerError: lastAuthRawError || 'POP3 인증 실패',
  });
  throw new MailAuthError(fallbackDiag.headline, fallbackDiag);
}

function quoteImapString(value: string): string {
  return `"${value.replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`;
}

async function readImapTaggedResponse(
  reader: MailLineSocket,
  tag: string
): Promise<{ status: 'OK' | 'NO' | 'BAD'; statusLine: string; untagged: string[] }> {
  const untagged: string[] = [];
  const prefix = `${tag} `;
  while (true) {
    const line = await reader.readLine();
    if (line.startsWith(prefix)) {
      const rest = line.slice(prefix.length).trim();
      const upper = rest.toUpperCase();
      const status = upper.startsWith('OK') ? 'OK' : upper.startsWith('NO') ? 'NO' : 'BAD';
      return { status, statusLine: line, untagged };
    }
    untagged.push(line);
  }
}

/**
 * Pure Node.js RFC 3501 IMAP4rev1 (993 SSL / 143) client.
 * Uses BODY.PEEK[] so fetching messages NEVER mutates unread (\Seen) state on the server.
 */
export async function fetchImapMessages(params: ImapConnectionParams): Promise<Pop3FetchResult> {
  const email = String(params.email || '').trim();
  const password = String(params.password || '');
  const host = String(params.imapHost || '').trim();
  const port = Number(params.imapPort) || (params.imapTls ? 993 : 143);
  const timeoutMs = params.timeoutMs ?? DEFAULT_MAIL_IO_TIMEOUT_MS;
  const maxMessages = Math.max(1, Math.min(80, Number(params.maxMessages) || 30));

  if (!email || /[\r\n\0]/.test(email)) {
    throw new Error('유효한 이메일 주소를 입력해 주세요.');
  }
  if (!password || /[\r\n\0]/.test(password)) {
    throw new Error('유효한 비밀번호를 입력해 주세요.');
  }

  await assertSafeMailEndpoint(host, port, 'imap', Boolean(params.allowLoopbackForTest));

  const usernameCandidates = buildLoginUsernameCandidates(email, params.usernameFormat);
  let lastAuthRawError = '';

  for (let attemptIdx = 0; attemptIdx < usernameCandidates.length; attemptIdx += 1) {
    const loginUser = usernameCandidates[attemptIdx];
    const socket = await openMailSocket({
      host,
      port,
      useTls: params.imapTls,
      timeoutMs,
      allowLoopbackForTest: Boolean(params.allowLoopbackForTest),
      passwordForSanitize: password,
      protocolLabel: 'IMAP',
    });

    const reader = new MailLineSocket(socket, timeoutMs);
    let tagSeq = 1;
    const nextTag = () => `A${String(tagSeq++).padStart(3, '0')}`;

    try {
      const banner = await reader.readLine();
      if (!/^\*\s+(OK|PREAUTH)\b/i.test(banner)) {
        throw new Error(`IMAP 서버 인사 응답 오류: ${banner}`);
      }

      const loginTag = nextTag();
      reader.send(`${loginTag} LOGIN ${quoteImapString(loginUser)} ${quoteImapString(password)}`);
      const loginResp = await readImapTaggedResponse(reader, loginTag);
      if (loginResp.status !== 'OK') {
        lastAuthRawError = sanitizeMailError(loginResp.statusLine, password);
        if (attemptIdx + 1 < usernameCandidates.length) {
          continue;
        }
        const diagnosis = diagnoseMailAuthError({
          providerId: params.providerId,
          protocol: 'imap',
          rawServerError: lastAuthRawError,
        });
        throw new MailAuthError(`${diagnosis.headline} (${lastAuthRawError})`, diagnosis);
      }

      const selectTag = nextTag();
      reader.send(`${selectTag} SELECT INBOX`);
      const selectResp = await readImapTaggedResponse(reader, selectTag);
      if (selectResp.status !== 'OK') {
        throw new Error(`IMAP INBOX 선택 실패: ${selectResp.statusLine}`);
      }

      let totalMailboxCount = 0;
      for (const line of selectResp.untagged) {
        const m = line.match(/^\*\s+(\d+)\s+EXISTS\b/i);
        if (m) {
          totalMailboxCount = Number.parseInt(m[1], 10);
        }
      }

      const emails: MailItem[] = [];
      let totalOctets = 0;

      if (totalMailboxCount > 0) {
        const startSeq = Math.max(1, totalMailboxCount - maxMessages + 1);
        const MAX_LITERAL_BYTES = 4 * 1024 * 1024;

        for (let seq = totalMailboxCount; seq >= startSeq; seq -= 1) {
          const fetchTag = nextTag();
          reader.send(`${fetchTag} FETCH ${seq} (FLAGS BODY.PEEK[])`);

          let rawBody = '';
          let isSeen = false;
          let isFlagged = false;

          while (true) {
            const line = await reader.readLine();
            if (line.startsWith(`${fetchTag} `)) {
              break;
            }
            if (/\\Seen\b/i.test(line)) isSeen = true;
            if (/\\Flagged\b/i.test(line)) isFlagged = true;

            const litMatch = line.match(/\{(\d+)\}\s*$/);
            if (litMatch) {
              const byteCount = Number.parseInt(litMatch[1], 10);
              const literalBuf = await reader.readExactBytes(byteCount);
              totalOctets += byteCount;
              if (byteCount <= MAX_LITERAL_BYTES) {
                rawBody = literalBuf.toString('utf8');
              } else {
                rawBody = literalBuf.subarray(0, MAX_LITERAL_BYTES).toString('utf8');
              }
            }
          }

          if (!rawBody) continue;

          const parsed = parseEmlMessage(rawBody, {
            idPrefix: `imap-${host.replace(/[^a-z0-9]/gi, '')}-${seq}`,
            folderHint: 'inbox',
            category: 'IMAP 수신',
          });
          emails.push({
            ...parsed,
            isRead: isSeen,
            isFlagged,
            folderName: `IMAP (${email})`,
            source: 'outlook-live',
          });
        }
      }

      try {
        const logoutTag = nextTag();
        reader.send(`${logoutTag} LOGOUT`);
        await readImapTaggedResponse(reader, logoutTag);
      } catch {
        // ignore logout close race
      }

      return {
        protocolUsed: 'imap',
        loginUsernameUsed: loginUser,
        totalMailboxCount,
        totalOctets,
        fetchedCount: emails.length,
        emails,
        banner,
      };
    } catch (err) {
      if (err instanceof MailAuthError) throw err;
      throw new Error(sanitizeMailError(err instanceof Error ? err.message : String(err), password));
    } finally {
      reader.destroy();
    }
  }

  const fallbackDiag = diagnoseMailAuthError({
    providerId: params.providerId,
    protocol: 'imap',
    rawServerError: lastAuthRawError || 'IMAP 인증 실패',
  });
  throw new MailAuthError(fallbackDiag.headline, fallbackDiag);
}

function assertNoCrlf(value: string, fieldName: string): string {
  if (/[\r\n\0]/.test(value)) {
    throw new Error(`${fieldName}에 허용되지 않는 줄바꿈/제어 문자가 포함되어 있습니다.`);
  }
  return value.trim();
}

function encodeMimeHeaderUtf8(value: string): string {
  if (/^[\x20-\x7E]*$/.test(value)) return value;
  return `=?UTF-8?B?${Buffer.from(value, 'utf8').toString('base64')}?=`;
}

async function readSmtpReply(reader: MailLineSocket): Promise<{ code: number; text: string }> {
  const lines: string[] = [];
  let code = 0;
  while (true) {
    const line = await reader.readLine();
    lines.push(line);
    const match = line.match(/^(\d{3})([ -])(.*)$/);
    if (match) {
      code = Number.parseInt(match[1], 10);
      if (match[2] === ' ') break;
    } else {
      break;
    }
  }
  return { code, text: lines.join('\n') };
}

function wrapBase6476(b64: string): string {
  const clean = b64.replace(/\s+/g, '');
  const lines: string[] = [];
  for (let i = 0; i < clean.length; i += 76) {
    lines.push(clean.slice(i, i + 76));
  }
  return lines.join('\r\n');
}

/**
 * Builds an RFC 5322 / MIME SMTP DATA payload. When `htmlBody` is provided, emits
 * `multipart/alternative` (`text/plain` + `text/html` or `multipart/related` with
 * inline `cid:` parts extracted from embedded `data:image/...;base64,...` business card images).
 */
export function buildSmtpMimePayload(params: {
  from: string;
  to: string[];
  cc?: string[];
  subject: string;
  textBody: string;
  htmlBody?: string;
}): string {
  const ccList = params.cc || [];
  const baseHeaders = [
    `From: <${params.from}>`,
    `To: ${params.to.join(', ')}`,
    ...(ccList.length > 0 ? [`Cc: ${ccList.join(', ')}`] : []),
    `Subject: ${encodeMimeHeaderUtf8(params.subject)}`,
    `Date: ${new Date().toUTCString()}`,
    'MIME-Version: 1.0',
  ];

  const html = String(params.htmlBody || '').trim();
  const normalizedText = String(params.textBody || '').replace(/\r?\n/g, '\r\n');

  if (!html) {
    const dotStuffedBody = normalizedText.replace(/(^|\r\n)\./g, '$1..');
    return [
      ...baseHeaders,
      'Content-Type: text/plain; charset=UTF-8',
      'Content-Transfer-Encoding: 8bit',
      '',
      dotStuffedBody,
      '.',
    ].join('\r\n');
  }

  const inlineImages: Array<{ cid: string; mimeType: string; base64Data: string }> = [];
  let imgCounter = 0;
  const cidHtml = html.replace(
    /(src\s*=\s*)(["'])(data:image\/([a-z0-9+.-]+);base64,([A-Za-z0-9+/=\s]+))\2/gi,
    (_full, attrPrefix: string, quote: string, _uri: string, subtype: string, rawB64: string) => {
      imgCounter += 1;
      const cid = `img${imgCounter}@bmail.local`;
      const cleanMime = `image/${String(subtype || 'png').toLowerCase()}`;
      inlineImages.push({
        cid,
        mimeType: cleanMime,
        base64Data: rawB64.replace(/\s+/g, ''),
      });
      return `${attrPrefix}${quote}cid:${cid}${quote}`;
    }
  );

  const randSuffix = crypto.randomBytes(6).toString('hex');
  const altBoundary = `----=_Alt_${randSuffix}`;
  const relBoundary = `----=_Rel_${randSuffix}`;

  const textPart = [
    `--${altBoundary}`,
    'Content-Type: text/plain; charset=UTF-8',
    'Content-Transfer-Encoding: base64',
    '',
    wrapBase6476(Buffer.from(normalizedText || '(HTML 메일)', 'utf8').toString('base64')),
  ].join('\r\n');

  const wrappedHtml = `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body style="font-family:'Malgun Gothic','Segoe UI',sans-serif;font-size:10pt;color:#242424;">${cidHtml}</body></html>`;
  const htmlB64 = wrapBase6476(Buffer.from(wrappedHtml, 'utf8').toString('base64'));

  let richSection: string;
  if (inlineImages.length > 0) {
    const imgParts = inlineImages.map((img) =>
      [
        `--${relBoundary}`,
        `Content-Type: ${img.mimeType}; name="${img.cid.split('@')[0]}"`,
        'Content-Transfer-Encoding: base64',
        `Content-ID: <${img.cid}>`,
        'Content-Disposition: inline',
        '',
        wrapBase6476(img.base64Data),
      ].join('\r\n')
    );

    richSection = [
      `--${altBoundary}`,
      `Content-Type: multipart/related; type="text/html"; boundary="${relBoundary}"`,
      '',
      `--${relBoundary}`,
      'Content-Type: text/html; charset=UTF-8',
      'Content-Transfer-Encoding: base64',
      '',
      htmlB64,
      ...imgParts,
      `--${relBoundary}--`,
    ].join('\r\n');
  } else {
    richSection = [
      `--${altBoundary}`,
      'Content-Type: text/html; charset=UTF-8',
      'Content-Transfer-Encoding: base64',
      '',
      htmlB64,
    ].join('\r\n');
  }

  const fullMime = [
    ...baseHeaders,
    `Content-Type: multipart/alternative; boundary="${altBoundary}"`,
    '',
    textPart,
    richSection,
    `--${altBoundary}--`,
    '',
  ].join('\r\n');

  const dotStuffedMime = fullMime.replace(/(^|\r\n)\./g, '$1..');
  return `${dotStuffedMime}\r\n.`;
}

/**
 * Sends a UTF-8 email (plain text or multipart HTML + inline cid business card images)
 * via SMTP over TLS (465) or STARTTLS (587) using the connected company mail credentials.
 */
export async function sendSmtpMessage(params: SmtpSendParams): Promise<{ acceptedTo: string[] }> {
  const email = assertNoCrlf(params.email, '발신 이메일');
  const password = String(params.password || '');
  if (!password || /[\r\n\0]/.test(password)) {
    throw new Error('유효한 SMTP 비밀번호가 없습니다.');
  }

  const recipients = params.to
    .map((addr) => assertNoCrlf(addr, '받는 사람'))
    .filter((addr) => extractDomainFromEmail(addr) !== '');
  const ccList = (params.cc || [])
    .map((addr) => assertNoCrlf(addr, '참조'))
    .filter((addr) => extractDomainFromEmail(addr) !== '');

  if (recipients.length === 0) {
    throw new Error('유효한 수신자 이메일 주소를 1개 이상 입력해 주세요.');
  }

  const subject = assertNoCrlf(params.subject || '(제목 없음)', '제목');
  const body = String(params.body || '').replace(/\r?\n/g, '\r\n');
  const host = String(params.smtpHost || '').trim();
  const port = Number(params.smtpPort) || (params.smtpTls ? 465 : 587);
  const timeoutMs = params.timeoutMs ?? DEFAULT_MAIL_IO_TIMEOUT_MS;

  await assertSafeMailEndpoint(host, port, 'smtp', Boolean(params.allowLoopbackForTest));

  let socket = await openMailSocket({
    host,
    port,
    useTls: params.smtpTls,
    timeoutMs,
    allowLoopbackForTest: Boolean(params.allowLoopbackForTest),
    passwordForSanitize: password,
    protocolLabel: 'SMTP',
  });

  let reader = new MailLineSocket(socket, timeoutMs);
  try {
    const greeting = await readSmtpReply(reader);
    if (greeting.code < 200 || greeting.code >= 300) {
      throw new Error(`SMTP 인사 응답 오류: ${greeting.text}`);
    }

    reader.send('EHLO bmail.local');
    let ehlo = await readSmtpReply(reader);
    if (ehlo.code < 200 || ehlo.code >= 300) {
      throw new Error(`SMTP EHLO 실패: ${ehlo.text}`);
    }

    if (!params.smtpTls && ehlo.text.toUpperCase().includes('STARTTLS')) {
      reader.send('STARTTLS');
      const startTlsResp = await readSmtpReply(reader);
      if (startTlsResp.code === 220) {
        socket = await new Promise<tls.TLSSocket>((resolve, reject) => {
          const upgraded = tls.connect({
            socket,
            servername: host,
            rejectUnauthorized: !params.allowLoopbackForTest,
          });
          upgraded.once('secureConnect', () => resolve(upgraded));
          upgraded.once('error', reject);
        });
        reader = new MailLineSocket(socket, timeoutMs);
        reader.send('EHLO bmail.local');
        ehlo = await readSmtpReply(reader);
      }
    }

    reader.send('AUTH LOGIN');
    const authStep1 = await readSmtpReply(reader);
    if (authStep1.code !== 334) {
      throw new Error(`SMTP AUTH LOGIN 지원 오류: ${authStep1.text}`);
    }
    reader.send(Buffer.from(email, 'utf8').toString('base64'));
    const authStep2 = await readSmtpReply(reader);
    if (authStep2.code !== 334) {
      throw new Error('SMTP 계정 인증에 실패했습니다.');
    }
    reader.send(Buffer.from(password, 'utf8').toString('base64'));
    const authStep3 = await readSmtpReply(reader);
    if (authStep3.code !== 235) {
      throw new Error('SMTP 비밀번호 인증에 실패했습니다. 웹메일 SMTP 허용 및 앱 비밀번호를 확인하세요.');
    }

    reader.send(`MAIL FROM:<${email}>`);
    const mailFromResp = await readSmtpReply(reader);
    if (mailFromResp.code < 200 || mailFromResp.code >= 300) {
      throw new Error(`SMTP 발신자 거부: ${mailFromResp.text}`);
    }

    const allRecipients = [...recipients, ...ccList];
    for (const rcpt of allRecipients) {
      reader.send(`RCPT TO:<${rcpt}>`);
      const rcptResp = await readSmtpReply(reader);
      if (rcptResp.code < 200 || rcptResp.code >= 300) {
        throw new Error(`수신자(${rcpt}) 거부: ${rcptResp.text}`);
      }
    }

    reader.send('DATA');
    const dataResp = await readSmtpReply(reader);
    if (dataResp.code !== 354) {
      throw new Error(`SMTP DATA 시작 오류: ${dataResp.text}`);
    }

    const mimePayload = buildSmtpMimePayload({
      from: email,
      to: recipients,
      cc: ccList,
      subject,
      textBody: body,
      htmlBody: params.htmlBody,
    });

    reader.send(mimePayload);
    const sendResp = await readSmtpReply(reader);
    if (sendResp.code < 200 || sendResp.code >= 300) {
      throw new Error(`메일 발송이 거부되었습니다: ${sendResp.text}`);
    }

    try {
      reader.send('QUIT');
      await readSmtpReply(reader);
    } catch {
      // ignore
    }

    return { acceptedTo: allRecipients };
  } catch (err) {
    throw new Error(sanitizeMailError(err instanceof Error ? err.message : String(err), password));
  } finally {
    reader.destroy();
  }
}

export interface Pop3ServerSession {
  token: string;
  email: string;
  password: string;
  providerId: MailServicePresetId;
  providerName: string;
  protocol: MailReceiveProtocol;
  usernameFormat: UsernameFormatPolicy;
  pop3Host: string;
  pop3Port: number;
  pop3Tls: boolean;
  imapHost: string;
  imapPort: number;
  imapTls: boolean;
  smtpHost: string;
  smtpPort: number;
  smtpTls: boolean;
  createdAt: number;
  expiresAt: number;
}

const SESSION_TTL_MS = 30 * 24 * 60 * 60 * 1000;
const MAX_ACTIVE_SESSIONS = 100;

interface EncryptedSessionEnvelope {
  version: 1;
  iv: string;
  tag: string;
  ciphertext: string;
}

function deriveLocalSessionKey(explicitSecret?: string): Buffer {
  const secret =
    explicitSecret?.trim() ||
    process.env.EMV_MASTER_KEY?.trim() ||
    process.env.BMAIL_SESSION_KEY?.trim() ||
    `bmail-local-session:${os.hostname()}:${os.userInfo().username}`;
  return crypto.scryptSync(secret, 'bmail-pop3-session-salt-v1', 32);
}

/**
 * Encrypted persistent POP3/IMAP/SMTP session store (30-day sliding TTL).
 * When `baseDir` is provided, sessions are stored strictly AES-256-GCM encrypted at rest.
 */
export class Pop3SessionStore {
  private readonly sessions = new Map<string, Pop3ServerSession>();
  private readonly filePath: string | null;
  private readonly encryptionKey: Buffer;

  constructor(options?: { baseDir?: string; secretKey?: string }) {
    this.filePath = options?.baseDir
      ? path.join(path.resolve(options.baseDir), 'pop3-sessions.enc.json')
      : null;
    this.encryptionKey = deriveLocalSessionKey(options?.secretKey);
    if (this.filePath) {
      this.loadFromDisk();
    }
  }

  create(data: Omit<Pop3ServerSession, 'token' | 'createdAt' | 'expiresAt'>): Pop3ServerSession {
    this.purgeExpired(false);
    if (this.sessions.size >= MAX_ACTIVE_SESSIONS) {
      const oldestKey = this.sessions.keys().next().value;
      if (oldestKey) this.sessions.delete(oldestKey);
    }
    const token = crypto.randomBytes(32).toString('hex');
    const now = Date.now();
    const record: Pop3ServerSession = {
      ...data,
      token,
      createdAt: now,
      expiresAt: now + SESSION_TTL_MS,
    };
    this.sessions.set(token, record);
    this.saveToDisk();
    return record;
  }

  get(token?: string | null): Pop3ServerSession | null {
    if (!token) return null;
    const rec = this.sessions.get(token);
    if (!rec) return null;
    const now = Date.now();
    if (now > rec.expiresAt) {
      this.sessions.delete(token);
      this.saveToDisk();
      return null;
    }
    // Sliding TTL: keep active user logged in
    if (rec.expiresAt - now < SESSION_TTL_MS - 60_000) {
      rec.expiresAt = now + SESSION_TTL_MS;
      this.saveToDisk();
    }
    return rec;
  }

  getLatest(): Pop3ServerSession | null {
    this.purgeExpired(true);
    let latest: Pop3ServerSession | null = null;
    for (const rec of this.sessions.values()) {
      if (!latest || rec.expiresAt > latest.expiresAt || rec.createdAt > latest.createdAt) {
        latest = rec;
      }
    }
    if (!latest) return null;
    const now = Date.now();
    if (latest.expiresAt - now < SESSION_TTL_MS - 60_000) {
      latest.expiresAt = now + SESSION_TTL_MS;
      this.saveToDisk();
    }
    return latest;
  }

  listAccounts(): Array<{
    token: string;
    account: ReturnType<typeof toPublicPop3AccountInfo>;
    expiresAt: number;
  }> {
    this.purgeExpired(true);
    return Array.from(this.sessions.values()).map((rec) => ({
      token: rec.token,
      account: toPublicPop3AccountInfo(rec),
      expiresAt: rec.expiresAt,
    }));
  }

  delete(token?: string | null): boolean {
    if (!token) return false;
    const deleted = this.sessions.delete(token);
    if (deleted) {
      this.saveToDisk();
    }
    return deleted;
  }

  private purgeExpired(persistIfChanged = false): void {
    const now = Date.now();
    let changed = false;
    for (const [k, v] of this.sessions.entries()) {
      if (now > v.expiresAt) {
        this.sessions.delete(k);
        changed = true;
      }
    }
    if (changed && persistIfChanged) {
      this.saveToDisk();
    }
  }

  private saveToDisk(): void {
    if (!this.filePath) return;
    try {
      fs.mkdirSync(path.dirname(this.filePath), { recursive: true });
      const payload = JSON.stringify(Array.from(this.sessions.values()));
      const iv = crypto.randomBytes(12);
      const cipher = crypto.createCipheriv('aes-256-gcm', this.encryptionKey, iv);
      const encrypted = Buffer.concat([cipher.update(payload, 'utf8'), cipher.final()]);
      const tag = cipher.getAuthTag();
      const envelope: EncryptedSessionEnvelope = {
        version: 1,
        iv: iv.toString('hex'),
        tag: tag.toString('hex'),
        ciphertext: encrypted.toString('hex'),
      };
      fs.writeFileSync(this.filePath, JSON.stringify(envelope), 'utf8');
    } catch {
      // Best-effort persistence: never crash in-memory session operation
    }
  }

  private loadFromDisk(): void {
    if (!this.filePath || !fs.existsSync(this.filePath)) return;
    try {
      const raw = fs.readFileSync(this.filePath, 'utf8');
      const env = JSON.parse(raw) as Partial<EncryptedSessionEnvelope>;
      if (env.version !== 1 || !env.iv || !env.tag || !env.ciphertext) return;
      const decipher = crypto.createDecipheriv(
        'aes-256-gcm',
        this.encryptionKey,
        Buffer.from(env.iv, 'hex')
      );
      decipher.setAuthTag(Buffer.from(env.tag, 'hex'));
      const plain = Buffer.concat([
        decipher.update(Buffer.from(env.ciphertext, 'hex')),
        decipher.final(),
      ]).toString('utf8');
      const list = JSON.parse(plain);
      if (!Array.isArray(list)) return;
      const now = Date.now();
      for (const item of list) {
        if (item && typeof item === 'object' && typeof item.token === 'string' && Number(item.expiresAt) > now) {
          this.sessions.set(item.token, item as Pop3ServerSession);
        }
      }
    } catch {
      // Ignore corrupt or different-machine encrypted file safely
    }
  }
}

export function toPublicPop3AccountInfo(session: Pop3ServerSession) {
  return {
    email: session.email,
    providerId: session.providerId,
    providerName: session.providerName,
    protocol: session.protocol,
    pop3Host: session.protocol === 'imap' ? session.imapHost : session.pop3Host,
    pop3Port: session.protocol === 'imap' ? session.imapPort : session.pop3Port,
    pop3Tls: session.protocol === 'imap' ? session.imapTls : session.pop3Tls,
    imapHost: session.imapHost,
    imapPort: session.imapPort,
    imapTls: session.imapTls,
    smtpHost: session.smtpHost,
    smtpPort: session.smtpPort,
    smtpTls: session.smtpTls,
  };
}

// ============================================================================
// Smart Auto-Heal Engine & Live Server Egress IP/Country Diagnostics
// ============================================================================

const COUNTRY_NAME_KO: Record<string, string> = {
  KR: '대한민국',
  TW: '대만',
  US: '미국',
  JP: '일본',
  SG: '싱가포르',
  HK: '홍콩',
  DE: '독일',
  GB: '영국',
  NL: '네덜란드',
  BE: '벨기에',
  AU: '호주',
  IN: '인도',
};

let cachedEgressInfo: { value: ServerEgressInfo; expiresAt: number } | null = null;
const EGRESS_CACHE_TTL_MS = 10 * 60 * 1000;

/**
 * Safely detects the current server's public IPv4 and country so users can
 * whitelist the exact cloud IP/country when Korean mail servers block foreign IPs.
 */
export async function detectServerEgressInfo(timeoutMs = 2500): Promise<ServerEgressInfo> {
  if (cachedEgressInfo && Date.now() < cachedEgressInfo.expiresAt) {
    return cachedEgressInfo.value;
  }

  const fallback: ServerEgressInfo = {
    ipv4: null,
    countryCode: null,
    countryName: null,
    city: null,
    isKorea: null,
  };

  const fetchJsonWithTimeout = async (url: string): Promise<Record<string, unknown> | null> => {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    try {
      const resp = await fetch(url, {
        signal: controller.signal,
        headers: { Accept: 'application/json' },
      });
      if (!resp.ok) return null;
      return (await resp.json()) as Record<string, unknown>;
    } catch {
      return null;
    } finally {
      clearTimeout(timer);
    }
  };

  try {
    const [ip4Res, geoRes] = await Promise.all([
      fetchJsonWithTimeout('https://api4.ipify.org?format=json'),
      fetchJsonWithTimeout('https://ipwho.is/'),
    ]);

    const ipv4 =
      typeof ip4Res?.ip === 'string' && net.isIPv4(ip4Res.ip)
        ? ip4Res.ip
        : typeof geoRes?.ip === 'string' && net.isIPv4(geoRes.ip)
          ? geoRes.ip
          : null;

    const countryCode =
      typeof geoRes?.country_code === 'string' && geoRes.country_code.length === 2
        ? geoRes.country_code.toUpperCase()
        : null;
    const rawCountry = typeof geoRes?.country === 'string' ? geoRes.country : null;
    const countryName = countryCode
      ? COUNTRY_NAME_KO[countryCode] || rawCountry || countryCode
      : rawCountry;
    const city = typeof geoRes?.city === 'string' ? geoRes.city : null;
    const isKorea = countryCode ? countryCode === 'KR' : null;

    const value: ServerEgressInfo = { ipv4, countryCode, countryName, city, isKorea };
    cachedEgressInfo = { value, expiresAt: Date.now() + EGRESS_CACHE_TTL_MS };
    return value;
  } catch {
    return fallback;
  }
}

/**
 * Generates deduplicated password normalization candidates to automatically heal
 * copy-paste whitespace or hyphens commonly embedded in 16-character App Passwords.
 */
export function buildPasswordCandidates(
  rawPassword: string
): Array<{ value: string; label: string; isNormalized: boolean }> {
  const candidates: Array<{ value: string; label: string; isNormalized: boolean }> = [];
  const seen = new Set<string>();
  const add = (value: string, label: string, isNormalized: boolean) => {
    if (!value || /[\r\n\0]/.test(value) || seen.has(value)) return;
    seen.add(value);
    candidates.push({ value, label, isNormalized });
  };

  const cleanLine = String(rawPassword || '').replace(/[\r\n]+/g, '');
  add(cleanLine, '입력한 비밀번호 원문', false);

  const trimmed = cleanLine.trim();
  if (trimmed !== cleanLine) {
    add(trimmed, '앞뒤 공백 자동 제거', true);
  }

  const noSpaces = trimmed.replace(/\s+/g, '');
  if (noSpaces !== trimmed) {
    add(noSpaces, '앱 비밀번호 공백 자동 제거', true);
  }

  const noSpacesOrHyphens = trimmed.replace(/[\s-]+/g, '');
  if (noSpacesOrHyphens !== noSpaces && noSpacesOrHyphens !== trimmed) {
    add(noSpacesOrHyphens, '앱 비밀번호 공백·하이픈(-) 자동 제거', true);
  }

  return candidates;
}

/**
 * Produces exact, step-by-step interactive resolution instructions tailored to the
 * provider, error code, and live detected server IP/country.
 */
export function buildDetailedResolutionSteps(params: {
  providerId: MailServicePresetId;
  protocol: MailReceiveProtocol;
  rawServerError: string;
  serverEgress?: ServerEgressInfo;
}): ResolutionGuideStep[] {
  const preset = getPresetById(params.providerId);
  const egress = params.serverEgress;
  const countryBadge =
    egress?.countryName && egress?.countryCode
      ? `${egress.countryName} (${egress.countryCode})`
      : '해외(클라우드 서버)';

  if (params.providerId === 'mailplug' || params.rawServerError.includes('#5.7.5')) {
    const steps: ResolutionGuideStep[] = [];

    if (egress?.isKorea === false) {
      steps.push({
        id: 'mailplug-overseas-ip',
        priority: 'critical',
        title: `1단계 (가장 흔한 원인): 해외 IP 차단 해제 — 현재 접속 국가: ${countryBadge}`,
        menuPath: '메일플러그 웹메일 > 우측 상단 프로필(또는 설정) > [보안 설정] > [해외 로그인 차단 / 국가·IP 접근 설정]',
        detail: egress?.ipv4
          ? `앱 비밀번호가 정확해도 현재 클라우드 서버가 ${countryBadge} IP(${egress.ipv4})에서 접속하기 때문에 메일플러그가 #5.7.5로 즉시 차단합니다. [해외 로그인 차단]에서 '${egress.countryName || '대만/미국'}'을 허용하거나 아래 접속 IP를 예외 IP에 추가한 뒤 저장하세요. (회사 관리자 정책으로 잠겨 있다면 사내 관리자에게 예외 IP 등록을 요청하거나 Windows 로컬 .exe에서 실행하세요.)`
          : `앱 비밀번호가 정확해도 현재 클라우드 서버가 해외 IP에서 접속하기 때문에 메일플러그가 #5.7.5로 차단합니다. [해외 로그인 차단] 설정을 일시 허용하거나 Windows 로컬 .exe(한국 사내 IP)에서 실행하세요.`,
        copyValue: egress?.ipv4 || undefined,
        copyLabel: egress?.ipv4 ? `서버 접속 IP 복사 (${egress.ipv4})` : undefined,
      });
    }

    steps.push(
      {
        id: 'mailplug-pop3-enable',
        priority: 'critical',
        title: '2단계: 웹메일 POP3 사용 허용 ON 확인 (30일 미사용 시 자동 OFF)',
        menuPath: '메일플러그 웹메일 > 좌측 하단 [환경설정] > [POP3/IMAP 설정]',
        detail:
          '최초 계정이나 30일간 외부 프로그램 사용 이력이 없는 경우 보안상 자동으로 [사용 안 함]으로 꺼집니다. 반드시 [POP3 사용함]을 체크하고 [저장] 버튼을 눌러 주세요. (※ 2025년 4월 1일 이후 개통된 그룹웨어/기업메일은 IMAP 이용이 제한되므로 POP3 995 SSL을 사용해야 합니다.)',
        helpUrl: 'https://www.mailplug-help.com/',
        helpLabel: '메일플러그 고객센터 열기',
      },
      {
        id: 'mailplug-app-pw',
        priority: 'recommended',
        title: '3단계: 앱 비밀번호 재발급 및 관리자 POP3 차단 여부 확인',
        menuPath: '메일플러그 웹메일 > 우측 상단 프로필 > [로그인 보안 설정] > [앱 비밀번호]',
        detail:
          '기존 앱 비밀번호가 만료되었거나 회사 관리자 페이지([관리자 > 메일 관리 > POP3/IMAP 사용 설정])에서 사원 외부 연동을 막아 둔 경우에도 동일한 #5.7.5 오류가 발생합니다. 새 앱 비밀번호를 발급받아 아래에 붙여넣고 [자동 해결로 다시 연결]을 누르세요.',
      },
      {
        id: 'mailplug-local-fallback',
        priority: 'alternative',
        title: '4단계 (사내 정책으로 클라우드 연동 불가 시 즉시 대안)',
        menuPath: '현재 PC Outlook 파일 불러오기 (.pst / .ost / .eml) 또는 Windows 설치형 .exe',
        detail:
          '사내 보안 정책상 해외 클라우드 접속을 열 수 없는 경우, 하단의 [원본 .pst/.eml 불러오기]로 즉시 메일을 분석하거나 사내 PC(한국 IP)에서 설치형 .exe로 실행하면 해외 IP 차단 없이 바로 연결됩니다.',
      }
    );

    return steps;
  }

  const genericSteps: ResolutionGuideStep[] = [];
  if (preset.overseasIpWarning && egress?.isKorea !== true) {
    genericSteps.push({
      id: 'generic-overseas-ip',
      priority: 'critical',
      title: `1단계: 해외 IP 로그인 차단 확인 (현재 접속: ${countryBadge})`,
      menuPath: `${preset.label} 웹메일 > [보안 설정] > [해외 로그인 차단 / 접속 IP 관리]`,
      detail: egress?.ipv4
        ? `현재 클라우드 서버가 ${countryBadge} IP(${egress.ipv4})로 접속 중입니다. 국내 메일 서비스의 해외 로그인 차단이 켜져 있으면 앱 비밀번호가 맞아도 거부됩니다.`
        : `현재 클라우드 서버가 해외 IP로 접속 중입니다. 웹메일 보안 설정에서 해외 로그인 차단을 확인해 주세요.`,
      copyValue: egress?.ipv4 || undefined,
      copyLabel: egress?.ipv4 ? `서버 접속 IP 복사 (${egress.ipv4})` : undefined,
    });
  }

  genericSteps.push(
    {
      id: 'generic-app-pw',
      priority: 'critical',
      title: `${genericSteps.length + 1}단계: 비밀번호 / 앱 비밀번호 확인`,
      menuPath: `${preset.label} 웹메일 > [보안 설정 / 앱 비밀번호]`,
      detail: preset.appPasswordGuide,
    },
    {
      id: 'generic-proto',
      priority: 'recommended',
      title: `${genericSteps.length + 2}단계: POP3 / IMAP 외부 연동 허용 확인`,
      menuPath: `${preset.label} 웹메일 > [환경설정] > [POP3/IMAP 설정]`,
      detail: preset.protocolSetupGuide,
    }
  );

  return genericSteps;
}

function extractCompactServerCode(rawError: string): string {
  const m = rawError.match(/\(#\d+\.\d+\.\d+\)|\b553\b|\b535\b|\bNO\b|\b-ERR\b/i);
  if (m) return m[0];
  return rawError.slice(0, 64);
}

export interface AutoHealConnectParams {
  email: string;
  password: string;
  discovered: DiscoveredPop3Config;
  requestedProtocol: MailReceiveProtocol;
  pop3Host: string;
  pop3Port: number;
  pop3Tls: boolean;
  imapHost: string;
  imapPort: number;
  imapTls: boolean;
  maxMessages?: number;
  allowLoopbackForTest?: boolean;
  fetchPop3Impl?: typeof fetchPop3Messages;
  fetchImapImpl?: typeof fetchImapMessages;
  detectEgressImpl?: typeof detectServerEgressInfo;
}

export interface AutoHealConnectResult {
  fetched: Pop3FetchResult;
  effectiveProtocol: MailReceiveProtocol;
  effectiveHost: string;
  effectivePort: number;
  effectiveTls: boolean;
  effectivePassword: string;
  attempts: AutoHealAttempt[];
  healedBy: string | null;
}

/**
 * Smart multi-strategy mail connection with bounded self-healing (max 4 attempts):
 * 1) Primary protocol/host + exact password
 * 2) Automatic App Password whitespace/hyphen normalization (if password had spaces/hyphens)
 * 3) Automatic alternate protocol fallback (POP3 <-> IMAP) when provider supports both
 * 4) Automatic alternate domain host fallback (e.g. pop3.{domain}) if different from primary
 *
 * Strictly capped at 4 attempts to prevent corporate account lockout.
 */
export async function connectMailWithAutoHeal(
  params: AutoHealConnectParams
): Promise<AutoHealConnectResult> {
  const runPop3 = params.fetchPop3Impl ?? fetchPop3Messages;
  const runImap = params.fetchImapImpl ?? fetchImapMessages;
  const getEgress = params.detectEgressImpl ?? detectServerEgressInfo;

  const passwordCandidates = buildPasswordCandidates(params.password);
  const rawPw = passwordCandidates[0]?.value || params.password;
  const normalizedCandidates = passwordCandidates.filter((c) => c.isNormalized);
  const primaryNormalizedPw =
    normalizedCandidates.length > 0 ? normalizedCandidates[normalizedCandidates.length - 1] : null;
  const secondaryNormalizedPw =
    normalizedCandidates.length > 1 ? normalizedCandidates[0] : null;

  interface PlannedStep {
    protocol: MailReceiveProtocol;
    host: string;
    port: number;
    tls: boolean;
    password: string;
    strategyLabel: string;
    healedNote: string | null;
  }

  const plan: PlannedStep[] = [];
  const seenKeys = new Set<string>();
  const addPlan = (step: PlannedStep) => {
    if (!step.host || !step.password) return;
    const key = `${step.protocol}|${step.host}:${step.port}|${step.tls}|${step.password}`;
    if (seenKeys.has(key) || plan.length >= 5) return;
    seenKeys.add(key);
    plan.push(step);
  };

  const primaryHost = params.requestedProtocol === 'imap' ? params.imapHost : params.pop3Host;
  const primaryPort = params.requestedProtocol === 'imap' ? params.imapPort : params.pop3Port;
  const primaryTls = params.requestedProtocol === 'imap' ? params.imapTls : params.pop3Tls;

  // Step 1: Requested protocol + exact input password
  addPlan({
    protocol: params.requestedProtocol,
    host: primaryHost,
    port: primaryPort,
    tls: primaryTls,
    password: rawPw,
    strategyLabel: `${params.requestedProtocol.toUpperCase()} 기본 연결`,
    healedNote: null,
  });

  // Step 2: Most-normalized App Password (strips copy-pasted spaces / hyphens) and trimmed fallback
  if (primaryNormalizedPw) {
    addPlan({
      protocol: params.requestedProtocol,
      host: primaryHost,
      port: primaryPort,
      tls: primaryTls,
      password: primaryNormalizedPw.value,
      strategyLabel: primaryNormalizedPw.label,
      healedNote: `${primaryNormalizedPw.label}로 자동 보정하여 연결했습니다.`,
    });
  }
  if (secondaryNormalizedPw && secondaryNormalizedPw.value !== primaryNormalizedPw?.value) {
    addPlan({
      protocol: params.requestedProtocol,
      host: primaryHost,
      port: primaryPort,
      tls: primaryTls,
      password: secondaryNormalizedPw.value,
      strategyLabel: secondaryNormalizedPw.label,
      healedNote: `${secondaryNormalizedPw.label}로 자동 보정하여 연결했습니다.`,
    });
  }

  // Step 3: Alternate protocol auto-fallback (POP3 <-> IMAP) if supported
  const altProtocol: MailReceiveProtocol = params.requestedProtocol === 'imap' ? 'pop3' : 'imap';
  const altHost = altProtocol === 'imap' ? params.imapHost : params.pop3Host;
  const altPort = altProtocol === 'imap' ? params.imapPort : params.pop3Port;
  const altTls = altProtocol === 'imap' ? params.imapTls : params.pop3Tls;
  if (params.discovered.supportedProtocols.includes(altProtocol) && altHost) {
    addPlan({
      protocol: altProtocol,
      host: altHost,
      port: altPort,
      tls: altTls,
      password: primaryNormalizedPw ? primaryNormalizedPw.value : rawPw,
      strategyLabel: `${altProtocol.toUpperCase()} 프로토콜 자동 전환`,
      healedNote: `${params.requestedProtocol.toUpperCase()} 대신 ${altProtocol.toUpperCase()}(${altHost}:${altPort})로 자동 전환해 연결했습니다.`,
    });
  }

  // Step 4: Alternate custom-domain host if provider uses shared host (e.g. pop3.{domain})
  const preset = getPresetById(params.discovered.detectedProviderId);
  const domain = params.discovered.domain;
  const altTemplates =
    params.requestedProtocol === 'imap' ? preset.imapAltTemplates : preset.pop3AltTemplates;
  if (altTemplates && altTemplates.length > 0 && domain) {
    const altDomainHost = altTemplates[0].hostTemplate.replace('{domain}', domain);
    if (altDomainHost && altDomainHost !== primaryHost) {
      addPlan({
        protocol: params.requestedProtocol,
        host: altDomainHost,
        port: altTemplates[0].port,
        tls: altTemplates[0].tls,
        password: primaryNormalizedPw ? primaryNormalizedPw.value : rawPw,
        strategyLabel: `도메인 전용 서버(${altDomainHost}) 자동 전환`,
        healedNote: `도메인 전용 서버(${altDomainHost}:${altTemplates[0].port})로 자동 전환해 연결했습니다.`,
      });
    }
  }

  const attempts: AutoHealAttempt[] = [];
  let lastAuthError: MailAuthError | null = null;
  let lastRawError = '메일 서버 인증에 실패했습니다.';

  for (let i = 0; i < plan.length; i += 1) {
    const step = plan[i];
    try {
      const fetched =
        step.protocol === 'imap'
          ? await runImap({
              email: params.email,
              password: step.password,
              imapHost: step.host,
              imapPort: step.port,
              imapTls: step.tls,
              providerId: params.discovered.detectedProviderId,
              usernameFormat: params.discovered.usernameFormat,
              maxMessages: params.maxMessages,
              allowLoopbackForTest: params.allowLoopbackForTest,
            })
          : await runPop3({
              email: params.email,
              password: step.password,
              pop3Host: step.host,
              pop3Port: step.port,
              pop3Tls: step.tls,
              providerId: params.discovered.detectedProviderId,
              usernameFormat: params.discovered.usernameFormat,
              maxMessages: params.maxMessages,
              allowLoopbackForTest: params.allowLoopbackForTest,
            });

      attempts.push({
        step: i + 1,
        protocol: step.protocol,
        host: step.host,
        port: step.port,
        strategyLabel: step.strategyLabel,
        outcome: 'success',
      });

      return {
        fetched,
        effectiveProtocol: step.protocol,
        effectiveHost: step.host,
        effectivePort: step.port,
        effectiveTls: step.tls,
        effectivePassword: step.password,
        attempts,
        healedBy: step.healedNote,
      };
    } catch (err) {
      const rawMsg =
        err instanceof MailAuthError
          ? err.diagnosis.rawServerError || err.message
          : err instanceof Error
            ? err.message
            : '연결 실패';
      const safeRaw = sanitizeMailError(rawMsg, params.password);
      lastRawError = safeRaw;
      if (err instanceof MailAuthError) {
        lastAuthError = err;
      }
      attempts.push({
        step: i + 1,
        protocol: step.protocol,
        host: step.host,
        port: step.port,
        strategyLabel: step.strategyLabel,
        outcome: 'failed',
        serverCode: extractCompactServerCode(safeRaw),
      });
    }
  }

  // All automated strategies failed -> gather live server egress IP/country & detailed resolution wizard
  const serverEgress = await getEgress();
  const baseDiagnosis =
    lastAuthError?.diagnosis ||
    diagnoseMailAuthError({
      providerId: params.discovered.detectedProviderId,
      protocol: params.requestedProtocol,
      rawServerError: lastRawError,
    });

  const resolutionSteps = buildDetailedResolutionSteps({
    providerId: params.discovered.detectedProviderId,
    protocol: params.requestedProtocol,
    rawServerError: baseDiagnosis.rawServerError || lastRawError,
    serverEgress,
  });

  const enrichedDiagnosis: MailConnectionDiagnosis = {
    ...baseDiagnosis,
    serverEgress,
    autoHealAttempts: attempts,
    resolutionSteps,
  };

  throw new MailAuthError(enrichedDiagnosis.headline, enrichedDiagnosis);
}

