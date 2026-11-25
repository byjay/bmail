import test from 'node:test';
import assert from 'node:assert/strict';
import {
  buildUnifiedAddressBook,
  buildWindowsOutlookInspectorScript,
  exportContactsToVcf,
  getEmployeeStorageKey,
  parseWindowsOutlookProfileJson,
} from './outlookLocalBridge';
import type { MailItem } from '../types/mail';

test('getEmployeeStorageKey isolates storage per employee email cleanly', () => {
  assert.equal(getEmployeeStorageKey('maillog_projects', ''), 'maillog_projects');
  assert.equal(
    getEmployeeStorageKey('maillog_projects', 'Kim@Shipyard.CO.KR'),
    'maillog_projects__u_kim@shipyard.co.kr'
  );
  assert.notEqual(
    getEmployeeStorageKey('maillog_projects', 'emp01@corp.kr'),
    getEmployeeStorageKey('maillog_projects', 'emp02@corp.kr')
  );
});

test('buildWindowsOutlookInspectorScript outputs safe PowerShell script without password collection', () => {
  const ps1 = buildWindowsOutlookInspectorScript();
  assert.ok(ps1.includes('Outlook.Application'));
  assert.ok(ps1.includes('GetDefaultFolder(10)'));
  assert.ok(ps1.includes('bmail-outlook-pc-profile.json'));
  assert.ok(!ps1.toLowerCase().includes('pop3 password'));
});

test('parseWindowsOutlookProfileJson validates accounts/contacts and ignores extra secret fields', () => {
  const parsed = parseWindowsOutlookProfileJson({
    machineName: 'ENG-PC-07',
    outlookVersion: '16.0.17928',
    defaultProfileName: 'Outlook',
    accounts: [
      {
        displayName: '홍길동',
        smtpAddress: 'Hong@Company.co.kr',
        accountType: 'POP3',
        pop3Host: 'pop3.mailplug.co.kr',
        pop3Port: 995,
        deliveryStorePath: 'C:\\Users\\Hong\\Documents\\Outlook Files\\hong.pst',
        password: 'should-never-be-kept',
      },
      { displayName: 'Invalid', smtpAddress: 'not-an-email' },
    ],
    pstPaths: ['C:\\Users\\Hong\\Documents\\Outlook Files\\hong.pst'],
    contacts: [
      {
        fullName: '박발주',
        email: 'Buyer@Partner.co.kr',
        company: '파트너조선',
        department: '구매팀',
        phone: '02-1234-5678',
      },
    ],
    folders: ['개인 폴더 / 받은 편지함', '개인 폴더 / PRJ-H101'],
    signatureNames: ['기본서명'],
  });

  assert.equal(parsed.accounts.length, 1);
  assert.equal(parsed.accounts[0].smtpAddress, 'hong@company.co.kr');
  assert.equal('password' in parsed.accounts[0], false);
  assert.equal(parsed.contacts.length, 1);
  assert.equal(parsed.contacts[0].email, 'buyer@partner.co.kr');
  assert.equal(parsed.pstPaths.length, 1);
});

test('buildUnifiedAddressBook merges PC contacts with PST/POP3 emails and exports valid vCard (.vcf)', () => {
  const sampleMails: MailItem[] = [
    {
      id: 'm1',
      folder: 'inbox',
      subject: '도면 확인',
      fromName: '김설계',
      fromEmail: 'design@yard.co.kr',
      to: ['이영업 <sales@company.co.kr>'],
      date: '2026-11-20',
      timeLabel: '09:00',
      timestamp: '2026-11-20T09:00:00Z',
      snippet: '확인 요청',
      body: '본문',
      isRead: false,
      source: 'pst-import',
    },
  ];

  const contacts = buildUnifiedAddressBook(sampleMails, [
    {
      id: 'c1',
      fullName: '박발주',
      email: 'buyer@partner.co.kr',
      company: '파트너조선',
      source: 'outlook-pc',
    },
  ]);

  assert.equal(contacts.length, 3);
  const vcf = exportContactsToVcf(contacts);
  assert.ok(vcf.includes('BEGIN:VCARD'));
  assert.ok(vcf.includes('EMAIL;TYPE=INTERNET:design@yard.co.kr'));
  assert.ok(vcf.includes('ORG:파트너조선'));
});
