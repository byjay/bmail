import test from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import type {
  SedmsIngestMessagePayload,
  SedmsOutboundDispatchRef,
  SedmsProjectCatalogItem,
} from '../types/sgwSedms';
import {
  buildNasMailStorageKey,
  buildSedmsOutboundDispatchHeaders,
  canonDrawingNo,
  createFlexibleDrawingPattern,
  matchMailDeterministically,
  normalizeAlphaNumeric,
  SedmsMailHubStore,
  signSedmsIngestPayload,
  verifySedmsIngestSignature,
} from './sedmsMailHub';

// Strictly synthetic fixtures only (20261120-MAILHUB S4: zero hardcoded vessel/company/person literals)
const SYNTHETIC_PROJECTS: SedmsProjectCatalogItem[] = [
  {
    tenantId: 'TENANT-ALPHA',
    projectId: 'PRJ-101',
    projectCode: 'H-101',
    vesselAliases: ['VSL-ALPHA-01', 'HULL-101'],
    drawingNos: ['DWG-E1001', 'DWG-E1002', 'DWG-E1003', 'DWG-SHARED-900'],
    companyDomains: ['vendor-a.example.test'],
    companyEmails: ['contact-a@partner.example.test'],
    sealedDrawings: {
      DWGE1001: 2,
    },
  },
  {
    tenantId: 'TENANT-ALPHA',
    projectId: 'PRJ-102',
    projectCode: 'H-102',
    vesselAliases: ['VSL-ALPHA-02', 'HULL-102'],
    drawingNos: ['DWG-M2001', 'DWG-M2002', 'DWG-SHARED-900'],
    companyDomains: ['vendor-b.example.test'],
    companyEmails: ['contact-b@partner.example.test'],
    sealedDrawings: {},
  },
  {
    // Completely separate tenant with identical codes/drawings to prove S3 strict tenant isolation
    tenantId: 'TENANT-BETA',
    projectId: 'PRJ-999',
    projectCode: 'H-101',
    vesselAliases: ['VSL-ALPHA-01'],
    drawingNos: ['DWG-E1001', 'DWG-BETA-001'],
    companyDomains: ['vendor-a.example.test'],
    companyEmails: ['contact-a@partner.example.test'],
    sealedDrawings: {},
  },
];

const SYNTHETIC_DISPATCHES: SedmsOutboundDispatchRef[] = [
  {
    tenantId: 'TENANT-ALPHA',
    projectId: 'PRJ-101',
    distributionId: 'DIST-101-A',
    outboundMessageId: '<sedms-dist-101-a@sedms.local>',
    drawingNoKey: 'DWGE1001',
  },
  {
    tenantId: 'TENANT-ALPHA',
    projectId: 'PRJ-102',
    distributionId: 'DIST-102-B',
    outboundMessageId: '<sedms-dist-102-b@sedms.local>',
    drawingNoKey: 'DWGM2001',
  },
  {
    tenantId: 'TENANT-BETA',
    projectId: 'PRJ-999',
    distributionId: 'DIST-999-Z',
    outboundMessageId: '<sedms-dist-999-z@sedms.local>',
    drawingNoKey: 'DWGBETA001',
  },
];

function makeMsg(
  idx: number,
  overrides: Partial<SedmsIngestMessagePayload> = {}
): SedmsIngestMessagePayload {
  return {
    tenantId: 'TENANT-ALPHA',
    accountId: 'acc-01',
    messageId: `<std-msg-${String(idx).padStart(2, '0')}@example.test>`,
    direction: 'in',
    from: 'sender@external.example.test',
    to: ['inbox@tenant-alpha.example.test'],
    subject: `Standard Test Subject ${idx}`,
    bodyText: `Standard Test Body ${idx}`,
    sentAt: `2026-11-${String((idx % 28) + 1).padStart(2, '0')}T09:00:00Z`,
    ...overrides,
  };
}

test('SEDMS drawingNoKey 정규화 및 유연 매칭 패턴 검증', () => {
  assert.equal(normalizeAlphaNumeric('dwg-e1001 / rev.2'), 'DWGE1001REV2');
  assert.equal(canonDrawingNo('DWG_E-1001'), 'DWGE1001');
  const pat = createFlexibleDrawingPattern('DWG-E1001');
  assert.ok(pat);
  assert.equal(pat.test('첨부: dwg_e 1001_final.pdf'), true);
  assert.equal(pat.test('unrelated-DWG-E1002.pdf'), false);
  assert.equal(createFlexibleDrawingPattern('AB'), null);
});

test('SEDMS 표준 테스트 메일 30건 전수 매칭 정확도 검증 (Rule 1~5 + 충돌 → review_queue)', () => {
  interface StandardCase {
    msg: SedmsIngestMessagePayload;
    expectedDecision: 'auto_link' | 'review_queue';
    expectedRule?: 1 | 2 | 3 | 4;
    expectedProjectId?: string;
  }

  const cases: StandardCase[] = [
    // 1~5: Rule 1 (In-Reply-To / References -> auto_link 100%)
    {
      msg: makeMsg(1, { inReplyTo: '<sedms-dist-101-a@sedms.local>', subject: 'Re: Drawing package' }),
      expectedDecision: 'auto_link',
      expectedRule: 1,
      expectedProjectId: 'PRJ-101',
    },
    {
      msg: makeMsg(2, { references: ['<other@test>', '<sedms-dist-102-b@sedms.local>'] }),
      expectedDecision: 'auto_link',
      expectedRule: 1,
      expectedProjectId: 'PRJ-102',
    },
    {
      msg: makeMsg(3, { inReplyTo: 'sedms-dist-101-a@sedms.local' }), // without angle brackets
      expectedDecision: 'auto_link',
      expectedRule: 1,
      expectedProjectId: 'PRJ-101',
    },
    {
      msg: makeMsg(4, {
        inReplyTo: '<sedms-dist-101-a@sedms.local>',
        subject: 'Re: Comment on H-102', // Rule 1 takes highest priority over body text
      }),
      expectedDecision: 'auto_link',
      expectedRule: 1,
      expectedProjectId: 'PRJ-101',
    },
    {
      msg: makeMsg(5, {
        tenantId: 'TENANT-BETA',
        inReplyTo: '<sedms-dist-999-z@sedms.local>',
      }),
      expectedDecision: 'auto_link',
      expectedRule: 1,
      expectedProjectId: 'PRJ-999',
    },

    // 6~13: Rule 2 (drawingNoKey in subject/body/attachment -> unique project auto_link)
    {
      msg: makeMsg(6, { subject: 'Please review DWG-E1001 rev.3' }),
      expectedDecision: 'auto_link',
      expectedRule: 2,
      expectedProjectId: 'PRJ-101',
    },
    {
      msg: makeMsg(7, { bodyText: 'Updated file dwg_e 1002 attached for approval.' }),
      expectedDecision: 'auto_link',
      expectedRule: 2,
      expectedProjectId: 'PRJ-101',
    },
    {
      msg: makeMsg(8, {
        attachments: [{ filename: 'FINAL_DWG.E.1003_signed.pdf', mime: 'application/pdf', size: 100, sha256: 'a1' }],
      }),
      expectedDecision: 'auto_link',
      expectedRule: 2,
      expectedProjectId: 'PRJ-101',
    },
    {
      msg: makeMsg(9, { subject: 'Mechanical layout DWG-M2001 submitted' }),
      expectedDecision: 'auto_link',
      expectedRule: 2,
      expectedProjectId: 'PRJ-102',
    },
    {
      msg: makeMsg(10, { bodyText: 'See DWG M 2002 comments.' }),
      expectedDecision: 'auto_link',
      expectedRule: 2,
      expectedProjectId: 'PRJ-102',
    },
    {
      msg: makeMsg(11, { subject: '[H-101] DWG-E1001 and DWG-E1002 package' }), // Rule 2 + Rule 3 agree on PRJ-101
      expectedDecision: 'auto_link',
      expectedRule: 2,
      expectedProjectId: 'PRJ-101',
    },
    {
      msg: makeMsg(12, { subject: '[VSL-ALPHA-02] DWG-M2002 updated' }), // Rule 2 + Rule 3 agree on PRJ-102
      expectedDecision: 'auto_link',
      expectedRule: 2,
      expectedProjectId: 'PRJ-102',
    },
    {
      msg: makeMsg(13, { tenantId: 'TENANT-BETA', subject: 'Submit DWG-BETA-001' }),
      expectedDecision: 'auto_link',
      expectedRule: 2,
      expectedProjectId: 'PRJ-999',
    },

    // 14~19: Rule 3 (projectCode / vesselAliases DB token -> unique project auto_link)
    {
      msg: makeMsg(14, { subject: '[H-101] Weekly status update' }),
      expectedDecision: 'auto_link',
      expectedRule: 3,
      expectedProjectId: 'PRJ-101',
    },
    {
      msg: makeMsg(15, { bodyText: 'Schedule confirmed for VSL-ALPHA-01 sea trial.' }),
      expectedDecision: 'auto_link',
      expectedRule: 3,
      expectedProjectId: 'PRJ-101',
    },
    {
      msg: makeMsg(16, { subject: 'HULL-101 inspection memo' }),
      expectedDecision: 'auto_link',
      expectedRule: 3,
      expectedProjectId: 'PRJ-101',
    },
    {
      msg: makeMsg(17, { subject: 'Action required for H-102 cabling' }),
      expectedDecision: 'auto_link',
      expectedRule: 3,
      expectedProjectId: 'PRJ-102',
    },
    {
      msg: makeMsg(18, { bodyText: 'VSL-ALPHA-02 meeting agenda' }),
      expectedDecision: 'auto_link',
      expectedRule: 3,
      expectedProjectId: 'PRJ-102',
    },
    {
      msg: makeMsg(19, { subject: 'HULL-102 kickoff notes' }),
      expectedDecision: 'auto_link',
      expectedRule: 3,
      expectedProjectId: 'PRJ-102',
    },

    // 20~23: Conflict / Ambiguity -> MUST route to review_queue (never blindly auto-link)
    {
      msg: makeMsg(20, { subject: 'H-102 comment on DWG-E1001' }), // Rule 2(PRJ-101) vs Rule 3(PRJ-102) conflict
      expectedDecision: 'review_queue',
    },
    {
      msg: makeMsg(21, { subject: 'Shared standard drawing DWG-SHARED-900' }), // Exists in PRJ-101 & PRJ-102
      expectedDecision: 'review_queue',
    },
    {
      msg: makeMsg(22, { subject: 'Combined progress H-101 and H-102' }), // Multiple Rule 3 projects
      expectedDecision: 'review_queue',
    },
    {
      msg: makeMsg(23, { subject: 'DWG-E1001 vs DWG-M2001 interface check' }), // Multiple Rule 2 projects
      expectedDecision: 'review_queue',
    },

    // 24~27: Rule 4 (Company domain / email -> ALWAYS review_queue with candidate)
    {
      msg: makeMsg(24, { from: 'sales@vendor-a.example.test', subject: 'Quotation inquiry' }),
      expectedDecision: 'review_queue',
      expectedRule: 4,
    },
    {
      msg: makeMsg(25, { from: 'contact-a@partner.example.test', subject: 'General schedule question' }),
      expectedDecision: 'review_queue',
      expectedRule: 4,
    },
    {
      msg: makeMsg(26, { from: 'eng@vendor-b.example.test', subject: 'Invoice attached' }),
      expectedDecision: 'review_queue',
      expectedRule: 4,
    },
    {
      msg: makeMsg(27, { cc: ['contact-b@partner.example.test'], subject: 'Follow up' }),
      expectedDecision: 'review_queue',
      expectedRule: 4,
    },

    // 28~30: Rule 5 (Unclassified -> honest empty review_queue)
    {
      msg: makeMsg(28, { from: 'unknown@other.example.test', subject: 'Office holiday notice' }),
      expectedDecision: 'review_queue',
    },
    {
      msg: makeMsg(29, { from: 'noreply@billing.example.test', subject: 'Monthly receipt' }),
      expectedDecision: 'review_queue',
    },
    {
      msg: makeMsg(30, {
        tenantId: 'TENANT-ALPHA',
        inReplyTo: '<sedms-dist-999-z@sedms.local>', // Cross-tenant dispatch ID MUST NOT match!
        subject: 'Cross-tenant header test',
      }),
      expectedDecision: 'review_queue',
    },
  ];

  assert.equal(cases.length, 30);

  for (let i = 0; i < cases.length; i += 1) {
    const tc = cases[i];
    const res = matchMailDeterministically({
      message: tc.msg,
      projects: SYNTHETIC_PROJECTS,
      outboundDispatches: SYNTHETIC_DISPATCHES,
    });
    assert.equal(res.decision, tc.expectedDecision, `Case #${i + 1} decision mismatch`);
    if (tc.expectedDecision === 'auto_link') {
      assert.ok(res.links.length >= 1, `Case #${i + 1} expected auto links`);
      assert.equal(res.links[0].rule, tc.expectedRule, `Case #${i + 1} rule mismatch`);
      assert.equal(res.links[0].projectId, tc.expectedProjectId, `Case #${i + 1} project mismatch`);
      assert.equal(res.links[0].tenantId, tc.msg.tenantId, `Case #${i + 1} tenant isolation mismatch`);
    } else if (tc.expectedRule === 4) {
      assert.ok(res.candidates.some((c) => c.rule === 4), `Case #${i + 1} expected Rule 4 candidate`);
    }
    // Strict S3 check: NO candidate or link may ever belong to another tenant
    for (const link of res.links) {
      assert.equal(link.tenantId, tc.msg.tenantId);
    }
    for (const cand of res.candidates) {
      const proj = SYNTHETIC_PROJECTS.find((p) => p.projectId === cand.projectId);
      assert.equal(proj?.tenantId, tc.msg.tenantId);
    }
  }
});

test('SEDMS 멱등 수집(3회 중복 수집=1건) + 첨부 sha256 중복 제거 + NAS append-only 경로 규칙', () => {
  const store = new SedmsMailHubStore();
  store.setCatalog(SYNTHETIC_PROJECTS, SYNTHETIC_DISPATCHES);

  const sha = crypto.createHash('sha256').update('same-drawing-binary').digest('hex');
  const firstPayload = makeMsg(1, {
    messageId: '<idem-001@example.test>',
    subject: '[H-101] DWG-E1001 rev.3',
    attachments: [
      { filename: 'DWG-E1001_rev3.pdf', mime: 'application/pdf', size: 4096, sha256: sha },
    ],
  });

  const r1 = store.ingestMessage(firstPayload);
  const r2 = store.ingestMessage(firstPayload);
  const r3 = store.ingestMessage(firstPayload);

  assert.equal(r1.duplicate, false);
  assert.equal(r2.duplicate, true);
  assert.equal(r3.duplicate, true);
  assert.equal(store.listProjectMail('TENANT-ALPHA', 'PRJ-101').length, 1);

  // Append-only NAS storage path format: <projectRoot>/_mail/YYYYMM/<safe-message-id>/<sha12>_<filename>
  const nasPath = r1.stored.attachments[0].storageKey;
  assert.match(nasPath, /^PRJ-101\/_mail\/202611\/_idem-001_example\.test_\/[a-f0-9]{12}_DWG-E1001_rev3\.pdf$/);

  // Second distinct mail with identical attachment sha256 reuses storageKey (dedup)
  const secondPayload = makeMsg(2, {
    messageId: '<idem-002@example.test>',
    subject: '[H-101] Re-sending DWG-E1001',
    attachments: [
      { filename: 'DWG-E1001_copy.pdf', mime: 'application/pdf', size: 4096, sha256: sha },
    ],
  });
  const rSecond = store.ingestMessage(secondPayload);
  assert.equal(rSecond.duplicate, false);
  assert.equal(rSecond.dedupedAttachmentCount, 1);
  assert.equal(rSecond.stored.attachments[0].storageKey, nasPath);

  // Cross-tenant query sees 0 messages from TENANT-ALPHA
  assert.equal(store.listProjectMail('TENANT-BETA', 'PRJ-101').length, 0);
  assert.equal(store.getMessageById('TENANT-BETA', r1.stored.id), null);

  // Direct buildNasMailStorageKey sanity
  const built = buildNasMailStorageKey({
    projectRoot: 'PRJ-101',
    sentAt: '2026-11-20T12:00:00Z',
    messageId: '<abc/def@test>',
    sha256: '0123456789abcdef0123456789abcdef',
    filename: 'file name?.pdf',
  });
  assert.equal(built, 'PRJ-101/_mail/202611/_abc_def_test_/0123456789ab_file_name_.pdf');
});

test('SEDMS M0 테넌트별 서버 간 HMAC 서명 검증 (401/403/200)', () => {
  const secrets = { 'TENANT-ALPHA': 'alpha-ingest-hmac-secret-999' };
  const rawBody = JSON.stringify({ messageId: '<hmac-1@test>' });
  const nowSec = 1_800_000_000;
  const sig = signSedmsIngestPayload({
    tenantId: 'TENANT-ALPHA',
    timestampSec: nowSec,
    rawBody,
    tenantSecret: secrets['TENANT-ALPHA'],
  });

  // Valid -> 200
  assert.equal(
    verifySedmsIngestSignature({
      tenantId: 'TENANT-ALPHA',
      timestampHeader: String(nowSec),
      signatureHeader: sig,
      rawBody,
      tenantSecrets: secrets,
      nowSec,
    }).statusCode,
    200
  );

  // Missing header -> 401
  assert.equal(
    verifySedmsIngestSignature({
      tenantId: 'TENANT-ALPHA',
      rawBody,
      tenantSecrets: secrets,
      nowSec,
    }).statusCode,
    401
  );

  // Stale timestamp -> 401
  assert.equal(
    verifySedmsIngestSignature({
      tenantId: 'TENANT-ALPHA',
      timestampHeader: String(nowSec - 600),
      signatureHeader: sig,
      rawBody,
      tenantSecrets: secrets,
      nowSec,
    }).statusCode,
    401
  );

  // Wrong tenant / bad signature -> 403
  assert.equal(
    verifySedmsIngestSignature({
      tenantId: 'TENANT-ALPHA',
      timestampHeader: String(nowSec),
      signatureHeader: '0'.repeat(64),
      rawBody,
      tenantSecrets: secrets,
      nowSec,
    }).statusCode,
    403
  );
  assert.equal(
    verifySedmsIngestSignature({
      tenantId: 'TENANT-BETA',
      timestampHeader: String(nowSec),
      signatureHeader: sig,
      rawBody,
      tenantSecrets: secrets,
      nowSec,
    }).statusCode,
    403
  );
});

test('SEDMS M4 출도 발신 헤더 → 회신 Rule 1 왕복 + 확인 대기열 확정 + M5 봉인 역행 409 차단', () => {
  const store = new SedmsMailHubStore();
  // Deep copy projects so sealedDrawings mutation stays isolated to this test
  store.setCatalog(
    SYNTHETIC_PROJECTS.map((p) => ({
      ...p,
      sealedDrawings: { ...(p.sealedDrawings || {}) },
    })),
    []
  );

  // M4: Outbound dispatch generates Message-ID & X-SEDMS-* headers
  const outbound = buildSedmsOutboundDispatchHeaders({
    tenantId: 'TENANT-ALPHA',
    projectId: 'PRJ-101',
    distributionId: 'DIST-ROUNDTRIP-01',
    domain: 'sedms.example.test',
  });
  assert.equal(outbound.headers['X-SEDMS-Tenant'], 'TENANT-ALPHA');
  assert.equal(outbound.headers['X-SEDMS-Project'], 'PRJ-101');
  assert.equal(outbound.headers['X-SEDMS-Distribution'], 'DIST-ROUNDTRIP-01');

  store.registerOutboundDispatch({
    tenantId: 'TENANT-ALPHA',
    projectId: 'PRJ-101',
    distributionId: 'DIST-ROUNDTRIP-01',
    outboundMessageId: outbound.messageId,
    drawingNoKey: 'DWGE1001',
  });

  // Inbound reply matches via Rule 1 automatically
  const reply = store.ingestMessage(
    makeMsg(99, {
      messageId: '<vendor-reply-99@example.test>',
      inReplyTo: outbound.messageId,
      subject: 'Re: Transmittal',
      bodyText: 'Attached updated drawing',
      attachments: [
        {
          filename: 'DWG-E1001_updated.pdf',
          mime: 'application/pdf',
          size: 8192,
          sha256: crypto.createHash('sha256').update('rev-file').digest('hex'),
        },
      ],
    })
  );
  assert.equal(reply.stored.links.length, 1);
  assert.equal(reply.stored.links[0].rule, 1);
  assert.equal(reply.stored.links[0].distributionId, 'DIST-ROUNDTRIP-01');

  // M5 Seal Rule (S6): DWG-E1001 is sealed at Rev.2 -> Rev.1 and Rev.2 MUST return HTTP 409
  const attId = reply.stored.attachments[0].id;
  const rev1Attempt = store.attachMailFileToDrawing({
    tenantId: 'TENANT-ALPHA',
    messageId: reply.stored.messageId,
    attachmentId: attId,
    projectId: 'PRJ-101',
    drawingNo: 'DWG-E1001',
    proposedRevision: 1,
    actorSub: 'sub-manager-01',
  });
  assert.equal(rev1Attempt.ok, false);
  assert.equal(rev1Attempt.statusCode, 409);

  const rev2Attempt = store.attachMailFileToDrawing({
    tenantId: 'TENANT-ALPHA',
    messageId: reply.stored.messageId,
    attachmentId: attId,
    projectId: 'PRJ-101',
    drawingNo: 'DWG-E1001',
    proposedRevision: 2,
    actorSub: 'sub-manager-01',
  });
  assert.equal(rev2Attempt.ok, false);
  assert.equal(rev2Attempt.statusCode, 409);

  // Forward-only Rev.3 MUST succeed (HTTP 200)
  const rev3Attempt = store.attachMailFileToDrawing({
    tenantId: 'TENANT-ALPHA',
    messageId: reply.stored.messageId,
    attachmentId: attId,
    projectId: 'PRJ-101',
    drawingNo: 'DWG-E1001',
    proposedRevision: 3,
    actorSub: 'sub-manager-01',
  });
  assert.equal(rev3Attempt.ok, true);
  assert.equal(rev3Attempt.statusCode, 200);
  assert.equal(rev3Attempt.acceptedRevision, 3);

  // Review queue manual confirmation resolves without deleting message history
  const queued = store.ingestMessage(
    makeMsg(100, {
      messageId: '<unclassified-100@example.test>',
      subject: 'General inquiry without drawing number',
      bodyText: 'Need clarification',
    })
  );
  assert.ok(queued.reviewQueueItem);
  assert.equal(store.listReviewQueue('TENANT-ALPHA', 'open').length, 1);

  const resolved = store.resolveReviewItem({
    tenantId: 'TENANT-ALPHA',
    messageId: '<unclassified-100@example.test>',
    actorSub: 'sub-manager-01',
    action: 'confirmed',
    projectId: 'PRJ-102',
  });
  assert.equal(resolved.ok, true);
  assert.equal(store.listReviewQueue('TENANT-ALPHA', 'open').length, 0);
  assert.equal(store.listProjectMail('TENANT-ALPHA', 'PRJ-102').length, 1);
});
