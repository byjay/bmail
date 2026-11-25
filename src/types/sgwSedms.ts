export type SgwMailboxFolder = 'inbox' | 'sent' | 'drafts' | 'trash';

export interface SgwUserIdentity {
  sub: string;
  name: string;
  department: string;
  email: string;
  role: 'user' | 'admin' | string;
  authSource: 'sgw-sso' | 'sgw-me' | string;
}

export interface SgwEmployee {
  sub: string;
  name: string;
  department?: string;
  email?: string;
  title?: string;
}

export interface SgwMailRecipient {
  sub: string;
  name: string;
  type: 'to' | 'cc' | 'bcc';
  isRead: boolean;
  readAt?: string;
  deletedAt?: string | null;
  purged?: boolean;
}

export interface SgwMailAttachment {
  id: string;
  filename: string;
  mimeType: string;
  sizeBytes: number;
  sha256: string;
  storageKey: string;
  scanStatus?: 'clean' | 'pending' | 'blocked';
  downloadToken?: string;
}

export interface SgwMailViewItem {
  id: string;
  threadId: string;
  senderSub: string;
  senderName: string;
  recipients: SgwMailRecipient[];
  subject: string;
  body: string;
  folder: SgwMailboxFolder;
  isDraft: boolean;
  isRead: boolean;
  sentAt: string;
  createdAt: string;
  updatedAt: string;
  attachments: SgwMailAttachment[];
  legacyInternalMailId?: string;
}

export interface SgwMailLimits {
  maxSubjectLength: number;
  maxBodyLength: number;
  maxAttachmentBytes: number;
  allowedAttachmentExtensions: string[];
}

export interface SgwUserAiPreferences {
  sub: string;
  aiEnabled: boolean;
  features: {
    summarize: boolean;
    replyDraft: boolean;
    classifyAndRules: boolean;
  };
  tone: 'formal' | 'concise' | 'friendly';
  length: 'short' | 'medium' | 'detailed';
  language: 'ko' | 'en' | 'auto';
  signature: string;
  allowTraining: boolean;
  updatedAt: string;
}

export interface DomainGlossaryEntry {
  id: string;
  ownerSub: string;
  term: string;
  definition: string;
  category: string;
  createdAt: string;
}

export interface SgwMailRule {
  id: string;
  ownerSub: string;
  name: string;
  matchKeyword: string;
  targetCategory: string;
  enabled: boolean;
  createdAt: string;
}

export interface SgwAuditLog {
  id: string;
  actorSub: string;
  action: string;
  targetId: string;
  createdAt: string;
  detail?: string;
}

export interface LegacyInternalMailRow {
  id: string | number;
  senderId: string;
  senderName?: string;
  recipientId: string;
  recipientName?: string;
  subject: string;
  body: string;
  isRead: number | boolean;
  createdAt: string;
}

export interface MigrationDryRunReport {
  totalSourceRows: number;
  validRows: number;
  skippedDuplicateIds: string[];
  selfSendRowIds: string[];
  missingFieldErrors: Array<{ rowId: string; reason: string }>;
  encodingWarnings: Array<{ rowId: string; reason: string }>;
}

export interface MigrationSampleCheck {
  legacyId: string;
  newMailId: string;
  senderSub: string;
  recipientSub: string;
  isReadMatch: boolean;
}

export interface MigrationExecutionReport {
  sourceTotal: number;
  migratedCount: number;
  skippedCount: number;
  verifiedCountMatch: boolean;
  sampleVerificationById: MigrationSampleCheck[];
  executedAt: string;
}

export interface SedmsProjectCatalogItem {
  tenantId: string;
  projectId: string;
  projectCode: string;
  vesselAliases: string[];
  drawingNos: string[];
  companyDomains: string[];
  companyEmails: string[];
  sealedDrawings?: Record<string, number>;
}

export interface SedmsOutboundDispatchRef {
  tenantId: string;
  projectId: string;
  distributionId: string;
  outboundMessageId: string;
  drawingNoKey?: string;
}

export interface SedmsIngestAttachmentPayload {
  filename: string;
  mime: string;
  size: number;
  sha256: string;
}

export interface SedmsIngestMessagePayload {
  tenantId: string;
  accountId: string;
  messageId: string;
  direction: 'in' | 'out';
  from: string;
  to: string[];
  cc?: string[];
  subject: string;
  bodyText: string;
  sentAt: string;
  inReplyTo?: string;
  references?: string[];
  attachments?: SedmsIngestAttachmentPayload[];
}

export interface SedmsMatchCandidate {
  projectId: string;
  rule: 1 | 2 | 3 | 4;
  score: number;
  evidence: string;
  drawingNoKey?: string;
  distributionId?: string;
}

export interface SedmsMailLink {
  id: string;
  tenantId: string;
  messageId: string;
  projectId: string;
  rule: 1 | 2 | 3 | 4 | 5;
  drawingNoKey?: string | null;
  distributionId?: string;
  confidence: number;
  createdAt: string;
  confirmedBy?: string;
}

export interface SedmsStoredAttachment {
  id: string;
  filename: string;
  mime: string;
  size: number;
  sha256: string;
  storageKey: string;
}

export interface SedmsStoredMessage {
  id: string;
  tenantId: string;
  accountId: string;
  messageId: string;
  direction: 'in' | 'out';
  from: string;
  to: string[];
  cc: string[];
  subject: string;
  bodyText: string;
  sentAt: string;
  inReplyTo?: string;
  references: string[];
  attachments: SedmsStoredAttachment[];
  links: SedmsMailLink[];
  ingestedAt: string;
}

export interface SedmsReviewQueueItem {
  id: string;
  tenantId: string;
  messageId: string;
  subject: string;
  from: string;
  sentAt: string;
  reason: string;
  candidates: SedmsMatchCandidate[];
  state: 'open' | 'confirmed' | 'changed' | 'rejected';
  resolvedBy?: string;
  resolvedAt?: string;
}
