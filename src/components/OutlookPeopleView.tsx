import React, { useEffect, useMemo, useState } from 'react';
import {
  Building2,
  Download,
  Mail,
  Phone,
  Search,
  UserCheck,
  Users,
} from 'lucide-react';
import type { OutlookContactRecord } from '../services/outlookLocalBridge';

interface OutlookPeopleViewProps {
  initialContacts?: OutlookContactRecord[];
  onComposeTo: (email: string, name?: string) => void;
  onFilterBySender: (email: string) => void;
}

export const OutlookPeopleView: React.FC<OutlookPeopleViewProps> = ({
  initialContacts = [],
  onComposeTo,
  onFilterBySender,
}) => {
  const [contacts, setContacts] = useState<OutlookContactRecord[]>(initialContacts);
  const [loading, setLoading] = useState(initialContacts.length === 0);
  const [search, setSearch] = useState('');
  const [companyFilter, setCompanyFilter] = useState<string>('ALL');
  const [selectedId, setSelectedId] = useState<string>('');

  useEffect(() => {
    let cancelled = false;
    fetch('/api/local/contacts')
      .then((res) => (res.ok ? res.json() : Promise.reject(new Error('연락처 로드 실패'))))
      .then((data: { contacts?: OutlookContactRecord[] }) => {
        if (cancelled) return;
        if (Array.isArray(data.contacts)) {
          setContacts(data.contacts);
          if (data.contacts.length > 0 && !selectedId) {
            setSelectedId(data.contacts[0].id);
          }
        }
      })
      .catch(() => {
        // Fallback to initialContacts silently
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, []);

  const companies = useMemo(() => {
    const counts = new Map<string, number>();
    for (const c of contacts) {
      const domain = c.email.includes('@') ? c.email.split('@')[1].toLowerCase() : '';
      const label = (c.company || domain || '기타').trim();
      counts.set(label, (counts.get(label) || 0) + 1);
    }
    return Array.from(counts.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 12);
  }, [contacts]);

  const filteredContacts = useMemo(() => {
    const q = search.trim().toLowerCase();
    return contacts.filter((c) => {
      if (companyFilter !== 'ALL') {
        const domain = c.email.includes('@') ? c.email.split('@')[1].toLowerCase() : '';
        const label = (c.company || domain || '기타').trim();
        if (label !== companyFilter) return false;
      }
      if (!q) return true;
      return `${c.fullName} ${c.email} ${c.company || ''} ${c.department || ''} ${c.jobTitle || ''}`
        .toLowerCase()
        .includes(q);
    });
  }, [contacts, companyFilter, search]);

  const selectedContact = useMemo(
    () =>
      filteredContacts.find((c) => c.id === selectedId) ||
      filteredContacts[0] ||
      null,
    [filteredContacts, selectedId]
  );

  const handleDownloadVCard = (contact: OutlookContactRecord) => {
    const lines = [
      'BEGIN:VCARD',
      'VERSION:3.0',
      `FN:${contact.fullName}`,
      `EMAIL;TYPE=INTERNET:${contact.email}`,
      contact.company ? `ORG:${contact.company}` : '',
      contact.jobTitle ? `TITLE:${contact.jobTitle}` : '',
      contact.phone ? `TEL;TYPE=WORK:${contact.phone}` : '',
      'END:VCARD',
    ]
      .filter(Boolean)
      .join('\r\n');
    const blob = new Blob([lines], { type: 'text/vcard;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${contact.fullName.replace(/[\\/:*?"<>|]/g, '_') || 'contact'}.vcf`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleDownloadCsv = () => {
    const header = ['Name', 'Email', 'Company', 'Department', 'JobTitle', 'Phone'];
    const esc = (v?: string) => `"${String(v || '').replace(/"/g, '""')}"`;
    const rows = filteredContacts.map((c) =>
      [esc(c.fullName), esc(c.email), esc(c.company), esc(c.department), esc(c.jobTitle), esc(c.phone)].join(',')
    );
    const csv = `\uFEFF${header.join(',')}\r\n${rows.join('\r\n')}`;
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `outlook-contacts-${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div className="flex-1 flex flex-col min-h-0 bg-[#FAF9F8] text-[#242424]">
      {/* Top Toolbar */}
      <div className="h-12 px-4 bg-white border-b border-[#EDEBE9] flex items-center justify-between gap-3 shrink-0">
        <div className="flex items-center gap-2">
          <Users className="w-4 h-4 text-[#0F6CBD]" />
          <h2 className="text-sm font-bold text-[#242424]">사람 · Outlook 주소록</h2>
          <span className="px-2 py-0.5 text-[11px] font-semibold bg-[#EBF3FC] text-[#0F6CBD] rounded-full tabular-nums">
            {contacts.length.toLocaleString()}명
          </span>
        </div>
        <div className="flex items-center gap-2">
          <div className="relative w-56 sm:w-72">
            <Search className="w-3.5 h-3.5 text-[#616161] absolute left-2.5 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="이름, 이메일, 회사 검색..."
              aria-label="연락처 검색"
              className="w-full pl-8 pr-3 py-1.5 text-xs bg-[#F5F5F5] border border-[#D1D1D1] rounded-[4px] focus:bg-white focus:outline-none focus:border-[#0F6CBD]"
            />
          </div>
          <button
            type="button"
            onClick={handleDownloadCsv}
            disabled={filteredContacts.length === 0}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-[#242424] bg-white border border-[#D1D1D1] hover:bg-[#F5F5F5] rounded-[4px] disabled:opacity-40"
          >
            <Download className="w-3.5 h-3.5 text-[#0F6CBD]" />
            CSV 내보내기
          </button>
        </div>
      </div>

      {/* 3-Column Outlook People Body */}
      <div className="flex-1 flex min-h-0 overflow-hidden">
        {/* Left: Company / Domain Filter */}
        <aside className="hidden md:flex w-52 shrink-0 flex-col bg-[#F5F5F5] border-r border-[#E0E0E0] p-2 overflow-y-auto space-y-0.5">
          <button
            type="button"
            onClick={() => setCompanyFilter('ALL')}
            className={`w-full flex items-center justify-between px-2.5 py-2 text-xs rounded-[4px] text-left ${
              companyFilter === 'ALL'
                ? 'bg-white text-[#0F6CBD] font-semibold shadow-2xs'
                : 'text-[#323130] hover:bg-black/[0.04]'
            }`}
          >
            <span>전체 연락처</span>
            <span className="text-[11px] tabular-nums">{contacts.length}</span>
          </button>
          <div className="pt-2 pb-1 px-2 text-[10px] font-semibold text-[#616161]">주요 회사·도메인</div>
          {companies.map(([label, cnt]) => (
            <button
              key={label}
              type="button"
              onClick={() => setCompanyFilter(label)}
              className={`w-full flex items-center justify-between gap-2 px-2.5 py-1.5 text-xs rounded-[4px] text-left ${
                companyFilter === label
                  ? 'bg-[#EBF3FC] text-[#0F6CBD] font-semibold'
                  : 'text-[#424242] hover:bg-black/[0.04]'
              }`}
            >
              <span className="truncate">{label}</span>
              <span className="text-[11px] text-[#616161] tabular-nums shrink-0">{cnt}</span>
            </button>
          ))}
        </aside>

        {/* Middle: Contacts List */}
        <div className="w-full sm:w-80 shrink-0 bg-white border-r border-[#EDEBE9] overflow-y-auto divide-y divide-[#F3F2F1]">
          {loading ? (
            <div className="p-6 text-xs text-[#616161]">Outlook 주소록을 불러오는 중입니다...</div>
          ) : filteredContacts.length === 0 ? (
            <div className="p-8 text-center text-xs text-[#616161]">
              일치하는 연락처가 없습니다.
            </div>
          ) : (
            filteredContacts.slice(0, 400).map((c) => {
              const active = selectedContact?.id === c.id;
              return (
                <button
                  key={c.id}
                  type="button"
                  onClick={() => setSelectedId(c.id)}
                  className={`w-full px-3.5 py-2.5 flex items-center gap-3 text-left transition-colors ${
                    active ? 'bg-[#EBF3FC]' : 'hover:bg-[#FAF9F8]'
                  }`}
                >
                  <span className="w-8 h-8 rounded-full bg-[#0F6CBD]/10 text-[#0F6CBD] font-bold text-xs flex items-center justify-center shrink-0">
                    {(c.fullName || c.email || '?').slice(0, 2).toUpperCase()}
                  </span>
                  <div className="min-w-0 flex-1">
                    <div className="text-xs font-semibold text-[#242424] truncate">{c.fullName}</div>
                    <div className="text-[11px] text-[#616161] truncate">{c.email}</div>
                    {c.company && (
                      <div className="text-[10px] text-[#0F6CBD] truncate">{c.company}</div>
                    )}
                  </div>
                </button>
              );
            })
          )}
        </div>

        {/* Right: Contact Detail Card */}
        <div className="hidden sm:flex flex-1 overflow-y-auto p-6 bg-[#FAF9F8]">
          {selectedContact ? (
            <div className="w-full max-w-xl bg-white rounded-[8px] border border-[#E0E0E0] p-6 space-y-5 shadow-2xs h-fit">
              <div className="flex items-start justify-between gap-4">
                <div className="flex items-center gap-4 min-w-0">
                  <div className="w-14 h-14 rounded-full bg-[#0F6CBD] text-white text-lg font-bold flex items-center justify-center shrink-0">
                    {(selectedContact.fullName || selectedContact.email).slice(0, 2).toUpperCase()}
                  </div>
                  <div className="min-w-0">
                    <h3 className="text-lg font-bold text-[#242424] truncate">
                      {selectedContact.fullName}
                    </h3>
                    <div className="text-xs text-[#616161] font-mono truncate">
                      {selectedContact.email}
                    </div>
                    {(selectedContact.company || selectedContact.jobTitle) && (
                      <div className="text-xs text-[#0F6CBD] font-medium mt-1 truncate">
                        {[selectedContact.company, selectedContact.department, selectedContact.jobTitle]
                          .filter(Boolean)
                          .join(' · ')}
                      </div>
                    )}
                  </div>
                </div>
                <span className="px-2 py-0.5 text-[10px] rounded bg-[#F5F5F5] border border-[#E0E0E0] text-[#616161] shrink-0">
                  {selectedContact.source === 'outlook-pc' ? 'Outlook PC' : 'PST 자동 추출'}
                </span>
              </div>

              <div className="flex flex-wrap gap-2 pt-1 border-t border-[#EDEBE9]">
                <button
                  type="button"
                  onClick={() => onComposeTo(selectedContact.email, selectedContact.fullName)}
                  className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-[4px]"
                >
                  <Mail className="w-3.5 h-3.5" />
                  새 메일 쓰기
                </button>
                <button
                  type="button"
                  onClick={() => onFilterBySender(selectedContact.email)}
                  className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium text-[#242424] bg-white border border-[#D1D1D1] hover:bg-[#F5F5F5] rounded-[4px]"
                >
                  <UserCheck className="w-3.5 h-3.5 text-[#0F6CBD]" />
                  관련 메일 보기
                </button>
                <button
                  type="button"
                  onClick={() => handleDownloadVCard(selectedContact)}
                  className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium text-[#242424] bg-white border border-[#D1D1D1] hover:bg-[#F5F5F5] rounded-[4px]"
                >
                  <Download className="w-3.5 h-3.5 text-[#0F6CBD]" />
                  vCard (.vcf) 저장
                </button>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs pt-2 border-t border-[#EDEBE9]">
                <div className="flex items-center gap-2 text-[#424242]">
                  <Mail className="w-3.5 h-3.5 text-[#616161] shrink-0" />
                  <span className="truncate">{selectedContact.email}</span>
                </div>
                {selectedContact.company && (
                  <div className="flex items-center gap-2 text-[#424242]">
                    <Building2 className="w-3.5 h-3.5 text-[#616161] shrink-0" />
                    <span className="truncate">{selectedContact.company}</span>
                  </div>
                )}
                {selectedContact.phone && (
                  <div className="flex items-center gap-2 text-[#424242]">
                    <Phone className="w-3.5 h-3.5 text-[#616161] shrink-0" />
                    <span className="font-mono">{selectedContact.phone}</span>
                  </div>
                )}
              </div>
            </div>
          ) : (
            <div className="m-auto text-xs text-[#616161]">
              왼쪽 목록에서 연락처를 선택해 상세 정보를 확인하세요.
            </div>
          )}
        </div>
      </div>
    </div>
  );
};
export default OutlookPeopleView;
