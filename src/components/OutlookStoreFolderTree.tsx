import React, { useMemo, useState } from 'react';
import {
  Archive,
  ChevronDown,
  ChevronRight,
  FileEdit,
  Folder,
  FolderOpen,
  Inbox,
  Search,
  Send,
  Star,
  Trash2,
} from 'lucide-react';
import type { LocalFolderRecord } from '../services/localMailDb';
import type { OutlookFavoriteLink } from '../services/localOutlookMigrator';

export interface OutlookStoreFolderTreeProps {
  folders: LocalFolderRecord[];
  outlookFavorites?: OutlookFavoriteLink[];
  favoriteFolderIds: string[];
  activeFolderFilter: string;
  activeStoreFilter?: string;
  onSelectFolder: (folderPath: string, storeName?: string) => void;
  onToggleFavoriteFolder?: (folderId: string) => void;
  isSyncingLiveOutlook?: boolean;
  onSyncLiveOutlook?: () => void;
  totalEmailsCount?: number;
  totalUnreadCount?: number;
}

interface FolderTreeNode {
  name: string;
  fullPath: string;
  storeName: string;
  itemCount: number;
  unreadCount: number;
  isDraftOrOutbox: boolean;
  children: Map<string, FolderTreeNode>;
}

const SYSTEM_FOLDER_ORDER: Record<string, number> = {
  '받은 편지함': 1,
  inbox: 1,
  '임시 보관함': 2,
  drafts: 2,
  '보낸 편지함': 3,
  'sent items': 3,
  '지운 편지함': 4,
  'deleted items': 4,
  '보낼 편지함': 5,
  outbox: 5,
  '정크 메일': 6,
  'junk email': 6,
  'rss 피드': 7,
  '검색 폴더': 8,
};

function isDraftOrOutboxFolder(name: string): boolean {
  const lower = name.trim().toLowerCase();
  return (
    lower === '임시 보관함' ||
    lower === 'drafts' ||
    lower === '보낼 편지함' ||
    lower === 'outbox'
  );
}

export const OutlookStoreFolderTree: React.FC<OutlookStoreFolderTreeProps> = ({
  folders,
  outlookFavorites = [],
  favoriteFolderIds,
  activeFolderFilter,
  activeStoreFilter,
  onSelectFolder,
  onToggleFavoriteFolder,
  isSyncingLiveOutlook = false,
  onSyncLiveOutlook,
  totalEmailsCount = 0,
  totalUnreadCount = 0,
}) => {
  const [searchQuery, setSearchQuery] = useState('');
  const [favoritesExpanded, setFavoritesExpanded] = useState(true);
  const [collapsedStores, setCollapsedStores] = useState<Record<string, boolean>>({});
  const [collapsedNodes, setCollapsedNodes] = useState<Record<string, boolean>>({});

  // Build hierarchical tree per storeName (`김봉정`, `백업`, etc.)
  const storeTrees = useMemo(() => {
    const storesMap = new Map<string, Map<string, FolderTreeNode>>();
    const q = searchQuery.trim().toLowerCase();

    for (const f of folders) {
      const storeName = (f.storeName || 'Outlook').trim();
      const rawParts = f.folderPath
        .split('/')
        .map((s) => s.trim())
        .filter(Boolean);
      if (rawParts.length === 0) continue;
      if (q && !f.folderPath.toLowerCase().includes(q) && !storeName.toLowerCase().includes(q)) {
        continue;
      }

      if (!storesMap.has(storeName)) {
        storesMap.set(storeName, new Map());
      }
      let currentLevel = storesMap.get(storeName)!;
      let builtPath = '';

      for (let i = 0; i < rawParts.length; i += 1) {
        const seg = rawParts[i];
        builtPath = builtPath ? `${builtPath} / ${seg}` : seg;
        let node = currentLevel.get(seg);
        if (!node) {
          node = {
            name: seg,
            fullPath: builtPath,
            storeName,
            itemCount: 0,
            unreadCount: 0,
            isDraftOrOutbox: isDraftOrOutboxFolder(seg),
            children: new Map(),
          };
          currentLevel.set(seg, node);
        }
        if (i === rawParts.length - 1) {
          node.itemCount = Math.max(node.itemCount, Number(f.itemCount || 0));
          node.unreadCount = Math.max(node.unreadCount, Number(f.unreadCount || 0));
        }
        currentLevel = node.children;
      }
    }

    // Order stores: primary user store first (e.g. 김봉정), backup stores next
    return Array.from(storesMap.entries()).sort((a, b) => {
      const aIsBackup = /백업|backup|archive/i.test(a[0]) ? 1 : 0;
      const bIsBackup = /백업|backup|archive/i.test(b[0]) ? 1 : 0;
      if (aIsBackup !== bIsBackup) return aIsBackup - bIsBackup;
      return a[0].localeCompare(b[0], 'ko');
    });
  }, [folders, searchQuery]);

  // Resolve Favorites list combining Outlook New2026.xml favorites + user pinned folders
  const resolvedFavorites = useMemo(() => {
    const items: Array<{
      key: string;
      label: string;
      folderPath: string;
      storeName?: string;
      unreadCount: number;
      itemCount: number;
      isDraft: boolean;
    }> = [];
    const seen = new Set<string>();

    const findMatchingFolder = (nameHint: string): LocalFolderRecord | undefined => {
      const lower = nameHint.trim().toLowerCase();
      return (
        folders.find((f) => f.folderPath.toLowerCase() === lower) ||
        folders.find((f) => {
          const leaf = f.folderPath.split('/').pop()?.trim().toLowerCase();
          return leaf === lower;
        })
      );
    };

    for (const xmlFav of outlookFavorites) {
      const matched = findMatchingFolder(xmlFav.name);
      const pathKey = matched ? `${matched.storeName || ''}::${matched.folderPath}` : xmlFav.name;
      if (seen.has(pathKey.toLowerCase())) continue;
      seen.add(pathKey.toLowerCase());
      items.push({
        key: pathKey,
        label: xmlFav.name,
        folderPath: matched ? matched.folderPath : xmlFav.name,
        storeName: matched?.storeName,
        unreadCount: Number(matched?.unreadCount || 0),
        itemCount: Number(matched?.itemCount || 0),
        isDraft: isDraftOrOutboxFolder(xmlFav.name),
      });
    }

    // Ensure core default favorites if XML had none
    if (items.length === 0) {
      for (const defaultName of ['받은 편지함', '보낸 편지함', '임시 보관함']) {
        const matched = findMatchingFolder(defaultName);
        items.push({
          key: defaultName,
          label: defaultName,
          folderPath: matched ? matched.folderPath : defaultName,
          storeName: matched?.storeName,
          unreadCount: Number(matched?.unreadCount || 0),
          itemCount: Number(matched?.itemCount || 0),
          isDraft: isDraftOrOutboxFolder(defaultName),
        });
      }
    }

    for (const favId of favoriteFolderIds) {
      if (favId.startsWith('project:')) continue;
      const matched = findMatchingFolder(favId);
      if (!matched) continue;
      const pathKey = `${matched.storeName || ''}::${matched.folderPath}`;
      if (seen.has(pathKey.toLowerCase())) continue;
      seen.add(pathKey.toLowerCase());
      const leaf = matched.folderPath.split('/').pop()?.trim() || matched.folderPath;
      items.push({
        key: pathKey,
        label: leaf,
        folderPath: matched.folderPath,
        storeName: matched.storeName,
        unreadCount: Number(matched.unreadCount || 0),
        itemCount: Number(matched.itemCount || 0),
        isDraft: isDraftOrOutboxFolder(leaf),
      });
    }

    return items;
  }, [outlookFavorites, favoriteFolderIds, folders]);

  const sortNodes = (nodesMap: Map<string, FolderTreeNode>): FolderTreeNode[] =>
    Array.from(nodesMap.values()).sort((a, b) => {
      const oa = SYSTEM_FOLDER_ORDER[a.name.toLowerCase()] ?? 99;
      const ob = SYSTEM_FOLDER_ORDER[b.name.toLowerCase()] ?? 99;
      if (oa !== ob) return oa - ob;
      return a.name.localeCompare(b.name, 'ko');
    });

  const isSystemFolder = (name: string): boolean => {
    const lower = name.trim().toLowerCase();
    return Boolean(SYSTEM_FOLDER_ORDER[lower]);
  };

  const renderFolderIcon = (name: string, isOpen: boolean) => {
    const lower = name.toLowerCase();
    if (lower === '받은 편지함' || lower === 'inbox') {
      return <Inbox className="w-3.5 h-3.5 text-[#0F6CBD] shrink-0" />;
    }
    if (lower === '보낸 편지함' || lower === 'sent items') {
      return <Send className="w-3.5 h-3.5 text-slate-600 shrink-0" />;
    }
    if (lower === '임시 보관함' || lower === 'drafts' || lower === '보낼 편지함') {
      return <FileEdit className="w-3.5 h-3.5 text-slate-600 shrink-0" />;
    }
    if (lower === '지운 편지함' || lower === 'deleted items') {
      return <Trash2 className="w-3.5 h-3.5 text-slate-600 shrink-0" />;
    }
    if (isOpen) {
      return <FolderOpen className="w-3.5 h-3.5 text-[#0F6CBD] shrink-0" />;
    }
    return <Folder className="w-3.5 h-3.5 text-slate-600 shrink-0" />;
  };

  const renderNode = (node: FolderTreeNode, depth: number): React.ReactNode => {
    const nodeKey = `${node.storeName}::${node.fullPath}`;
    const hasChildren = node.children.size > 0;
    const isDefaultCollapsed =
      depth >= 1 &&
      node.name !== '받은 편지함' &&
      node.name !== '22K FD' &&
      !searchQuery.trim();
    const isCollapsed =
      collapsedNodes[nodeKey] !== undefined ? collapsedNodes[nodeKey] : isDefaultCollapsed;
    const isActive =
      activeFolderFilter === node.fullPath &&
      (!activeStoreFilter || activeStoreFilter === node.storeName);
    const showIcon = depth === 0 || isSystemFolder(node.name);

    return (
      <div key={nodeKey}>
        <div
          className={`group flex items-center justify-between pr-2 py-[3px] rounded-[3px] cursor-pointer text-[12px] transition-colors ${
            isActive
              ? 'bg-[#CFE4FA] text-[#0F6CBD] font-bold'
              : 'text-[#242424] hover:bg-black/[0.05]'
          }`}
          style={{ paddingLeft: `${8 + depth * 12}px` }}
          onClick={() => onSelectFolder(node.fullPath, node.storeName)}
        >
          <div className="flex items-center gap-1 min-w-0 flex-1">
            {hasChildren ? (
              <button
                type="button"
                onClick={(e) => {
                  e.stopPropagation();
                  setCollapsedNodes((prev) => ({
                    ...prev,
                    [nodeKey]: !isCollapsed,
                  }));
                }}
                className="w-4 h-4 flex items-center justify-center text-slate-500 hover:text-slate-900 shrink-0"
                aria-label={isCollapsed ? `${node.name} 펼치기` : `${node.name} 접기`}
              >
                {isCollapsed ? (
                  <ChevronRight className="w-3 h-3" />
                ) : (
                  <ChevronDown className="w-3 h-3" />
                )}
              </button>
            ) : (
              <span className="w-4 shrink-0" />
            )}
            {showIcon && renderFolderIcon(node.name, !isCollapsed && hasChildren)}
            <span
              className={`truncate ${
                node.unreadCount > 0 || isActive ? 'font-bold text-[#242424]' : 'font-normal'
              }`}
            >
              {node.name}
            </span>
          </div>

          <div className="flex items-center gap-1 shrink-0 ml-1">
            {node.isDraftOrOutbox && node.itemCount > 0 ? (
              <span className="text-[11px] font-semibold text-[#0F6CBD]">
                [{node.itemCount}]
              </span>
            ) : node.unreadCount > 0 ? (
              <span className="text-[11px] font-bold text-[#0F6CBD]">
                {node.unreadCount}
              </span>
            ) : node.itemCount > 0 ? (
              <span className="text-[10px] text-slate-400 opacity-0 group-hover:opacity-100">
                {node.itemCount}
              </span>
            ) : null}

            {onToggleFavoriteFolder && (
              <button
                type="button"
                onClick={(e) => {
                  e.stopPropagation();
                  onToggleFavoriteFolder(node.fullPath);
                }}
                title="즐겨찾기 고정/해제"
                className="opacity-0 group-hover:opacity-100 text-slate-400 hover:text-amber-500"
              >
                <Star className="w-3 h-3" />
              </button>
            )}
          </div>
        </div>

        {hasChildren && !isCollapsed && (
          <div>{sortNodes(node.children).map((child) => renderNode(child, depth + 1))}</div>
        )}
      </div>
    );
  };

  const syncTitleHint = onSyncLiveOutlook
    ? isSyncingLiveOutlook
      ? 'Outlook 증분 동기화 중...'
      : '상단 리본에서 신규 메일 증분 동기화를 실행할 수 있습니다'
    : undefined;

  return (
    <div className="flex flex-col h-full select-none text-[#242424]" title={syncTitleHint}>
      {/* Compact Folder Quick Filter (shown only when tree has > 12 folders) */}
      {folders.length > 12 && (
        <div className="px-2 py-1.5 border-b border-[#E1DFDD] shrink-0">
          <div className="relative flex items-center">
            <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2 pointer-events-none" />
            <input
              type="search"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="폴더 검색 (예: 22K, 선보)"
              className="w-full pl-7 pr-2 py-1 text-[11px] bg-white border border-[#E1DFDD] rounded-[3px] focus:outline-none focus:border-[#0F6CBD]"
            />
          </div>
        </div>
      )}

      {/* Scrollable Classic Outlook Tree */}
      <div className="flex-1 overflow-y-auto py-1.5 px-1 space-y-2">
        {/* 1. Favorites Section (New2026.xml) */}
        <div>
          <button
            type="button"
            onClick={() => setFavoritesExpanded((v) => !v)}
            className="w-full flex items-center gap-1 px-1.5 py-1 text-[12px] font-bold text-[#242424] hover:bg-black/[0.04] rounded"
          >
            {favoritesExpanded ? (
              <ChevronDown className="w-3.5 h-3.5 text-slate-600" />
            ) : (
              <ChevronRight className="w-3.5 h-3.5 text-slate-600" />
            )}
            <span>즐겨찾기</span>
          </button>

          {favoritesExpanded && (
            <div className="mt-0.5 space-y-0.5">
              {resolvedFavorites.map((fav) => {
                const isActive = activeFolderFilter === fav.folderPath;
                return (
                  <div
                    key={fav.key}
                    onClick={() => onSelectFolder(fav.folderPath, fav.storeName)}
                    className={`flex items-center justify-between pl-6 pr-2 py-[3px] rounded-[3px] cursor-pointer text-[12px] ${
                      isActive
                        ? 'bg-[#CFE4FA] text-[#0F6CBD] font-bold'
                        : 'text-[#242424] hover:bg-black/[0.05]'
                    }`}
                  >
                    <div className="flex items-center gap-1.5 min-w-0">
                      {renderFolderIcon(fav.label, false)}
                      <span className={`truncate ${fav.unreadCount > 0 ? 'font-bold' : ''}`}>
                        {fav.label}
                      </span>
                    </div>
                    {fav.isDraft && fav.itemCount > 0 ? (
                      <span className="text-[11px] font-semibold text-[#0F6CBD]">
                        [{fav.itemCount}]
                      </span>
                    ) : fav.unreadCount > 0 ? (
                      <span className="text-[11px] font-bold text-[#0F6CBD]">
                        {fav.unreadCount}
                      </span>
                    ) : null}
                  </div>
                );
              })}
            </div>
          )}
        </div>

        {/* 2. Per-Store Collapsible Folder Trees (김봉정, 백업, etc.) */}
        {storeTrees.map(([storeName, rootNodes]) => {
          const isStoreCollapsed = Boolean(collapsedStores[storeName]);
          return (
            <div key={storeName} className="pt-1 border-t border-black/[0.06]">
              <button
                type="button"
                onClick={() =>
                  setCollapsedStores((prev) => ({
                    ...prev,
                    [storeName]: !isStoreCollapsed,
                  }))
                }
                className="w-full flex items-center justify-between px-1.5 py-1 text-[12px] font-bold text-[#242424] hover:bg-black/[0.04] rounded"
              >
                <div className="flex items-center gap-1 min-w-0">
                  {isStoreCollapsed ? (
                    <ChevronRight className="w-3.5 h-3.5 text-slate-600 shrink-0" />
                  ) : (
                    <ChevronDown className="w-3.5 h-3.5 text-slate-600 shrink-0" />
                  )}
                  <Archive className="w-3.5 h-3.5 text-[#0F6CBD] shrink-0" />
                  <span className="truncate">{storeName}</span>
                </div>
              </button>

              {!isStoreCollapsed && (
                <div className="mt-0.5 space-y-0.5">
                  {sortNodes(rootNodes).map((node) => renderNode(node, 1))}
                </div>
              )}
            </div>
          );
        })}
      </div>

      {/* Classic Outlook Bottom Status Bar */}
      <div className="px-2.5 py-1 bg-[#F0F0F0] border-t border-[#E0E0E0] text-[11px] text-[#424242] flex items-center justify-between shrink-0">
        <span>항목: {totalEmailsCount.toLocaleString()}</span>
        <span className="font-semibold text-[#0F6CBD]">읽지 않음: {totalUnreadCount.toLocaleString()}</span>
      </div>
    </div>
  );
};
