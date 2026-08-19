export const RECENT_DOCUMENTS_LIMIT = 10;
export const RECENT_DOCUMENTS_STORAGE_PREFIX = "mindoodb-teamedit-recent-documents";

export interface RecentDocumentEntry {
  databaseId: string;
  documentId: string;
  title: string;
  type: "markdown" | "word";
  openedAt: number;
}

export interface RecentDocumentsScope {
  appId: string;
  tenantId?: string;
  userId?: string;
}

function sanitizeKeyPart(value: string): string {
  return value.trim().replace(/[:\s]+/g, "-");
}

/**
 * localStorage key for the recents list. Always includes the Haven app id so
 * two apps on the same origin keep separate lists; tenant and user keep
 * identities from mixing documents they cannot open.
 */
export function createRecentDocumentsStorageKey(scope: RecentDocumentsScope): string {
  const appId = sanitizeKeyPart(scope.appId);
  if (!appId) {
    return "";
  }
  return [
    RECENT_DOCUMENTS_STORAGE_PREFIX,
    appId,
    sanitizeKeyPart(scope.tenantId ?? ""),
    sanitizeKeyPart(scope.userId ?? ""),
  ].join(":");
}

export function rememberRecentDocument(
  entries: readonly RecentDocumentEntry[],
  next: Omit<RecentDocumentEntry, "openedAt"> & { openedAt?: number },
  now = Date.now(),
): RecentDocumentEntry[] {
  const databaseId = next.databaseId.trim();
  const documentId = next.documentId.trim();
  if (!databaseId || !documentId) {
    return [...entries];
  }
  const remembered: RecentDocumentEntry = {
    databaseId,
    documentId,
    title: next.title.trim(),
    type: next.type === "word" ? "word" : "markdown",
    openedAt: next.openedAt ?? now,
  };
  return [
    remembered,
    ...entries.filter(
      (entry) =>
        entry.databaseId !== remembered.databaseId
        || entry.documentId !== remembered.documentId,
    ),
  ].slice(0, RECENT_DOCUMENTS_LIMIT);
}

export function removeRecentDocument(
  entries: readonly RecentDocumentEntry[],
  databaseId: string,
  documentId: string,
): RecentDocumentEntry[] {
  return entries.filter(
    (entry) => entry.databaseId !== databaseId || entry.documentId !== documentId,
  );
}

function isRecentDocumentEntry(value: unknown): value is RecentDocumentEntry {
  if (!value || typeof value !== "object") {
    return false;
  }
  const entry = value as Partial<RecentDocumentEntry>;
  return (
    typeof entry.databaseId === "string"
    && entry.databaseId.trim().length > 0
    && typeof entry.documentId === "string"
    && entry.documentId.trim().length > 0
    && typeof entry.title === "string"
    && (entry.type === "markdown" || entry.type === "word")
    && typeof entry.openedAt === "number"
    && Number.isFinite(entry.openedAt)
  );
}

export function parseRecentDocuments(raw: string | null): RecentDocumentEntry[] {
  if (!raw) {
    return [];
  }
  try {
    const parsed = JSON.parse(raw) as unknown;
    if (!Array.isArray(parsed)) {
      return [];
    }
    const entries: RecentDocumentEntry[] = [];
    for (const item of parsed) {
      if (!isRecentDocumentEntry(item)) {
        continue;
      }
      entries.push({
        databaseId: item.databaseId.trim(),
        documentId: item.documentId.trim(),
        title: item.title.trim(),
        type: item.type,
        openedAt: item.openedAt,
      });
      if (entries.length >= RECENT_DOCUMENTS_LIMIT) {
        break;
      }
    }
    return entries;
  } catch {
    return [];
  }
}

export function formatRecentDocumentMenuLabel(
  title: string,
  untitled: string,
  maxLength = 42,
): string {
  const label = title.trim() || untitled;
  if (label.length <= maxLength) {
    return label;
  }
  return `${label.slice(0, Math.max(1, maxLength - 1))}…`;
}

export function readRecentDocumentsFromStorage(key: string): RecentDocumentEntry[] {
  if (!key || typeof localStorage === "undefined") {
    return [];
  }
  try {
    return parseRecentDocuments(localStorage.getItem(key));
  } catch {
    return [];
  }
}

export function writeRecentDocumentsToStorage(
  key: string,
  entries: readonly RecentDocumentEntry[],
): void {
  if (!key || typeof localStorage === "undefined") {
    return;
  }
  try {
    localStorage.setItem(key, JSON.stringify(entries));
  } catch {
    // Quota or private-mode storage failures should not break editing.
  }
}
