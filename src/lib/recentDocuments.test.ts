import { afterEach, describe, expect, it } from "vitest";

import {
  RECENT_DOCUMENTS_LIMIT,
  createRecentDocumentsStorageKey,
  formatRecentDocumentMenuLabel,
  parseRecentDocuments,
  readRecentDocumentsFromStorage,
  rememberRecentDocument,
  removeRecentDocument,
  writeRecentDocumentsToStorage,
} from "./recentDocuments";

describe("recentDocuments", () => {
  afterEach(() => {
    localStorage.clear();
  });

  it("scopes the storage key to the Haven app id", () => {
    expect(
      createRecentDocumentsStorageKey({
        appId: "teamedit",
        tenantId: "acme",
        userId: "user-1",
      }),
    ).toBe("mindoodb-teamedit-recent-documents:teamedit:acme:user-1");
    expect(createRecentDocumentsStorageKey({ appId: "  " })).toBe("");
  });

  it("moves a reopened document to the front and caps the list at 10", () => {
    const seeded = Array.from({ length: RECENT_DOCUMENTS_LIMIT }, (_, index) => ({
      databaseId: "db",
      documentId: `doc-${index}`,
      title: `Doc ${index}`,
      type: "markdown" as const,
      openedAt: index,
    }));

    const withExisting = rememberRecentDocument(seeded, {
      databaseId: "db",
      documentId: "doc-3",
      title: "Renamed",
      type: "word",
    }, 100);
    expect(withExisting[0]).toMatchObject({
      documentId: "doc-3",
      title: "Renamed",
      type: "word",
      openedAt: 100,
    });
    expect(withExisting).toHaveLength(RECENT_DOCUMENTS_LIMIT);

    const overflow = rememberRecentDocument(withExisting, {
      databaseId: "db",
      documentId: "doc-new",
      title: "Newest",
      type: "markdown",
    }, 200);
    expect(overflow[0]?.documentId).toBe("doc-new");
    expect(overflow).toHaveLength(RECENT_DOCUMENTS_LIMIT);
    expect(overflow.some((entry) => entry.documentId === "doc-9")).toBe(false);
  });

  it("drops invalid stored JSON and round-trips valid entries", () => {
    expect(parseRecentDocuments("not-json")).toEqual([]);
    expect(parseRecentDocuments('{"documentId":"x"}')).toEqual([]);

    const entries = rememberRecentDocument([], {
      databaseId: "teamedit",
      documentId: "doc-1",
      title: " Notes ",
      type: "markdown",
    }, 1);
    writeRecentDocumentsToStorage("test-key", entries);
    expect(readRecentDocumentsFromStorage("test-key")).toEqual([
      {
        databaseId: "teamedit",
        documentId: "doc-1",
        title: "Notes",
        type: "markdown",
        openedAt: 1,
      },
    ]);
    expect(removeRecentDocument(entries, "teamedit", "doc-1")).toEqual([]);
  });

  it("truncates long menu labels", () => {
    expect(formatRecentDocumentMenuLabel("  ", "Untitled")).toBe("Untitled");
    expect(formatRecentDocumentMenuLabel("Short note", "Untitled")).toBe("Short note");
    expect(formatRecentDocumentMenuLabel("A".repeat(50), "Untitled", 10)).toBe("AAAAAAAAA…");
  });
});
