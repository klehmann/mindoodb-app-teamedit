import { describe, expect, it } from "vitest";
import type { MindooDBAppViewEntry } from "mindoodb-app-sdk";

import { createOpenViewDefinition, mapDocumentEntries } from "@/lib/viewOpen";

describe("viewOpen", () => {
  it("defaults File/Open to non-template documents", () => {
    expect(createOpenViewDefinition()).toMatchObject({
      id: "teamedit-open-tags-noTemplates-v1",
      filter: {
        mode: "expression",
      },
    });
  });

  it("creates separate template-filtered view definitions", () => {
    expect(createOpenViewDefinition("all").filter).toBeUndefined();
    expect(createOpenViewDefinition("onlyTemplates")).toMatchObject({
      id: "teamedit-open-tags-onlyTemplates-v1",
      filter: {
        mode: "expression",
      },
    });
  });

  it("maps document type discriminators for markdown and Word documents", () => {
    const entries: MindooDBAppViewEntry[] = [
      createDocumentEntry("markdown-doc", "Markdown", "markdown"),
      createDocumentEntry("word-doc", "Word", "word"),
    ];

    expect(mapDocumentEntries(entries)).toMatchObject([
      { id: "markdown-doc", title: "Markdown", type: "markdown" },
      { id: "word-doc", title: "Word", type: "word" },
    ]);
  });
});

function createDocumentEntry(
  docId: string,
  subject: string,
  type: "markdown" | "word",
): MindooDBAppViewEntry {
  return {
    key: docId,
    kind: "document",
    origin: "main",
    docId,
    level: 0,
    parentKey: null,
    categoryPath: [],
    categoryValue: null,
    columnValues: {
      subject,
      tags: [],
      type,
    },
    position: null,
    expanded: false,
    selected: false,
    isVisible: true,
  };
}
