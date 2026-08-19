import { describe, expect, it } from "vitest";
import type { MindooDBAppViewEntry } from "mindoodb-app-sdk";

import { createOpenViewDefinition, mapDocumentEntries, buildOpenCategoryTree, usableExistingTagNodes } from "@/lib/viewOpen";

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

  it("keeps hierarchical tag paths on category tree nodes", () => {
    const tree = buildOpenCategoryTree(
      [
        {
          key: "work",
          kind: "category",
          origin: "main",
          level: 0,
          parentKey: null,
          categoryPath: ["Work"],
          categoryValue: "Work",
          columnValues: {},
          position: null,
          expanded: false,
          selected: false,
          isVisible: true,
          descendantDocumentCount: 2,
          docId: null,
        },
        {
          key: "work-planning",
          kind: "category",
          origin: "main",
          level: 1,
          parentKey: "work",
          categoryPath: ["Work", "Planning"],
          categoryValue: "Planning",
          columnValues: {},
          position: null,
          expanded: false,
          selected: false,
          isVisible: true,
          descendantDocumentCount: 1,
          docId: null,
        },
      ],
      2,
    );

    expect(tree.roots[0]?.children).toMatchObject([
      {
        key: "work",
        tag: "Work",
        children: [{ key: "work-planning", tag: "Work\\Planning" }],
      },
    ]);
  });

  it("does not treat null category values as the tag null", () => {
    const tree = buildOpenCategoryTree(
      [
        {
          key: "untagged",
          kind: "category",
          origin: "main",
          level: 0,
          parentKey: null,
          categoryPath: [null],
          categoryValue: null,
          columnValues: {},
          position: null,
          expanded: false,
          selected: false,
          isVisible: true,
          descendantDocumentCount: 1,
          docId: null,
        },
        {
          key: "work",
          kind: "category",
          origin: "main",
          level: 0,
          parentKey: null,
          categoryPath: ["Work"],
          categoryValue: "Work",
          columnValues: {},
          position: null,
          expanded: false,
          selected: false,
          isVisible: true,
          descendantDocumentCount: 1,
          docId: null,
        },
      ],
      2,
    );

    expect(tree.roots[0]?.children).toMatchObject([
      { key: "untagged", label: "Untagged", tag: undefined },
      { key: "work", tag: "Work" },
    ]);
    expect(usableExistingTagNodes(tree.roots[0]?.children ?? [])).toMatchObject([
      { key: "work", tag: "Work" },
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
