import {
  createViewLanguage,
  type MindooDBAppViewDefinition,
  type MindooDBAppViewEntry,
  type MindooDBAppViewNavigator,
} from "mindoodb-app-sdk";

import { normalizeTags } from "@/lib/documentTags";

export const ALL_DOCUMENTS_NODE_KEY = "all";
export type OpenDocumentTemplateFilter = "all" | "noTemplates" | "onlyTemplates";

export interface OpenCategoryNode {
  key: string;
  label: string;
  count: number;
  children: OpenCategoryNode[];
}

export interface OpenDocumentRow {
  key: string;
  id: string;
  title: string;
  type: "markdown" | "word";
  tags: string[];
  detail: string;
}

export function createOpenViewDefinition(
  templateFilter: OpenDocumentTemplateFilter = "noTemplates",
): MindooDBAppViewDefinition {
  const v = createViewLanguage();
  const definition: MindooDBAppViewDefinition = {
    id: `teamedit-open-tags-${templateFilter}-v1`,
    title: "TeamEdit documents by tag",
    defaultExpand: "expanded",
    columns: [
      {
        name: "tags",
        title: "Tags",
        role: "category",
        expression: v.field("tags"),
        sorting: "ascending",
      },
      {
        name: "subject",
        title: "Title",
        role: "display",
        expression: v.field("subject"),
        sorting: "ascending",
      },
      {
        name: "type",
        title: "Type",
        role: "display",
        expression: v.field("type"),
        sorting: "ascending",
      },
    ],
  };
  if (templateFilter === "noTemplates") {
    definition.filter = {
      mode: "expression",
      expression: v.neq(v.field("istemplate"), true),
    };
  } else if (templateFilter === "onlyTemplates") {
    definition.filter = {
      mode: "expression",
      expression: v.eq(v.field("istemplate"), true),
    };
  }
  return definition;
}

export async function collectNavigatorEntries(navigator: MindooDBAppViewNavigator) {
  const entries: MindooDBAppViewEntry[] = [];
  let startPosition: string | null = null;
  do {
    const page = await navigator.entriesForward({ limit: 1000, startPosition });
    entries.push(...page.entries);
    startPosition = page.nextPosition;
  } while (startPosition);
  return entries;
}

export function buildOpenCategoryTree(categoryEntries: MindooDBAppViewEntry[], documentCount: number) {
  const root: OpenCategoryNode = {
    key: ALL_DOCUMENTS_NODE_KEY,
    label: "All documents",
    count: documentCount,
    children: [],
  };
  const nodesByKey = new Map<string, OpenCategoryNode>([[root.key, root]]);
  for (const entry of categoryEntries) {
    const node: OpenCategoryNode = {
      key: entry.key,
      label: readCategoryLabel(entry),
      count: entry.descendantDocumentCount ?? 0,
      children: [],
    };
    nodesByKey.set(node.key, node);
  }
  for (const entry of categoryEntries) {
    const node = nodesByKey.get(entry.key);
    if (!node) {
      continue;
    }
    const parent = entry.parentKey ? nodesByKey.get(entry.parentKey) : root;
    parent?.children.push(node);
  }
  return { roots: [root], nodesByKey };
}

export function dedupeDocumentEntries(entries: MindooDBAppViewEntry[]) {
  const rowsById = new Map<string, OpenDocumentRow>();
  for (const entry of entries) {
    const row = mapDocumentEntry(entry);
    if (row && !rowsById.has(row.id)) {
      rowsById.set(row.id, row);
    }
  }
  return Array.from(rowsById.values());
}

export function mapDocumentEntries(entries: MindooDBAppViewEntry[]) {
  return entries
    .map(mapDocumentEntry)
    .filter((row): row is OpenDocumentRow => row !== null);
}

function mapDocumentEntry(entry: MindooDBAppViewEntry) {
  if (entry.kind !== "document" || !entry.docId) {
    return null;
  }
  const tags = normalizeTags(entry.columnValues.tags);
  return {
    key: entry.key,
    id: entry.docId,
    title: readTitle(entry),
    type: readDocumentType(entry),
    tags,
    detail: tags.join(", ") || "Untagged",
  };
}

function readTitle(entry: MindooDBAppViewEntry) {
  const value = entry.columnValues.subject;
  return typeof value === "string" && value.trim() ? value : entry.docId ?? "Untitled document";
}

function readDocumentType(entry: MindooDBAppViewEntry): "markdown" | "word" {
  return entry.columnValues.type === "word" ? "word" : "markdown";
}

function readCategoryLabel(entry: MindooDBAppViewEntry) {
  const value = entry.categoryPath.at(-1);
  return value == null || value === "" ? "Untagged" : String(value);
}
