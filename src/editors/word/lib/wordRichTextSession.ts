import type { Document as DocxDocument } from "@eigenpal/docx-editor-core/types/document";
import type {
  MindooDBAppDatabase,
  MindooDBAppDocument,
  MindooDBAppDocumentRevisionId,
  MindooDBAppRichTextSpan,
} from "mindoodb-app-sdk";
import {
  createDefaultWordDocument,
  documentToRichTextSpans,
  richTextSpansToDocument,
} from "@/editors/word/lib/richTextSpans";
import { readRichTextSpansFromAutomergeSnapshot } from "@/editors/word/lib/wordAutomergeHandle";

export const WORD_RICH_TEXT_PATH = ["body"] as const;

export async function loadWordAutomergeSnapshot(
  database: MindooDBAppDatabase,
  docId: string,
  options?: { revisionId?: MindooDBAppDocumentRevisionId },
) {
  return database.documents.getAutomergeSnapshot(docId, options);
}

export function wordDocumentFromRichTextSpans(
  spans: MindooDBAppRichTextSpan[],
  comments: unknown,
): DocxDocument {
  return richTextSpansToDocument(spans, comments);
}

export async function seedWordRichTextDocument(
  database: MindooDBAppDatabase,
  document: MindooDBAppDocument,
  docxDocument: DocxDocument,
) {
  return database.documents.update(document.id, {
    richText: [{
      path: [...WORD_RICH_TEXT_PATH],
      spans: documentToRichTextSpans(docxDocument),
    }],
  });
}

export async function loadWordDocumentFromAutomerge(
  database: MindooDBAppDatabase,
  document: MindooDBAppDocument,
  options?: { revisionId?: MindooDBAppDocumentRevisionId },
) {
  const snapshot = await loadWordAutomergeSnapshot(database, document.id, options);
  const spans = readRichTextSpansFromAutomergeSnapshot(snapshot, [...WORD_RICH_TEXT_PATH]);
  const resolvedSpans = spans.length > 0
    ? spans
    : documentToRichTextSpans(createDefaultWordDocument());
  const wordDocument = wordDocumentFromRichTextSpans(resolvedSpans, document.data.comments);
  return {
    snapshot: {
      ...snapshot,
      spans: resolvedSpans,
    },
    wordDocument,
  };
}

/** @deprecated Use {@link loadWordDocumentFromAutomerge}. */
export const loadWordDocumentFromRichText = loadWordDocumentFromAutomerge;
