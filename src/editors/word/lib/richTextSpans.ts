import { createEmptyDocument } from "@eigenpal/docx-editor-core";
import { schema } from "@eigenpal/docx-editor-core/prosemirror/schema";
import { toProseDoc } from "@eigenpal/docx-editor-core/prosemirror/conversion";
import { updateDocumentContent } from "@eigenpal/docx-editor-core/prosemirror/conversion/fromProseDoc";
import type { Comment } from "@eigenpal/docx-editor-core/types/content";
import type { Document as DocxDocument } from "@eigenpal/docx-editor-core/types/document";
import type { MindooDBAppRichTextSpan } from "mindoodb-app-sdk";
import { createDocxAutomergeSchemaAdapter } from "@/editors/word/lib/automergeSchemaAdapter";
import {
  materializedRichTextToPlainText,
  pmNodeToSdkRichTextSpans,
  richTextSpansToPlainText,
  sdkRichTextSpansToProseDoc,
} from "@/editors/word/lib/automergeSpanUtils";

export { materializedRichTextToPlainText, richTextSpansToPlainText };

const A4_PAGE_WIDTH_TWIPS = 11906;
const A4_PAGE_HEIGHT_TWIPS = 16838;
const LETTER_PAGE_WIDTH_TWIPS = 12240;
const LETTER_PAGE_HEIGHT_TWIPS = 15840;

const adapter = createDocxAutomergeSchemaAdapter();

export function documentToRichTextSpans(document: DocxDocument): MindooDBAppRichTextSpan[] {
  const proseDoc = toProseDoc(document, {
    styles: document.package.styles,
  });
  return repairPendingCommentMarks(
    pmNodeToSdkRichTextSpans(adapter, proseDoc),
    commentsFromWordDocument(document),
  );
}

export function richTextSpansToDocument(
  spans: MindooDBAppRichTextSpan[],
  comments: unknown,
): DocxDocument {
  const repairedSpans = Array.isArray(comments)
    ? repairPendingCommentMarks(spans, comments)
    : spans;
  const baseDocument = createDefaultWordDocument();
  const pmDoc = sdkRichTextSpansToProseDoc(adapter, repairedSpans);
  const document = updateDocumentContent(baseDocument, pmDoc);
  if (Array.isArray(comments)) {
    attachCommentsToDocument(document, comments);
  }
  return document;
}

export function commentsFromWordDocument(document: DocxDocument) {
  return mergeCommentArrays(
    readLegacyPackageComments(document),
    document.package.document.comments,
  );
}

export function summarizeWordDocument(
  document: DocxDocument,
  spans?: MindooDBAppRichTextSpan[],
) {
  const canonicalComments = document.package.document.comments;
  const legacyComments = readLegacyPackageComments(document);
  const mergedComments = commentsFromWordDocument(document);
  const summary = {
    canonicalComments: Array.isArray(canonicalComments) ? canonicalComments.length : null,
    legacyComments: Array.isArray(legacyComments) ? legacyComments.length : null,
    mergedComments: mergedComments.length,
    commentIds: mergedComments.map((comment) => (comment as { id?: unknown }).id),
  };

  return spans ? { ...summary, ...summarizeRichTextSpans(spans) } : summary;
}

export function summarizeRichTextSpans(spans: MindooDBAppRichTextSpan[]) {
  const commentMarkValues = spans
    .filter((span) => span.type === "text" && Boolean(span.marks?.comment))
    .map((span) => span.type === "text" ? span.marks?.comment : undefined)
    .slice(0, 10);
  return {
    spans: spans.length,
    commentMarks: commentMarkValues.length,
    commentMarkValues,
  };
}

export function attachCommentsToDocument(document: DocxDocument, comments: unknown[]) {
  const cloned = cloneJsonSafe(comments) as Comment[];
  document.package.document.comments = cloned;
  (document.package as { comments?: unknown[] }).comments = cloneJsonSafe(comments);
}

export function plainTextToWordDocument(value: unknown): DocxDocument {
  const text = materializedRichTextToPlainText(value);
  return text ? createDefaultWordDocumentWithText(text) : createDefaultWordDocument();
}

export function createDefaultWordDocument(): DocxDocument {
  const pageSize = getDefaultWordPageSize();
  return createEmptyDocument({
    pageWidth: pageSize.width,
    pageHeight: pageSize.height,
    orientation: "portrait",
  });
}

function getDefaultWordPageSize() {
  return isLetterPaperLocale(getBrowserLocale())
    ? { width: LETTER_PAGE_WIDTH_TWIPS, height: LETTER_PAGE_HEIGHT_TWIPS }
    : { width: A4_PAGE_WIDTH_TWIPS, height: A4_PAGE_HEIGHT_TWIPS };
}

function getBrowserLocale() {
  return typeof navigator === "undefined" ? "" : navigator.language;
}

function isLetterPaperLocale(locale: string) {
  return locale.replace("_", "-").toLowerCase() === "en-us";
}

function createDefaultWordDocumentWithText(text: string) {
  const paragraph = schema.nodes.paragraph.create(null, schema.text(text));
  return updateDocumentContent(createDefaultWordDocument(), schema.nodes.doc.create(null, [paragraph]));
}

function repairPendingCommentMarks(
  spans: MindooDBAppRichTextSpan[],
  comments: unknown[],
): MindooDBAppRichTextSpan[] {
  const realCommentIds = comments
    .map((comment) => comment && typeof comment === "object" ? (comment as { id?: unknown }).id : undefined)
    .filter((id): id is number => typeof id === "number" && id >= 0);
  if (realCommentIds.length === 0) {
    return spans;
  }
  const replacementId = Math.max(...realCommentIds);
  let repaired = false;
  const nextSpans = spans.map((span) => {
    if (span.type !== "text" || !span.marks?.comment) {
      return span;
    }
    const attrs = parseMarkAttrs(span.marks.comment);
    if (attrs.commentId !== -1) {
      return span;
    }
    repaired = true;
    return {
      ...span,
      marks: {
        ...span.marks,
        comment: JSON.stringify({ ...attrs, commentId: replacementId }),
      },
    };
  });
  if (repaired) {
    console.warn("[Word rich-text] Repaired pending comment mark", {
      replacementId,
      commentIds: realCommentIds,
    });
  }
  return nextSpans;
}

function parseMarkAttrs(value: unknown) {
  if (value && typeof value === "object" && !Array.isArray(value)) {
    return value as Record<string, unknown>;
  }
  if (typeof value !== "string") {
    return {};
  }
  try {
    const parsed = JSON.parse(value);
    return parsed && typeof parsed === "object" && !Array.isArray(parsed)
      ? parsed as Record<string, unknown>
      : {};
  } catch {
    return {};
  }
}

function mergeCommentArrays(...sources: unknown[]) {
  const comments = new Map<number | string, unknown>();
  for (const source of sources) {
    if (!Array.isArray(source)) {
      continue;
    }
    for (const comment of source) {
      if (!comment || typeof comment !== "object") {
        continue;
      }
      const id = (comment as { id?: unknown }).id;
      comments.set(typeof id === "number" || typeof id === "string" ? id : comments.size, comment);
    }
  }
  return cloneJsonSafe([...comments.values()]);
}

function cloneJsonSafe<T>(value: T): T {
  if (value === undefined || value === null) {
    return value;
  }

  try {
    return JSON.parse(JSON.stringify(value)) as T;
  } catch {
    if (Array.isArray(value)) {
      return value.map((item) => cloneJsonSafe(item)) as T;
    }
    if (typeof value === "object") {
      return Object.fromEntries(
        Object.entries(value as Record<string, unknown>).map(([key, nestedValue]) => [
          key,
          cloneJsonSafe(nestedValue),
        ]),
      ) as T;
    }
    return value;
  }
}

function readLegacyPackageComments(document: DocxDocument) {
  return (document.package as { comments?: unknown }).comments;
}
