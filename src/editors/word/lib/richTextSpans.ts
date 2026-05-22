import { createEmptyDocument } from "@eigenpal/docx-editor-core";
import { schema } from "@eigenpal/docx-editor-core/prosemirror/schema";
import { toProseDoc } from "@eigenpal/docx-editor-core/prosemirror/conversion";
import { updateDocumentContent } from "@eigenpal/docx-editor-core/prosemirror/conversion/fromProseDoc";
import type { Comment } from "@eigenpal/docx-editor-core/types/content";
import type { Document as DocxDocument } from "@eigenpal/docx-editor-core/types/document";
import type {
  MindooDBAppRichTextMaterializeValue,
  MindooDBAppRichTextSpan,
} from "mindoodb-app-sdk";

type ProseMirrorLikeNode = {
  type: { name: string };
  attrs?: Record<string, unknown>;
  text?: string;
  marks?: ReadonlyArray<{ type: { name: string }; attrs?: Record<string, unknown> }>;
  forEach: (callback: (node: ProseMirrorLikeNode) => void) => void;
};

type RichTextTreeNode =
  | {
      kind: "block";
      type: string;
      attrs: Record<string, unknown>;
      isEmbed: boolean;
      children: RichTextTreeNode[];
    }
  | {
      kind: "text";
      text: string;
      marks?: Record<string, MindooDBAppRichTextMaterializeValue>;
    };

type ProseMirrorNode = ReturnType<typeof schema.node>;
type ProseMirrorMark = ReturnType<typeof schema.mark>;

const A4_PAGE_WIDTH_TWIPS = 11906;
const A4_PAGE_HEIGHT_TWIPS = 16838;
const LETTER_PAGE_WIDTH_TWIPS = 12240;
const LETTER_PAGE_HEIGHT_TWIPS = 15840;

const STRUCTURAL_BLOCKS = new Set([
  "paragraph",
  "heading",
  "blockquote",
  "textBox",
  "table",
  "tableRow",
  "tableCell",
  "tableHeader",
]);

const EMBED_BLOCKS = new Set([
  "field",
  "hardBreak",
  "image",
  "math",
  "shape",
  "tab",
]);

export function documentToRichTextSpans(document: DocxDocument): MindooDBAppRichTextSpan[] {
  const spans = repairPendingCommentMarks(spansFromDocument(document), commentsFromWordDocument(document));
  console.log("[Word rich-text] Serialized document", {
    ...summarizeWordDocument(document),
    commentMarks: summarizeRichTextSpans(spans).commentMarks,
    repairedCommentMarkValues: summarizeRichTextSpans(spans).commentMarkValues,
  });
  return spans;
}

export function commentsFromWordDocument(document: DocxDocument) {
  return mergeCommentArrays(
    readLegacyPackageComments(document),
    document.package.document.comments,
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
  const pmDoc = schema.nodes.doc.create(null, buildTopLevelNodes(buildRichTextTree(repairedSpans)));
  const document = updateDocumentContent(baseDocument, pmDoc);
  if (Array.isArray(comments)) {
    attachCommentsToDocument(document, comments);
  }
  console.log("[Word rich-text] Rebuilt document", {
    comments: Array.isArray(comments) ? comments.length : 0,
    commentMarks: countCommentMarks(repairedSpans),
    repairedCommentMarkValues: summarizeRichTextSpans(repairedSpans).commentMarkValues,
  });
  return document;
}

export function summarizeWordDocument(document: DocxDocument) {
  const canonicalComments = document.package.document.comments;
  const legacyComments = readLegacyPackageComments(document);
  const mergedComments = commentsFromWordDocument(document);
  const spans = spansFromDocument(document);
  return {
    canonicalComments: Array.isArray(canonicalComments) ? canonicalComments.length : null,
    legacyComments: Array.isArray(legacyComments) ? legacyComments.length : null,
    mergedComments: mergedComments.length,
    commentIds: mergedComments.map((comment) => (comment as { id?: unknown }).id),
    commentMarks: summarizeRichTextSpans(spans).commentMarks,
    commentMarkValues: summarizeRichTextSpans(spans).commentMarkValues,
  };
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
  const cloned = structuredClone(comments) as Comment[];
  document.package.document.comments = cloned;
  // Older @eigenpal/docx-editor-vue builds read sidebar comments from the
  // package root; keep that mirror until the app can depend on the fixed build.
  (document.package as { comments?: unknown[] }).comments = structuredClone(comments);
}

export function materializedRichTextToPlainText(value: unknown) {
  return typeof value === "string" ? value.replace(/\uFFFC/g, "") : "";
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

function spansFromDocument(document: DocxDocument): MindooDBAppRichTextSpan[] {
  const proseDoc = toProseDoc(document, {
    styles: document.package.styles,
  }) as unknown as ProseMirrorLikeNode;
  return nodeChildren(proseDoc).flatMap((child) => nodeToSpans(child, []));
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

function buildRichTextTree(spans: MindooDBAppRichTextSpan[]) {
  const roots: RichTextTreeNode[] = [];
  const stack: Array<Extract<RichTextTreeNode, { kind: "block" }>> = [];

  for (const span of spans) {
    if (span.type === "text") {
      const parent = stack.at(-1);
      if (!parent) {
        roots.push({
          kind: "block",
          type: "paragraph",
          attrs: {},
          isEmbed: false,
          children: [{
            kind: "text",
            text: span.value,
            marks: span.marks,
          }],
        });
      } else {
        parent.children.push({
          kind: "text",
          text: span.value,
          marks: span.marks,
        });
      }
      continue;
    }

    const blockType = richTextValueToString(span.value.type) || "paragraph";
    const parents = Array.isArray(span.value.parents) ? span.value.parents : [];
    stack.splice(parents.length);
    const block: Extract<RichTextTreeNode, { kind: "block" }> = {
      kind: "block",
      type: blockType,
      attrs: richTextValueToRecord(span.value.attrs),
      isEmbed: Boolean(span.value.isEmbed),
      children: [],
    };
    const parent = stack.at(-1);
    if (parent) {
      parent.children.push(block);
    } else {
      roots.push(block);
    }
    if (!block.isEmbed) {
      stack.push(block);
    }
  }

  return roots;
}

function buildTopLevelNodes(nodes: RichTextTreeNode[]): ProseMirrorNode[] {
  const built = nodes.flatMap((node) => buildNode(node));
  return built.length ? built : [schema.nodes.paragraph.create()];
}

function buildNode(node: RichTextTreeNode): ProseMirrorNode[] {
  if (node.kind === "text") {
    if (!node.text) {
      return [];
    }
    return [schema.text(node.text, marksFromRichText(node.marks))];
  }

  const nodeType = schema.nodes[node.type] ?? schema.nodes.paragraph;
  const children = node.children.flatMap((child) => buildNode(child));
  if (node.isEmbed) {
    try {
      return [nodeType.create(node.attrs)];
    } catch {
      return [];
    }
  }
  if (node.type === "table" && children.length === 0) {
    children.push(createEmptyTableRow());
  }
  if (node.type === "tableRow" && children.length === 0) {
    children.push(createEmptyTableCell());
  }
  if ((node.type === "tableCell" || node.type === "tableHeader" || node.type === "textBox") && children.length === 0) {
    children.push(schema.nodes.paragraph.create());
  }

  try {
    return [nodeType.createAndFill(node.attrs, children) ?? nodeType.create(node.attrs, children)];
  } catch {
    if (node.type === "paragraph") {
      return [schema.nodes.paragraph.create(null, children.filter((child) => child.isInline))];
    }
    return [schema.nodes.paragraph.create()];
  }
}

function createEmptyTableRow() {
  return schema.nodes.tableRow.create(null, [createEmptyTableCell()]);
}

function createEmptyTableCell() {
  return schema.nodes.tableCell.create(null, [schema.nodes.paragraph.create()]);
}

function marksFromRichText(
  marks: Record<string, MindooDBAppRichTextMaterializeValue> | undefined,
): ProseMirrorMark[] {
  if (!marks) {
    return [];
  }
  return Object.entries(marks).flatMap(([name, value]) => {
    const markType = schema.marks[name];
    if (!markType) {
      return [];
    }
    const attrs = parseMarkAttrs(value);
    return [markType.create(attrs)];
  });
}

function parseMarkAttrs(value: MindooDBAppRichTextMaterializeValue) {
  if (value && typeof value === "object" && !Array.isArray(value)) {
    return richTextValueToRecord(value);
  }
  if (typeof value !== "string") {
    return {};
  }
  try {
    const parsed = JSON.parse(value);
    return parsed && typeof parsed === "object" && !Array.isArray(parsed)
      ? parsed
      : {};
  } catch {
    return {};
  }
}

function richTextValueToString(value: unknown) {
  if (typeof value === "string") {
    return value;
  }
  if (value && typeof value === "object" && "type" in value && "value" in value) {
    const maybeImmutable = value as { type?: unknown; value?: unknown };
    if (maybeImmutable.type === "immutableString" && typeof maybeImmutable.value === "string") {
      return maybeImmutable.value;
    }
  }
  return "";
}

function richTextValueToRecord(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return {};
  }
  return Object.fromEntries(
    Object.entries(value).map(([key, entry]) => [
      key,
      normalizeRichTextValue(entry),
    ]),
  );
}

function normalizeRichTextValue(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map(normalizeRichTextValue);
  }
  if (value && typeof value === "object") {
    const text = richTextValueToString(value);
    if (text) {
      return text;
    }
    return richTextValueToRecord(value);
  }
  return value;
}

function nodeToSpans(
  node: ProseMirrorLikeNode,
  parents: string[],
): MindooDBAppRichTextSpan[] {
  if (node.text) {
    return [{
      type: "text",
      value: node.text,
      marks: marksToRichText(node.marks),
    }];
  }

  const spans: MindooDBAppRichTextSpan[] = [];
  const isBlock = STRUCTURAL_BLOCKS.has(node.type.name);
  const isEmbed = EMBED_BLOCKS.has(node.type.name);
  const nextParents = isBlock ? [...parents, node.type.name] : parents;

  if (isBlock || isEmbed) {
    spans.push({
      type: "block",
      value: {
        type: immutableString(node.type.name),
        parents: parents.map(immutableString),
        attrs: serializableAttrs(node.attrs),
        isEmbed,
      },
    });
  }

  if (isEmbed) {
    return spans;
  }

  for (const child of nodeChildren(node)) {
    spans.push(...nodeToSpans(child, nextParents));
  }

  return spans;
}

function nodeChildren(node: ProseMirrorLikeNode) {
  const children: ProseMirrorLikeNode[] = [];
  node.forEach((child) => children.push(child));
  return children;
}

function marksToRichText(
  marks: ProseMirrorLikeNode["marks"] | undefined,
) {
  if (!marks?.length) {
    return undefined;
  }
  return Object.fromEntries(
    marks.map((mark) => {
      const attrs = serializableAttrs(mark.attrs);
      return [
        mark.type.name,
        Object.keys(attrs).length === 0 ? true : JSON.stringify(attrs),
      ];
    }),
  );
}

function countCommentMarks(spans: MindooDBAppRichTextSpan[]) {
  return spans.filter((span) => span.type === "text" && Boolean(span.marks?.comment)).length;
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
  return structuredClone([...comments.values()]);
}

function readLegacyPackageComments(document: DocxDocument) {
  return (document.package as { comments?: unknown }).comments;
}

function serializableAttrs(
  attrs: Record<string, unknown> | undefined,
): Record<string, MindooDBAppRichTextMaterializeValue> {
  if (!attrs) {
    return {};
  }
  return Object.fromEntries(
    Object.entries(attrs)
      .map(([key, value]) => [key, sanitizeMaterializeValue(value)] as const)
      .filter((entry): entry is readonly [string, MindooDBAppRichTextMaterializeValue] => entry[1] !== undefined),
  );
}

function sanitizeMaterializeValue(
  value: unknown,
): MindooDBAppRichTextMaterializeValue | undefined {
  if (value === undefined) {
    return undefined;
  }
  if (
    value === null ||
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean"
  ) {
    return value;
  }
  if (Array.isArray(value)) {
    return value.map((entry) => sanitizeMaterializeValue(entry) ?? null);
  }
  if (typeof value === "object") {
    return serializableAttrs(value as Record<string, unknown>);
  }
  return String(value);
}

function immutableString(value: string) {
  return {
    type: "immutableString" as const,
    value,
  };
}
