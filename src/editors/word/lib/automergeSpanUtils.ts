/**
 * Conversion between ProseMirror documents and JSON-safe MindooDB rich-text spans.
 * Table structure follows the docx-editor-test approach on top of @automerge/prosemirror.
 */

import * as A from "@automerge/automerge";
import { pmNodeToSpans, pmDocFromSpans, type SchemaAdapter } from "@automerge/prosemirror";
import type { Node as ProseMirrorNode } from "prosemirror-model";
import type {
  MindooDBAppRichTextMaterializeValue,
  MindooDBAppRichTextSpan,
} from "mindoodb-app-sdk";

type NativeSpan = A.Span;

export function pmNodeToSdkRichTextSpans(
  adapter: SchemaAdapter,
  node: ProseMirrorNode,
): MindooDBAppRichTextSpan[] {
  return dehydrateRichTextSpans(
    normalizeNativeSpans(pmNodeToNativeRichTextSpans(adapter, node), "pmNodeToSdkRichTextSpans"),
  );
}

function pmNodeToNativeRichTextSpans(
  adapter: SchemaAdapter,
  node: ProseMirrorNode,
): NativeSpan[] {
  return node.type.name === "doc"
    ? nodeChildren(node).flatMap((child) => pmNodeToNativeRichTextSpans(adapter, child))
    : pmNodeToSpansWithDocxTables(adapter, node);
}

export function sdkRichTextSpansToProseDoc(
  adapter: SchemaAdapter,
  spans: MindooDBAppRichTextSpan[],
): ProseMirrorNode {
  const nativeSpans = reviveRichTextSpans(spans);
  return pmDocFromSpans(adapter, nativeSpans);
}

function pmNodeToSpansWithDocxTables(
  adapter: SchemaAdapter,
  node: ProseMirrorNode,
  parents: string[] = [],
): NativeSpan[] {
  if (node.type.name !== "table") {
    return prefixBlockParents(pmNodeToSpans(adapter, node) as NativeSpan[], parents);
  }

  const tableParents = [...parents, "table"];
  const spans: NativeSpan[] = [createBlockSpan("table", parents, node)];

  node.forEach((row: ProseMirrorNode) => {
    const rowParents = [...tableParents, "tableRow"];
    spans.push(createBlockSpan("tableRow", tableParents, row));

    row.forEach((cell: ProseMirrorNode) => {
      const cellType = cell.type.name === "tableHeader" ? "tableHeader" : "tableCell";
      spans.push(createBlockSpan(cellType, rowParents, cell));

      cell.forEach((child: ProseMirrorNode) => {
        spans.push(...pmNodeToSpansWithDocxTables(adapter, child, [...rowParents, cellType]));
      });
    });
  });

  return spans;
}

function createBlockSpan(type: string, parents: string[], node: ProseMirrorNode): NativeSpan {
  return {
    type: "block",
    value: {
      type: new A.ImmutableString(type),
      parents: parents.map((parent) => new A.ImmutableString(parent)),
      attrs: serializableNodeAttrs(node),
      isEmbed: false,
    },
  };
}

function prefixBlockParents(spans: NativeSpan[], parents: string[]): NativeSpan[] {
  if (parents.length === 0) {
    return spans;
  }

  return spans.map((span) => {
    if (span.type !== "block" || !isRecord(span.value)) {
      return span;
    }

    const existingParents = Array.isArray(span.value.parents) ? span.value.parents : [];
    return {
      ...span,
      value: {
        ...span.value,
        parents: [
          ...parents.map((parent) => new A.ImmutableString(parent)),
          ...existingParents,
        ],
      },
    };
  });
}

function serializableNodeAttrs(node: ProseMirrorNode): Record<string, A.MaterializeValue> {
  const attrs = sanitizeMaterializeValue(
    normalizeNodeAttrsForStorage(node),
  );

  return isRecord(attrs) ? (attrs as Record<string, A.MaterializeValue>) : {};
}

function normalizeNodeAttrsForStorage(node: ProseMirrorNode): Record<string, unknown> {
  const attrs = Object.fromEntries(Object.entries(node.attrs).filter(([, value]) => value !== null));
  if ((node.type.name !== "tableCell" && node.type.name !== "tableHeader") || typeof attrs.backgroundColor !== "string") {
    return attrs;
  }

  const originalFormatting = attrs._originalFormatting;
  if (!isRecord(originalFormatting)) {
    return attrs;
  }

  return {
    ...attrs,
    _originalFormatting: {
      ...originalFormatting,
      shading: {
        ...(isRecord(originalFormatting.shading) ? originalFormatting.shading : {}),
        fill: { rgb: attrs.backgroundColor },
      },
    },
  };
}

function normalizeNativeSpans(spans: NativeSpan[], context: string): NativeSpan[] {
  return spans.map((span, index) => {
    if (span.type !== "block") {
      return span;
    }

    if (isRecord(span.value)) {
      return {
        ...span,
        value: sanitizeBlockValue(span.value),
      };
    }

    console.error("[Word rich-text] Replacing invalid block span", { context, index, span });
    return {
      type: "block",
      value: {
        type: new A.ImmutableString("paragraph"),
        parents: [],
        attrs: { invalidBlockValue: String(span.value) },
        isEmbed: false,
      },
    };
  });
}

function sanitizeBlockValue(value: Record<string, unknown>): Record<string, A.MaterializeValue> {
  const rawAttrs = isRecord(value.attrs) ? value.attrs : {};
  const attrs = sanitizeMaterializeValue(rawAttrs);
  const rawParents = Array.isArray(value.parents) ? value.parents : [];

  return {
    type:
      typeof value.type === "string" || A.isImmutableString(value.type)
        ? (value.type as A.MaterializeValue)
        : new A.ImmutableString("paragraph"),
    parents: rawParents.map((parent) =>
      typeof parent === "string" || A.isImmutableString(parent)
        ? (parent as A.MaterializeValue)
        : new A.ImmutableString(String(parent)),
    ),
    attrs: isRecord(attrs) ? attrs : {},
    isEmbed: Boolean(value.isEmbed),
  };
}

export function dehydrateRichTextSpans(spans: NativeSpan[]): MindooDBAppRichTextSpan[] {
  return spans.map((span) => {
    if (span.type === "text") {
      return {
        type: "text",
        value: span.value,
        marks: span.marks
          ? dehydrateRichTextValue(span.marks) as Record<string, MindooDBAppRichTextMaterializeValue>
          : undefined,
      };
    }

    return {
      type: "block",
      value: dehydrateRichTextValue(span.value) as Record<string, MindooDBAppRichTextMaterializeValue>,
    };
  });
}

export function reviveRichTextSpans(spans: MindooDBAppRichTextSpan[]): NativeSpan[] {
  return spans.map((span) => {
    if (span.type === "text") {
      return {
        type: "text",
        value: span.value,
        marks: span.marks ? reviveRichTextValue(span.marks) : undefined,
      };
    }

    return {
      type: "block",
      value: reviveRichTextValue(span.value),
    };
  }) as NativeSpan[];
}

function dehydrateRichTextValue(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map((entry) => dehydrateRichTextValue(entry));
  }
  if (!value || typeof value !== "object") {
    return value;
  }
  if (A.isImmutableString(value)) {
    return { type: "immutableString", value: String(value) };
  }
  const record = value as Record<string, unknown>;
  return Object.fromEntries(
    Object.entries(record).map(([key, entry]) => [key, dehydrateRichTextValue(entry)]),
  );
}

function reviveRichTextValue(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map((entry) => reviveRichTextValue(entry));
  }
  if (!value || typeof value !== "object") {
    return value;
  }
  const record = value as Record<string, unknown>;
  if (record.type === "immutableString" && typeof record.value === "string") {
    return new A.ImmutableString(record.value);
  }
  return Object.fromEntries(
    Object.entries(record).map(([key, entry]) => [key, reviveRichTextValue(entry)]),
  );
}

function sanitizeMaterializeValue(value: unknown): A.MaterializeValue | undefined {
  if (value === undefined) {
    return undefined;
  }

  if (
    value === null ||
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean" ||
    value instanceof Date ||
    value instanceof Uint8Array ||
    A.isImmutableString(value)
  ) {
    return value as A.MaterializeValue;
  }

  if (Array.isArray(value)) {
    return value.map((item) => sanitizeMaterializeValue(item) ?? null) as A.MaterializeValue;
  }

  if (isRecord(value)) {
    const sanitized: Record<string, A.MaterializeValue> = {};
    for (const [key, nestedValue] of Object.entries(value)) {
      const nextValue = sanitizeMaterializeValue(nestedValue);
      if (nextValue !== undefined) {
        sanitized[key] = nextValue;
      }
    }
    return sanitized as A.MaterializeValue;
  }

  return String(value);
}

function nodeChildren(node: ProseMirrorNode): ProseMirrorNode[] {
  const children: ProseMirrorNode[] = [];
  node.forEach((child: ProseMirrorNode) => children.push(child));
  return children;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

export function richTextSpansToPlainText(spans: MindooDBAppRichTextSpan[]) {
  return spans
    .filter((span) => span.type === "text")
    .map((span) => (span.type === "text" ? span.value : ""))
    .join("");
}

export function materializedRichTextToPlainText(value: unknown) {
  return typeof value === "string" ? value.replace(/\uFFFC/g, "") : "";
}

export function updateSpansWithDiagnostics<T>(
  doc: A.Doc<T>,
  path: A.Prop[],
  rawSpans: NativeSpan[],
  config: A.UpdateSpansConfig | undefined,
  context: string,
): void {
  const normalizedSpans = normalizeNativeSpans(rawSpans, context);

  try {
    A.updateSpans(doc, path, normalizedSpans, config);
  } catch (error) {
    console.error("[Word rich-text] updateSpans failed", error, {
      context,
      path,
      config,
      rawSpans,
      normalizedSpans,
    });
    throw error;
  }
}
