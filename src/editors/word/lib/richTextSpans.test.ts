import { schema } from "@eigenpal/docx-editor-core/prosemirror/schema";
import { updateDocumentContent } from "@eigenpal/docx-editor-core/prosemirror/conversion/fromProseDoc";
import { describe, expect, it } from "vitest";
import { createDocxAutomergeSchemaAdapter } from "@/editors/word/lib/automergeSchemaAdapter";
import { pmNodeToSdkRichTextSpans } from "@/editors/word/lib/automergeSpanUtils";
import {
  attachCommentsToDocument,
  createDefaultWordDocument,
  documentToRichTextSpans,
  richTextSpansToDocument,
} from "@/editors/word/lib/richTextSpans";

const adapter = createDocxAutomergeSchemaAdapter();

describe("richTextSpans", () => {
  it("round-trips a default document through spans", () => {
    const source = createDefaultWordDocument();
    const spans = documentToRichTextSpans(source);
    const roundTripped = richTextSpansToDocument(spans, []);
    const roundTrippedSpans = documentToRichTextSpans(roundTripped);

    expect(roundTrippedSpans.length).toBeGreaterThan(0);
    expect(JSON.stringify(roundTrippedSpans)).toBe(JSON.stringify(spans));
  });

  it("preserves bold mark attrs through a round trip", () => {
    const source = createDefaultWordDocument();
    const paragraph = schema.nodes.paragraph.create(
      null,
      schema.text("Hello", [schema.marks.bold.create()]),
    );
    const proseDoc = schema.nodes.doc.create(null, [paragraph]);
    const markedDocument = updateDocumentContent(source, proseDoc);

    const spans = documentToRichTextSpans(markedDocument);
    const roundTripped = richTextSpansToDocument(spans, []);
    const textSpan = documentToRichTextSpans(roundTripped).find((span) => span.type === "text");

    expect(textSpan?.type).toBe("text");
    if (textSpan?.type === "text") {
      expect(textSpan.marks?.bold).toBe(true);
    }
  });

  it("preserves table structure through a round trip", () => {
    const cell = (text: string) =>
      schema.nodes.tableCell.create(null, [
        schema.nodes.paragraph.create(null, schema.text(text)),
      ]);
    const row = schema.nodes.tableRow.create(null, [cell("A1"), cell("A2")]);
    const table = schema.nodes.table.create(null, [row]);
    const proseDoc = schema.nodes.doc.create(null, [table]);
    const spans = pmNodeToSdkRichTextSpans(adapter, proseDoc);
    const blockTypes = spans
      .filter((span) => span.type === "block")
      .map((span) => (span.type === "block" ? span.value.type : null));

    expect(blockTypes).toContainEqual({ type: "immutableString", value: "table" });
    expect(blockTypes).toContainEqual({ type: "immutableString", value: "tableRow" });
    expect(blockTypes).toContainEqual({ type: "immutableString", value: "tableCell" });

    const roundTripped = richTextSpansToDocument(spans, []);
    const roundTrippedSpans = documentToRichTextSpans(roundTripped);
    const roundTrippedBlockTypes = roundTrippedSpans
      .filter((span) => span.type === "block")
      .map((span) => (span.type === "block" ? span.value.type : null));
    expect(roundTrippedBlockTypes).toEqual(blockTypes);
    expect(
      roundTrippedSpans
        .filter((span) => span.type === "text")
        .map((span) => (span.type === "text" ? span.value : "")),
    ).toEqual(["A1", "A2"]);
  });

  it("preserves resolved table cell backgrounds from themed DOCX shading", () => {
    const cell = schema.nodes.tableCell.create(
      {
        backgroundColor: "F4B183",
        _originalResolvedFill: "F4B183",
        _originalFormatting: {
          shading: {
            fill: {
              themeColor: "accent2",
              themeTint: "99",
            },
            pattern: "clear",
          },
        },
      },
      [schema.nodes.paragraph.create(null, schema.text("Themed"))],
    );
    const row = schema.nodes.tableRow.create(null, [cell]);
    const table = schema.nodes.table.create(null, [row]);
    const proseDoc = schema.nodes.doc.create(null, [table]);
    const spans = pmNodeToSdkRichTextSpans(adapter, proseDoc);

    const roundTripped = richTextSpansToDocument(spans, []);
    const roundTrippedSpans = documentToRichTextSpans(roundTripped);
    const cellSpan = roundTrippedSpans.find((span) =>
      span.type === "block" &&
      (span.value.type as { value?: string }).value === "tableCell"
    );

    expect(cellSpan?.type).toBe("block");
    if (cellSpan?.type === "block") {
      expect(cellSpan.value.attrs?.backgroundColor).toBe("F4B183");
      expect(cellSpan.value.attrs?._originalFormatting).toMatchObject({
        shading: {
          fill: {
            rgb: "F4B183",
          },
        },
      });
    }
  });

  it("attachCommentsToDocument clones proxy-like comment arrays", () => {
    const document = createDefaultWordDocument();
    const comments = new Proxy([{ id: 1, text: "Review this paragraph" }], {});
    attachCommentsToDocument(document, comments as unknown[]);
    expect(document.package.document.comments).toEqual([{ id: 1, text: "Review this paragraph" }]);
    expect(document.package.document.comments).not.toBe(comments);
  });
});
