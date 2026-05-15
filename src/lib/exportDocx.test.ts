import { strFromU8, unzipSync } from "fflate";
import { describe, expect, it, vi } from "vitest";

import { createAttachmentMarkdownUrl } from "@/lib/attachmentImages";
import { createExportFileName } from "@/lib/exportMarkdown";
import {
  collectNonImageAttachmentReferences,
  createDocxExportBlob,
  type ExportDocxOptions,
} from "@/lib/exportDocx";

async function readDocxEntries(blob: Blob) {
  return unzipSync(new Uint8Array(await blob.arrayBuffer()));
}

async function readDocumentXml(blob: Blob) {
  const entries = await readDocxEntries(blob);
  return strFromU8(entries["word/document.xml"]);
}

describe("exportDocx", () => {
  it("creates safe docx export file names", () => {
    expect(createExportFileName("../Sprint Notes", "docx")).toBe("..-Sprint Notes.docx");
    expect(createExportFileName("Weekly.docx", "docx")).toBe("Weekly.docx");
  });

  it("converts common TeamEdit markdown structures to a docx document", async () => {
    const blob = await createDocxExportBlob({
      title: "Sprint Notes",
      markdown: [
        "# Heading",
        "",
        "Paragraph with **bold**, *italic*, ==marked==, ~~gone~~, `code`, and [link](https://example.com).",
        "",
        "- [x] Done",
        "- Open",
        "",
        "| A | B |",
        "| - | - |",
        "| C | D |",
        "",
        "::: tip Useful",
        "Callout body",
        ":::",
        "",
        "```ts",
        "const value = 1;",
        "```",
      ].join("\n"),
    });

    const xml = await readDocumentXml(blob);
    expect(blob.type).toBe("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    expect(xml).toContain("Sprint Notes");
    expect(xml).toContain("Heading");
    expect(xml).toContain("Paragraph with");
    expect(xml).toContain("[x]");
    expect(xml).toContain("Callout body");
    expect(xml).toContain("const value = 1;");
    expect(xml).toContain("<w:tbl>");
  });

  it("falls back to image alt text when an image cannot be embedded", async () => {
    const blob = await createDocxExportBlob({
      title: "Images",
      markdown: "![Fallback caption](https://example.com/missing.png)",
      fetchImage: vi.fn<NonNullable<ExportDocxOptions["fetchImage"]>>().mockResolvedValue(null),
    });

    await expect(readDocumentXml(blob)).resolves.toContain("[Fallback caption]");
  });

  it("renders Mermaid diagrams to embedded docx media", async () => {
    const renderMermaid = vi.fn<NonNullable<ExportDocxOptions["renderMermaidSvg"]>>()
      .mockResolvedValue(`<svg width="120" height="60" viewBox="0 0 120 60"><text x="0" y="12">Diagram</text></svg>`);
    const svgToPngBytes = vi.fn<NonNullable<ExportDocxOptions["svgToPngBytes"]>>()
      .mockResolvedValue({
        bytes: new Uint8Array([137, 80, 78, 71]),
        type: "png",
        width: 120,
        height: 60,
      });

    const blob = await createDocxExportBlob({
      title: "Diagram",
      markdown: "```mermaid\ngraph TD\n  A --> B\n```",
      renderMermaidSvg: renderMermaid,
      svgToPngBytes,
    });
    const entries = await readDocxEntries(blob);

    expect(renderMermaid).toHaveBeenCalledWith("graph TD\n  A --> B\n", expect.any(Object));
    expect(svgToPngBytes).toHaveBeenCalledWith(expect.stringContaining("<svg"));
    expect(Object.keys(entries).some((path) => path.startsWith("word/media/"))).toBe(true);
  });

  it("collects non-image attachment links for the docx appendix", async () => {
    const attachmentUrl = createAttachmentMarkdownUrl("docs/spec.pdf");
    const references = collectNonImageAttachmentReferences(`[Spec](${attachmentUrl})`, [
      { attachmentId: "att-1", fileName: "docs/spec.pdf", mimeType: "application/pdf", size: 123 },
    ]);

    expect(references).toEqual([
      {
        label: "Spec",
        attachment: { attachmentId: "att-1", fileName: "docs/spec.pdf", mimeType: "application/pdf", size: 123 },
      },
    ]);

    const blob = await createDocxExportBlob({
      title: "Attachments",
      markdown: `[Spec](${attachmentUrl})`,
      attachments: [
        { attachmentId: "att-1", fileName: "docs/spec.pdf", mimeType: "application/pdf", size: 123 },
      ],
    });

    const xml = await readDocumentXml(blob);
    expect(xml).toContain("Document attachments");
    expect(xml).toContain("docs/spec.pdf");
  });
});
