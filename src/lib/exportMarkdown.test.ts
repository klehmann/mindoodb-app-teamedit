import { strFromU8, unzipSync } from "fflate";
import { describe, expect, it, vi } from "vitest";
import type { MindooDBAppDatabase } from "mindoodb-app-sdk";

import { createAttachmentMarkdownUrl } from "./attachmentImages";
import {
  createAttachmentExportPlan,
  createExportFileName,
  createMarkdownPackageBytes,
  rewriteMarkdownAttachmentUrls,
} from "./exportMarkdown";

describe("exportMarkdown", () => {
  it("creates safe export file names with the requested extension", () => {
    expect(createExportFileName("../Sprint Notes", "md")).toBe("..-Sprint Notes.md");
    expect(createExportFileName("Weekly.md", "md")).toBe("Weekly.md");
    expect(createExportFileName("", "zip")).toBe("Untitled document.zip");
  });

  it("plans unique attachment paths inside the zip package", () => {
    const plan = createAttachmentExportPlan([
      { attachmentId: "id-1", fileName: "teamedit-images/photo.png", mimeType: "image/png", size: 1 },
      { attachmentId: "id-2", fileName: "teamedit-images/photo.png", mimeType: "image/png", size: 1 },
      { attachmentId: "id-3", fileName: "../notes.txt", mimeType: "text/plain", size: 1 },
    ]);

    expect(plan.map((entry) => entry.exportPath)).toEqual([
      "attachments/teamedit-images/photo.png",
      "attachments/teamedit-images/photo-2.png",
      "attachments/attachment-1/notes.txt",
    ]);
  });

  it("rewrites TeamEdit attachment URLs to relative package paths", () => {
    const markdownUrl = createAttachmentMarkdownUrl("teamedit-images/photo one.png");
    const plan = createAttachmentExportPlan([
      { attachmentId: "id-1", fileName: "teamedit-images/photo one.png", mimeType: "image/png", size: 1 },
    ]);

    expect(rewriteMarkdownAttachmentUrls(`![Photo](${markdownUrl})`, plan))
      .toBe("![Photo](attachments/teamedit-images/photo%20one.png)");
  });

  it("builds a zip with rewritten markdown and attachment bytes", async () => {
    const read = vi.fn<() => Promise<Uint8Array | null>>()
      .mockResolvedValueOnce(new Uint8Array([104, 105]))
      .mockResolvedValueOnce(null);
    const close = vi.fn<() => Promise<void>>().mockResolvedValue(undefined);
    const database = {
      attachments: {
        openReadStream: vi.fn().mockResolvedValue({ read, close }),
      },
    } as unknown as MindooDBAppDatabase;
    const markdownUrl = createAttachmentMarkdownUrl("image.png");

    const packageBytes = await createMarkdownPackageBytes({
      database,
      documentId: "doc-1",
      markdown: `![Image](${markdownUrl})`,
      title: "Ignored by zip builder",
      attachments: [
        { attachmentId: "attachment-1", fileName: "image.png", mimeType: "image/png", size: 2 },
      ],
    });
    const entries = unzipSync(packageBytes);

    expect(strFromU8(entries["document.md"])).toBe("![Image](attachments/image.png)");
    expect([...entries["attachments/image.png"]]).toEqual([104, 105]);
    expect(close).toHaveBeenCalled();
  });
});
