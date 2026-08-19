import { strFromU8, unzipSync } from "fflate";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { MindooDBAppDatabase } from "mindoodb-app-sdk";

import { createAttachmentMarkdownUrl } from "./attachmentImages";
import {
  createAttachmentExportPlan,
  createExportFileName,
  createMarkdownPackageBytes,
  rewriteMarkdownAttachmentUrls,
  saveBlobToDisk,
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
      revisionId: "rev-1",
    });
    const entries = unzipSync(packageBytes);

    expect(strFromU8(entries["document.md"])).toBe("![Image](attachments/image.png)");
    expect([...entries["attachments/image.png"]]).toEqual([104, 105]);
    expect(database.attachments.openReadStream).toHaveBeenCalledWith("doc-1", "image.png", { revisionId: "rev-1" });
    expect(close).toHaveBeenCalled();
  });
});

describe("saveBlobToDisk", () => {
  const originalShowSaveFilePicker = (
    window as Window & { showSaveFilePicker?: unknown }
  ).showSaveFilePicker;
  let createObjectURL: ReturnType<typeof vi.fn>;
  let revokeObjectURL: ReturnType<typeof vi.fn>;
  let click: ReturnType<typeof vi.fn>;

  beforeEach(() => {
    createObjectURL = vi.fn(() => "blob:download");
    revokeObjectURL = vi.fn();
    click = vi.fn();
    vi.stubGlobal("URL", {
      createObjectURL,
      revokeObjectURL,
    });
    vi.spyOn(document, "createElement").mockImplementation((tagName: string) => {
      if (tagName === "a") {
        return {
          href: "",
          download: "",
          click,
        } as unknown as HTMLAnchorElement;
      }
      return document.createElement(tagName);
    });
  });

  afterEach(() => {
    vi.restoreAllMocks();
    if (originalShowSaveFilePicker) {
      Object.defineProperty(window, "showSaveFilePicker", {
        configurable: true,
        value: originalShowSaveFilePicker,
      });
    } else {
      delete (window as Window & { showSaveFilePicker?: unknown }).showSaveFilePicker;
    }
  });

  it("falls back to an anchor download when the save picker is blocked in an iframe", async () => {
    Object.defineProperty(window, "showSaveFilePicker", {
      configurable: true,
      value: vi.fn(async () => {
        throw new DOMException(
          "Failed to execute 'showSaveFilePicker' on 'Window': Cross origin sub frames aren't allowed to show a file picker.",
          "SecurityError",
        );
      }),
    });

    const saved = await saveBlobToDisk(
      new Blob(["docx"], { type: "application/octet-stream" }),
      "notes.docx",
    );

    expect(saved).toBe(true);
    expect(click).toHaveBeenCalledOnce();
    expect(createObjectURL).toHaveBeenCalledOnce();
  });

  it("does not fall back when the user cancels the save picker", async () => {
    Object.defineProperty(window, "showSaveFilePicker", {
      configurable: true,
      value: vi.fn(async () => {
        const error = new Error("The user aborted a request.");
        error.name = "AbortError";
        throw error;
      }),
    });

    const saved = await saveBlobToDisk(new Blob(["docx"]), "notes.docx");

    expect(saved).toBe(false);
    expect(click).not.toHaveBeenCalled();
  });
});
