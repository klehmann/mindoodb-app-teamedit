import { describe, expect, it, vi } from "vitest";
import type { MindooDBAppDatabase } from "mindoodb-app-sdk";

import {
  createAttachmentMarkdownUrl,
  createUniqueImageAttachmentName,
  extractAttachmentMarkdownUrls,
  formatAttachmentSize,
  parseAttachmentMarkdownUrl,
  readAttachmentBlob,
  sanitizeAttachmentFileName,
  uploadFileAttachment,
} from "./attachmentImages";

describe("attachmentImages", () => {
  it("round-trips stable markdown attachment URLs", () => {
    const url = createAttachmentMarkdownUrl("teamedit-images/image one.png");
    expect(url).toBe("mindoodb-attachment:teamedit-images%2Fimage%20one.png");
    expect(parseAttachmentMarkdownUrl(url)).toBe("teamedit-images/image one.png");
    expect(parseAttachmentMarkdownUrl("https://example.com/image.png")).toBeNull();
  });

  it("encodes markdown-sensitive attachment filename characters", () => {
    const url = createAttachmentMarkdownUrl("teamedit-images/photo (final)!.png");
    expect(url).toBe("mindoodb-attachment:teamedit-images%2Fphoto%20%28final%29%21.png");
    expect(parseAttachmentMarkdownUrl(url)).toBe("teamedit-images/photo (final)!.png");
  });

  it("sanitizes file names without hiding the original extension", () => {
    expect(sanitizeAttachmentFileName("../photo one.png")).toBe("..-photo one.png");
    expect(sanitizeAttachmentFileName("")).toBe("attachment");
    expect(createUniqueImageAttachmentName("photo.png")).toMatch(/^teamedit-images\/.+-photo\.png$/);
  });

  it("extracts TeamEdit attachment image URLs from markdown", () => {
    const first = createAttachmentMarkdownUrl("one.png");
    const second = createAttachmentMarkdownUrl("two.png");
    expect(extractAttachmentMarkdownUrls(`![one](${first})\n![remote](https://example.com/x.png)\n![two](${second} "Two")`))
      .toEqual([first, second]);
  });

  it("extracts legacy image URLs that contain raw balanced parentheses", () => {
    const legacyUrl = "mindoodb-attachment:teamedit-images%2Fphoto%20(1).png";
    expect(extractAttachmentMarkdownUrls(`![legacy](${legacyUrl})`)).toEqual([legacyUrl]);
  });

  it("formats attachment sizes for compact display", () => {
    expect(formatAttachmentSize(42)).toBe("42 B");
    expect(formatAttachmentSize(1536)).toBe("1.5 KB");
    expect(formatAttachmentSize(12 * 1024 * 1024)).toBe("12 MB");
  });

  it("streams uploads through the SDK attachment write stream", async () => {
    const write = vi.fn<(_: Uint8Array) => Promise<void>>().mockResolvedValue(undefined);
    const close = vi.fn<() => Promise<void>>().mockResolvedValue(undefined);
    const abort = vi.fn<() => Promise<void>>().mockResolvedValue(undefined);
    const openWriteStream = vi.fn().mockResolvedValue({ write, close, abort });
    const database = {
      attachments: { openWriteStream },
    } as unknown as MindooDBAppDatabase;

    await uploadFileAttachment(database, "doc-1", "hello.txt", new File(["hello"], "hello.txt", { type: "text/plain" }));

    expect(openWriteStream).toHaveBeenCalledWith("doc-1", "hello.txt", "text/plain");
    expect(write).toHaveBeenCalledWith(new Uint8Array([104, 101, 108, 108, 111]));
    expect(close).toHaveBeenCalled();
    expect(abort).not.toHaveBeenCalled();
  });

  it("reads attachment streams into blobs", async () => {
    const close = vi.fn<() => Promise<void>>().mockResolvedValue(undefined);
    const read = vi.fn<() => Promise<Uint8Array | null>>()
      .mockResolvedValueOnce(new Uint8Array([1, 2]))
      .mockResolvedValueOnce(new Uint8Array([3]))
      .mockResolvedValueOnce(null);
    const database = {
      attachments: {
        openReadStream: vi.fn().mockResolvedValue({ read, close }),
      },
    } as unknown as MindooDBAppDatabase;

    const blob = await readAttachmentBlob(database, "doc-1", "file.bin", "application/octet-stream");

    expect([...new Uint8Array(await blob.arrayBuffer())]).toEqual([1, 2, 3]);
    expect(close).toHaveBeenCalled();
  });
});
