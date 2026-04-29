import { zipSync } from "fflate";
import type { MindooDBAppAttachmentInfo, MindooDBAppDatabase } from "mindoodb-app-sdk";

import {
  parseAttachmentMarkdownUrl,
  readAttachmentBlob,
} from "@/lib/attachmentImages";

const MARKDOWN_MIME_TYPE = "text/markdown;charset=utf-8";
const ZIP_MIME_TYPE = "application/zip";

type BrowserFileHandle = {
  createWritable(): Promise<{
    write(data: Blob): Promise<void>;
    close(): Promise<void>;
  }>;
};

type SaveFilePickerWindow = Window & {
  showSaveFilePicker?: (options?: {
    suggestedName?: string;
    types?: Array<{
      description: string;
      accept: Record<string, string[]>;
    }>;
  }) => Promise<BrowserFileHandle>;
};

export interface ExportMarkdownPackageOptions {
  database: MindooDBAppDatabase;
  documentId: string;
  markdown: string;
  title: string;
  attachments: readonly MindooDBAppAttachmentInfo[];
}

export function createExportFileName(title: string, extension: "md" | "zip") {
  const baseName = title
    .normalize("NFKC")
    .replace(/[\\/]/g, "-")
    .replace(/[\u0000-\u001f\u007f]/g, "")
    .replace(/\s+/g, " ")
    .trim() || "Untitled document";
  return baseName.toLowerCase().endsWith(`.${extension}`) ? baseName : `${baseName}.${extension}`;
}

function downloadBlobWithAnchor(blob: Blob, fileName: string) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 0);
}

/**
 * Saves one generated file.
 *
 * Chromium browsers can show a real save dialog through the File System Access
 * API. Other browsers still get the same file via the traditional download
 * anchor fallback, which is why the app can offer one menu item everywhere.
 */
export async function saveBlobToDisk(blob: Blob, fileName: string) {
  const savePicker = (window as SaveFilePickerWindow).showSaveFilePicker;
  if (savePicker) {
    try {
      const handle = await savePicker({
        suggestedName: fileName,
        types: [{
          description: fileName.endsWith(".zip") ? "ZIP archive" : "Markdown file",
          accept: {
            [blob.type || "application/octet-stream"]: [fileName.slice(fileName.lastIndexOf("."))],
          },
        }],
      });
      const writable = await handle.createWritable();
      await writable.write(blob);
      await writable.close();
      return true;
    } catch (error) {
      if (error instanceof DOMException && error.name === "AbortError") {
        return false;
      }
      throw error;
    }
  }

  downloadBlobWithAnchor(blob, fileName);
  return true;
}

export function createMarkdownExportBlob(markdown: string) {
  return new Blob([markdown], { type: MARKDOWN_MIME_TYPE });
}

function sanitizeZipPathSegment(segment: string, fallback: string) {
  const sanitized = segment
    .normalize("NFKC")
    .replace(/[\\/]/g, "-")
    .replace(/[\u0000-\u001f\u007f]/g, "")
    .trim();
  return !sanitized || sanitized === "." || sanitized === ".." ? fallback : sanitized;
}

function createUniqueZipPath(usedPaths: Set<string>, requestedPath: string) {
  if (!usedPaths.has(requestedPath)) {
    usedPaths.add(requestedPath);
    return requestedPath;
  }

  const lastSlashIndex = requestedPath.lastIndexOf("/");
  const directory = lastSlashIndex >= 0 ? requestedPath.slice(0, lastSlashIndex + 1) : "";
  const fileName = lastSlashIndex >= 0 ? requestedPath.slice(lastSlashIndex + 1) : requestedPath;
  const dotIndex = fileName.lastIndexOf(".");
  const name = dotIndex > 0 ? fileName.slice(0, dotIndex) : fileName;
  const extension = dotIndex > 0 ? fileName.slice(dotIndex) : "";

  for (let copyIndex = 2; ; copyIndex += 1) {
    const candidate = `${directory}${name}-${copyIndex}${extension}`;
    if (!usedPaths.has(candidate)) {
      usedPaths.add(candidate);
      return candidate;
    }
  }
}

export function createAttachmentExportPlan(attachments: readonly MindooDBAppAttachmentInfo[]) {
  const usedPaths = new Set<string>(["document.md"]);
  return attachments.map((attachment) => {
    const sourceName = attachment.fileName || attachment.attachmentId;
    const safeSegments = sourceName
      .split(/[\\/]+/)
      .filter(Boolean)
      .map((segment, index) => sanitizeZipPathSegment(segment, `attachment-${index + 1}`));
    const safePath = createUniqueZipPath(usedPaths, `attachments/${safeSegments.join("/") || "attachment"}`);
    return {
      attachment,
      exportPath: safePath,
    };
  });
}

function encodeMarkdownPath(path: string) {
  return path.split("/").map((segment) => encodeURIComponent(segment)).join("/");
}

function copyBytesToArrayBuffer(bytes: Uint8Array) {
  const copy = new Uint8Array(bytes.byteLength);
  copy.set(bytes);
  return copy.buffer;
}

function encodeUtf8ForZip(value: string) {
  const encoded = new TextEncoder().encode(value);
  const copy = new Uint8Array(encoded.byteLength);
  copy.set(encoded);
  return copy;
}

export function rewriteMarkdownAttachmentUrls(
  markdown: string,
  exportPlan: ReturnType<typeof createAttachmentExportPlan>,
) {
  const pathsByAttachmentName = new Map<string, string>();
  for (const entry of exportPlan) {
    pathsByAttachmentName.set(entry.attachment.fileName, entry.exportPath);
    pathsByAttachmentName.set(entry.attachment.attachmentId, entry.exportPath);
  }

  return markdown.replace(/mindoodb-attachment:[^\s)"']+/g, (url) => {
    const attachmentName = parseAttachmentMarkdownUrl(url);
    const exportPath = attachmentName ? pathsByAttachmentName.get(attachmentName) : null;
    return exportPath ? encodeMarkdownPath(exportPath) : url;
  });
}

async function readAttachmentForExport(
  database: MindooDBAppDatabase,
  documentId: string,
  attachment: MindooDBAppAttachmentInfo,
) {
  try {
    return await readAttachmentBlob(database, documentId, attachment.fileName, attachment.mimeType);
  } catch (error) {
    if (attachment.attachmentId === attachment.fileName) {
      throw error;
    }
    return await readAttachmentBlob(database, documentId, attachment.attachmentId, attachment.mimeType);
  }
}

/**
 * Builds a portable package:
 *
 * - `document.md` contains the current markdown body.
 * - `attachments/...` contains every current MindooDoc attachment.
 * - TeamEdit attachment URLs are rewritten to point at the exported files.
 */
export async function createMarkdownPackageBytes(options: ExportMarkdownPackageOptions) {
  const exportPlan = createAttachmentExportPlan(options.attachments);
  const rewrittenMarkdown = rewriteMarkdownAttachmentUrls(options.markdown, exportPlan);
  const zipEntries: Record<string, Uint8Array> = {
    "document.md": encodeUtf8ForZip(rewrittenMarkdown),
  };

  for (const entry of exportPlan) {
    const blob = await readAttachmentForExport(options.database, options.documentId, entry.attachment);
    zipEntries[entry.exportPath] = new Uint8Array(await blob.arrayBuffer());
  }

  return zipSync(zipEntries);
}

export async function createMarkdownPackageBlob(options: ExportMarkdownPackageOptions) {
  return new Blob([copyBytesToArrayBuffer(await createMarkdownPackageBytes(options))], { type: ZIP_MIME_TYPE });
}

export async function exportMarkdownFile(markdown: string, title: string) {
  return await saveBlobToDisk(createMarkdownExportBlob(markdown), createExportFileName(title, "md"));
}

export async function exportMarkdownPackage(options: ExportMarkdownPackageOptions) {
  return await saveBlobToDisk(
    await createMarkdownPackageBlob(options),
    createExportFileName(options.title, "zip"),
  );
}
