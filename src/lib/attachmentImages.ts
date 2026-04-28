import type { MindooDBAppDatabase } from "mindoodb-app-sdk";

const ATTACHMENT_MARKDOWN_SCHEME = "mindoodb-attachment:";
const DEFAULT_CHUNK_SIZE = 64 * 1024;

/** Returns true when a markdown image URL points at a MindooDB document attachment. */
export function isAttachmentMarkdownUrl(url: string) {
  return url.startsWith(ATTACHMENT_MARKDOWN_SCHEME);
}

/**
 * Encodes an attachment name as a stable markdown URL.
 *
 * The markdown stores this app-owned URL instead of a `blob:` URL because blob
 * URLs are only valid in the current browser session.
 */
export function createAttachmentMarkdownUrl(attachmentName: string) {
  return `${ATTACHMENT_MARKDOWN_SCHEME}${encodeURIComponent(attachmentName)}`;
}

/** Decodes a TeamEdit attachment markdown URL back to the stored attachment name. */
export function parseAttachmentMarkdownUrl(url: string) {
  if (!isAttachmentMarkdownUrl(url)) {
    return null;
  }
  try {
    return decodeURIComponent(url.slice(ATTACHMENT_MARKDOWN_SCHEME.length));
  } catch {
    return null;
  }
}

/** Keeps user-provided filenames readable while avoiding path separators. */
export function sanitizeAttachmentFileName(fileName: string) {
  const sanitized = fileName
    .normalize("NFKC")
    .replace(/[\\/]/g, "-")
    .replace(/[\u0000-\u001f\u007f]/g, "")
    .replace(/\s+/g, " ")
    .trim();

  return sanitized || "attachment";
}

/** Creates a unique attachment name for images inserted from the editor. */
export function createUniqueImageAttachmentName(fileName: string) {
  const safeName = sanitizeAttachmentFileName(fileName);
  return `teamedit-images/${crypto.randomUUID()}-${safeName}`;
}

/** Streams a browser `File` into the SDK attachment write stream. */
export async function uploadFileAttachment(
  database: MindooDBAppDatabase,
  docId: string,
  attachmentName: string,
  file: File,
) {
  const stream = await database.attachments.openWriteStream(
    docId,
    attachmentName,
    file.type || "application/octet-stream",
  );

  try {
    const bytes = new Uint8Array(await file.arrayBuffer());
    for (let offset = 0; offset < bytes.length; offset += DEFAULT_CHUNK_SIZE) {
      await stream.write(bytes.slice(offset, offset + DEFAULT_CHUNK_SIZE));
    }
    await stream.close();
  } catch (error) {
    await stream.abort();
    throw error;
  }
}

/** Reads an SDK attachment stream into a browser `Blob`. */
export async function readAttachmentBlob(
  database: MindooDBAppDatabase,
  docId: string,
  attachmentName: string,
  mimeType = "application/octet-stream",
) {
  const stream = await database.attachments.openReadStream(docId, attachmentName);
  const chunks: Uint8Array[] = [];
  try {
    for (;;) {
      const chunk = await stream.read();
      if (!chunk) {
        break;
      }
      chunks.push(chunk);
    }
  } finally {
    await stream.close();
  }
  return new Blob(chunks, { type: mimeType });
}

/** Extracts TeamEdit attachment image URLs from markdown so previews can preload them. */
export function extractAttachmentMarkdownUrls(markdown: string) {
  const urls = new Set<string>();
  const markdownImagePattern = /!\[[^\]]*]\(([^)\s]+)(?:\s+"[^"]*")?\)/g;
  let match: RegExpExecArray | null;
  while ((match = markdownImagePattern.exec(markdown)) !== null) {
    const url = match[1];
    if (url && isAttachmentMarkdownUrl(url)) {
      urls.add(url);
    }
  }
  return [...urls];
}

/** Human-readable file size formatting shared by the attachment panel. */
export function formatAttachmentSize(size: number) {
  if (!Number.isFinite(size) || size < 1024) {
    return `${Math.max(0, Math.round(size || 0))} B`;
  }
  const units = ["KB", "MB", "GB", "TB"];
  let value = size / 1024;
  let unitIndex = 0;
  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024;
    unitIndex += 1;
  }
  return `${value.toFixed(value >= 10 ? 0 : 1)} ${units[unitIndex]}`;
}
