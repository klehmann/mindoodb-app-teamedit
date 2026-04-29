import { onBeforeUnmount, ref, type Ref } from "vue";
import type {
  MindooDBAppAttachmentInfo,
  MindooDBAppDatabase,
  MindooDBAppDocument,
  MindooDBAppDocumentRevisionId,
  MindooDBAppHistoricalDocument,
} from "mindoodb-app-sdk";

import {
  extractAttachmentMarkdownUrls,
  isAttachmentMarkdownUrl,
  parseAttachmentMarkdownUrl,
  readAttachmentBlob,
} from "@/lib/attachmentImages";

export interface UseAttachmentImageResolverOptions {
  database: Ref<MindooDBAppDatabase | null>;
  document: Readonly<Ref<MindooDBAppDocument | MindooDBAppHistoricalDocument | null>>;
  revisionId?: Ref<MindooDBAppDocumentRevisionId | null>;
}

/**
 * Resolves stable TeamEdit attachment markdown URLs to session-local object URLs.
 *
 * This keeps the document body portable (`mindoodb-attachment:...`) while giving
 * Milkdown and the preview pane regular browser image URLs to render.
 */
export function useAttachmentImageResolver(options: UseAttachmentImageResolverOptions) {
  const revision = ref(0);
  const objectUrls = new Map<string, string>();
  const pendingLoads = new Map<string, Promise<string>>();

  function findAttachmentMimeType(attachmentName: string) {
    const attachments: readonly MindooDBAppAttachmentInfo[] = options.document.value?.attachments ?? [];
    const attachment = attachments.find((entry) =>
      entry.fileName === attachmentName || entry.attachmentId === attachmentName);
    return attachment?.mimeType || "application/octet-stream";
  }

  function revokeCachedObjectUrl(attachmentName: string) {
    const objectUrl = objectUrls.get(attachmentName);
    if (objectUrl) {
      URL.revokeObjectURL(objectUrl);
      objectUrls.delete(attachmentName);
    }
  }

  async function loadAttachmentUrl(attachmentName: string) {
    const database = options.database.value;
    const document = options.document.value;
    if (!database || !document) {
      throw new Error("Open a document before resolving attachment images.");
    }

    const cached = objectUrls.get(attachmentName);
    if (cached) {
      return cached;
    }

    const pending = pendingLoads.get(attachmentName);
    if (pending) {
      return await pending;
    }

    const revisionId = options.revisionId?.value ?? undefined;
    const load = readAttachmentBlob(
      database,
      document.id,
      attachmentName,
      findAttachmentMimeType(attachmentName),
      revisionId ? { revisionId } : undefined,
    )
      .then((blob) => {
        revokeCachedObjectUrl(attachmentName);
        const objectUrl = URL.createObjectURL(blob);
        objectUrls.set(attachmentName, objectUrl);
        revision.value += 1;
        return objectUrl;
      })
      .finally(() => {
        pendingLoads.delete(attachmentName);
      });
    pendingLoads.set(attachmentName, load);
    return await load;
  }

  async function resolveImageUrl(url: string) {
    const attachmentName = parseAttachmentMarkdownUrl(url);
    return attachmentName ? await loadAttachmentUrl(attachmentName) : url;
  }

  function getCachedImageUrl(url: string) {
    const attachmentName = parseAttachmentMarkdownUrl(url);
    return attachmentName ? objectUrls.get(attachmentName) ?? url : url;
  }

  async function preloadMarkdownImages(markdown: string) {
    const urls = extractAttachmentMarkdownUrls(markdown);
    await Promise.all(urls.map((url) => resolveImageUrl(url).catch(() => url)));
  }

  function clear() {
    pendingLoads.clear();
    for (const objectUrl of objectUrls.values()) {
      URL.revokeObjectURL(objectUrl);
    }
    objectUrls.clear();
    revision.value += 1;
  }

  onBeforeUnmount(clear);

  return {
    revision,
    clear,
    getCachedImageUrl,
    isAttachmentImageUrl: isAttachmentMarkdownUrl,
    preloadMarkdownImages,
    resolveImageUrl,
  };
}
