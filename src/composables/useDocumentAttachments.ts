import { ref, type Ref } from "vue";
import {
  canPreviewAttachment,
  type MindooDBAppDatabase,
  type MindooDBAppDocument,
  type MindooDBAppDocumentRevisionId,
  type MindooDBAppHistoricalDocument,
} from "mindoodb-app-sdk";

import {
  formatAttachmentSize,
  readAttachmentBlob,
  sanitizeAttachmentFileName,
  uploadFileAttachment,
} from "@/lib/attachmentImages";

export interface UseDocumentAttachmentsOptions {
  database: Ref<MindooDBAppDatabase | null>;
  document: Readonly<
    Ref<MindooDBAppDocument | MindooDBAppHistoricalDocument | null>
  >;
  revisionId?: Ref<MindooDBAppDocumentRevisionId | null>;
  setStatus(message: string): void;
  onDocumentRefresh?(document: MindooDBAppDocument): void;
}

/**
 * Wraps the SDK attachment API for the active TeamEdit document.
 *
 * Apps upload and download binary data through streams because the SDK bridge
 * is message-based and should not require large files to fit in one payload.
 */
export function useDocumentAttachments(options: UseDocumentAttachmentsOptions) {
  const busyAction = ref<string | null>(null);
  const uploadInputKey = ref(0);

  async function refreshCurrentDocument() {
    const database = options.database.value;
    const document = options.document.value;
    if (!database || !document) {
      return;
    }
    const refreshed = await database.documents.get(document.id);
    if (refreshed) {
      options.onDocumentRefresh?.(refreshed);
    }
  }

  async function uploadAttachments(files: FileList | readonly File[] | null) {
    const database = options.database.value;
    const document = options.document.value;
    const selectedFiles = files ? Array.from(files) : [];
    if (!database || !document || selectedFiles.length === 0) {
      return;
    }

    busyAction.value = "Uploading attachment";
    try {
      for (const file of selectedFiles) {
        await uploadFileAttachment(
          database,
          document.id,
          sanitizeAttachmentFileName(file.name),
          file,
        );
      }
      await refreshCurrentDocument();
      uploadInputKey.value += 1;
      options.setStatus(
        `Uploaded ${selectedFiles.length} attachment${selectedFiles.length === 1 ? "" : "s"}.`,
      );
    } catch (error) {
      options.setStatus(
        error instanceof Error
          ? error.message
          : "The attachment upload failed.",
      );
    } finally {
      busyAction.value = null;
    }
  }

  async function scanAttachment() {
    const database = options.database.value;
    const document = options.document.value;
    if (!database || !document) {
      return;
    }

    busyAction.value = "Scanning attachment";
    try {
      const result = await database.attachments.scan(document.id, {
        defaultFileName: `scan-${document.id}.pdf`,
        preset: "a4-portrait",
        mimeType: "application/pdf",
      });
      if (result.ok) {
        await refreshCurrentDocument();
        uploadInputKey.value += 1;
        options.setStatus("Scanned document attached.");
      }
    } catch (error) {
      options.setStatus(
        error instanceof Error ? error.message : "The document scan failed.",
      );
    } finally {
      busyAction.value = null;
    }
  }

  async function previewAttachment(attachmentName: string) {
    const database = options.database.value;
    const document = options.document.value;
    if (!database || !document) {
      return;
    }

    busyAction.value = "Opening attachment preview";
    try {
      const revisionId = options.revisionId?.value ?? undefined;
      await database.attachments.openPreview(
        document.id,
        attachmentName,
        revisionId ? { revisionId } : undefined,
      );
    } catch (error) {
      options.setStatus(
        error instanceof Error
          ? error.message
          : "The attachment preview could not be opened.",
      );
    } finally {
      busyAction.value = null;
    }
  }

  async function downloadAttachment(attachmentName: string) {
    const database = options.database.value;
    const document = options.document.value;
    if (!database || !document || typeof window === "undefined") {
      return;
    }

    busyAction.value = "Downloading attachment";
    try {
      const attachment = document.attachments?.find(
        (entry) =>
          entry.fileName === attachmentName ||
          entry.attachmentId === attachmentName,
      );
      const revisionId = options.revisionId?.value ?? undefined;
      const blob = await readAttachmentBlob(
        database,
        document.id,
        attachmentName,
        attachment?.mimeType,
        revisionId ? { revisionId } : undefined,
      );
      const url = URL.createObjectURL(blob);
      const link = window.document.createElement("a");
      link.href = url;
      link.download = attachment?.fileName || attachmentName;
      link.click();
      URL.revokeObjectURL(url);
    } catch (error) {
      options.setStatus(
        error instanceof Error
          ? error.message
          : "The attachment could not be downloaded.",
      );
    } finally {
      busyAction.value = null;
    }
  }

  async function removeAttachment(attachmentName: string) {
    const database = options.database.value;
    const document = options.document.value;
    if (!database || !document) {
      return;
    }

    busyAction.value = "Removing attachment";
    try {
      await database.attachments.remove(document.id, attachmentName);
      await refreshCurrentDocument();
      uploadInputKey.value += 1;
      options.setStatus(`Removed attachment ${attachmentName}.`);
    } catch (error) {
      options.setStatus(
        error instanceof Error
          ? error.message
          : "The attachment could not be removed.",
      );
    } finally {
      busyAction.value = null;
    }
  }

  return {
    busyAction,
    canPreviewAttachment,
    downloadAttachment,
    formatAttachmentSize,
    previewAttachment,
    refreshCurrentDocument,
    removeAttachment,
    scanAttachment,
    uploadAttachments,
    uploadInputKey,
  };
}
