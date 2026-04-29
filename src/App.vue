<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from "vue";
import Button from "primevue/button";
import Dialog from "primevue/dialog";
import Menubar from "primevue/menubar";
import Splitter from "primevue/splitter";
import SplitterPanel from "primevue/splitterpanel";
import type { MenuItem } from "primevue/menuitem";
import {
  createMindooDBAppBridge,
  createMindooDBTextBuffer,
  type MindooDBAppDatabase,
  type MindooDBAppDatabaseInfo,
  type MindooDBAppDocument,
  type MindooDBAppDocumentHistoryEntry,
  type MindooDBAppDocumentRevisionId,
  type MindooDBAppDocumentSummary,
  type MindooDBAppHistoricalDocument,
  type MindooDBAppRuntime,
  type MindooDBAppSession,
  type MindooDBAppUiPreferences,
  type MindooDBTextBuffer,
} from "mindoodb-app-sdk";

import AttachmentPickerDialog from "@/components/AttachmentPickerDialog.vue";
import DocumentAttachmentsPanel from "@/components/DocumentAttachmentsPanel.vue";
import DocumentRevisionDialog from "@/components/DocumentRevisionDialog.vue";
import MilkdownMarkdownEditor from "@/components/MilkdownMarkdownEditor.vue";
import { useAttachmentImageResolver } from "@/composables/useAttachmentImageResolver";
import { useDocumentAttachments } from "@/composables/useDocumentAttachments";
import {
  type AttachmentInsertion,
  createAttachmentMarkdownUrl,
  createUniqueImageAttachmentName,
  uploadFileAttachment,
} from "@/lib/attachmentImages";
import { exportMarkdownFile, exportMarkdownPackage } from "@/lib/exportMarkdown";
import { applyImageRatios } from "@/lib/imageRatio";
import { createMarkdownRenderer, normalizeMarkdownForRendering } from "@/lib/markdownRendering";
import { renderMermaidPlaceholders } from "@/lib/mermaid";
import { renderAndPrintMarkdownWindow } from "@/lib/printMarkdown";
import { applyAppTheme } from "@/lib/theme";

const PREVIEW_PANE_SETTINGS_KEY = "mindoodb-teamedit-preview-pane";

type PreviewPanePosition = "right" | "bottom";

function readPreviewPaneSettings() {
  if (typeof localStorage === "undefined") {
    return { showPreviewPane: true, previewPanePosition: "right" as PreviewPanePosition };
  }
  try {
    const parsed = JSON.parse(localStorage.getItem(PREVIEW_PANE_SETTINGS_KEY) ?? "{}") as {
      showPreviewPane?: unknown;
      previewPanePosition?: unknown;
    };
    return {
      showPreviewPane: typeof parsed.showPreviewPane === "boolean" ? parsed.showPreviewPane : true,
      previewPanePosition: parsed.previewPanePosition === "bottom" ? "bottom" as const : "right" as const,
    };
  } catch {
    return { showPreviewPane: true, previewPanePosition: "right" as PreviewPanePosition };
  }
}

const previewPaneSettings = readPreviewPaneSettings();

// Bridge/session state is kept in this root component because TeamEdit is a
// small sample app with one active document at a time.
const session = ref<MindooDBAppSession | null>(null);
const databases = ref<MindooDBAppDatabaseInfo[]>([]);
const selectedDatabaseId = ref("");
const documents = ref<MindooDBAppDocumentSummary[]>([]);
const currentDatabase = ref<MindooDBAppDatabase | null>(null);
const currentDatabaseId = ref("");
const currentDocument = ref<MindooDBAppDocument | null>(null);
const viewingHistoricalSnapshot = ref<MindooDBAppHistoricalDocument | null>(null);
const currentRuntime = ref<MindooDBAppRuntime>("iframe");
const hostUiPreferences = ref<MindooDBAppUiPreferences>({
  iosMultitaskingOptimized: false,
});

// The editor always works with a local markdown string. The text buffer mirrors
// that string, records local edits, and flushes granular text patches on save.
const textBuffer = ref<MindooDBTextBuffer | null>(null);
const markdown = ref("");
const subject = ref("");
const savedSubject = ref("");
const isDirty = ref(false);
const status = ref("Connecting to Haven...");
const openDialogVisible = ref(false);
const refreshConfirmVisible = ref(false);
const infoDialogVisible = ref(false);
const deleteConfirmVisible = ref(false);
const attachmentPickerVisible = ref(false);
const revisionDialogVisible = ref(false);
const revisionEntries = ref<MindooDBAppDocumentHistoryEntry[]>([]);
const revisionLoading = ref(false);
const revisionErrorMessage = ref<string | null>(null);
// The slash-menu callback hands us a promise resolver that we keep around
// while the attachment picker dialog is open. The resolver is invoked from
// either the picker's `select` or `cancel` paths, never both, and is cleared
// once it fires so a closed-and-reopened dialog can never resolve a stale
// editor request.
let attachmentPickerResolver: ((value: AttachmentInsertion | null) => void) | null = null;
const showPreviewPane = ref(previewPaneSettings.showPreviewPane);
const previewPanePosition = ref<PreviewPanePosition>(previewPaneSettings.previewPanePosition);
const previewRoot = ref<HTMLElement | null>(null);
const selectedOpenDocId = ref("");
const copiedInfoLabel = ref<string | null>(null);
let cleanupTheme: (() => void) | null = null;
let cleanupUiPreferences: (() => void) | null = null;
let previewRenderGeneration = 0;

// Programmatic updates from load/reconcile must not be recorded as fresh local
// edits, otherwise opening a document or accepting merged content would mark it
// dirty immediately.
let suppressEditorUpdate = false;

const readableDatabases = computed(() => databases.value.filter((database) => database.capabilities.includes("read")));
const creatableDatabases = computed(() => databases.value.filter((database) => database.capabilities.includes("create")));
const selectedDatabaseInfo = computed(() => databases.value.find((database) => database.id === selectedDatabaseId.value) ?? null);
const currentDatabaseInfo = computed(() => databases.value.find((database) => database.id === currentDatabaseId.value) ?? null);
const currentCanUpdate = computed(() => currentDatabaseInfo.value?.capabilities.includes("update") ?? false);
const currentCanDelete = computed(() => currentDatabaseInfo.value?.capabilities.includes("delete") ?? false);
const currentCanBrowseHistory = computed(() => currentDatabaseInfo.value?.capabilities.includes("history") ?? false);
const isViewingHistorical = computed(() => viewingHistoricalSnapshot.value !== null);
const activeDocumentView = computed(() => viewingHistoricalSnapshot.value ?? currentDocument.value);
const activeRevisionId = computed<MindooDBAppDocumentRevisionId | null>(() => viewingHistoricalSnapshot.value?.revisionId ?? null);
const canUseAttachments = computed(() => currentDatabaseInfo.value?.capabilities.includes("attachments") ?? false);
const canManageAttachments = computed(() => canUseAttachments.value && currentCanUpdate.value && !isViewingHistorical.value);
const canCreate = computed(() => creatableDatabases.value.length > 0);
const subjectDirty = computed(() => subject.value !== savedSubject.value);
const hasLocalEdits = computed(() => isDirty.value || subjectDirty.value);
const canSave = computed(() => Boolean(!isViewingHistorical.value && hasLocalEdits.value && currentCanUpdate.value && currentDatabase.value && currentDocument.value));
const canDelete = computed(() => Boolean(!isViewingHistorical.value && currentCanDelete.value && currentDatabase.value && currentDocument.value));
const canRefresh = computed(() => Boolean(currentDatabase.value && currentDocument.value));
const canShowInfo = computed(() => Boolean(currentDatabaseId.value && currentDocument.value));
const canPrint = computed(() => Boolean(currentDocument.value));
const canExportMarkdown = computed(() => Boolean(currentDocument.value));
const activeAttachments = computed(() => activeDocumentView.value?.attachments ?? []);
const hasAttachments = computed(() => activeAttachments.value.length > 0);
const canExportMarkdownWithAttachments = computed(() =>
  Boolean(currentDatabase.value && currentDocument.value && hasAttachments.value));
const currentRevisionId = computed(() => {
  if (viewingHistoricalSnapshot.value?.revisionId) {
    return viewingHistoricalSnapshot.value.revisionId;
  }
  return revisionEntries.value.find((entry) => entry.isCurrent)?.revisionId ?? null;
});
const statusBadgeLabel = computed(() => {
  if (viewingHistoricalSnapshot.value) {
    return `Historical · ${formatRevisionDate(viewingHistoricalSnapshot.value.timestamp)}`;
  }
  return `Current · ${hasLocalEdits.value ? "Unsaved" : "Saved"}`;
});
const statusBadgeTooltip = computed(() => isViewingHistorical.value
  ? "You're viewing a historical revision. Click to pick a different version or return to the current one."
  : "You're viewing the current version. Click to browse older revisions.");
const splitterLayout = computed(() => previewPanePosition.value === "bottom" ? "vertical" : "horizontal");
const currentDatabaseLabel = computed(() => {
  const info = databases.value.find((database) => database.id === currentDatabaseId.value);
  if (info?.title) {
    return `${info.title} (${currentDatabaseId.value})`;
  }
  return currentDatabaseId.value || "-";
});
function menuCheckIcon(checked: boolean) {
  return checked ? "pi pi-check" : "pi pi-check menu-check-placeholder";
}

const imageResolver = useAttachmentImageResolver({
  database: currentDatabase,
  document: activeDocumentView,
  revisionId: activeRevisionId,
});

const {
  busyAction: attachmentBusyAction,
  canPreviewAttachment,
  downloadAttachment,
  formatAttachmentSize,
  previewAttachment,
  removeAttachment,
  uploadAttachments,
  uploadInputKey: attachmentUploadInputKey,
} = useDocumentAttachments({
  database: currentDatabase,
  document: activeDocumentView,
  revisionId: activeRevisionId,
  runtime: currentRuntime,
  onDocumentRefresh(document) {
    currentDocument.value = document;
  },
  setStatus(message) {
    status.value = message;
  },
});

const markdownIt = createMarkdownRenderer({
  resolveImageUrl: (url) => imageResolver.getCachedImageUrl(url),
});

const renderedMarkdown = computed(() => {
  imageResolver.revision.value;
  return markdownIt.render(normalizeMarkdownForRendering(markdown.value || ""));
});

const menuItems = computed<MenuItem[]>(() => [
  {
    label: "File",
    icon: "pi pi-file",
    items: [
      {
        label: "New",
        icon: "pi pi-file-plus",
        disabled: !canCreate.value,
        command: () => {
          void newFile();
        },
      },
      {
        label: "Open",
        icon: "pi pi-folder-open",
        disabled: readableDatabases.value.length === 0,
        command: () => {
          void openFileDialog();
        },
      },
      {
        label: "Save",
        icon: "pi pi-save",
        disabled: !canSave.value,
        command: () => {
          void saveFile();
        },
      },
      {
        label: "Info",
        icon: "pi pi-info-circle",
        disabled: !canShowInfo.value,
        command: () => {
          copiedInfoLabel.value = null;
          infoDialogVisible.value = true;
        },
      },
      {
        label: "Print",
        icon: "pi pi-print",
        disabled: !canPrint.value,
        command: () => {
          void printCurrentDocument();
        },
      },
      { separator: true },
      {
        label: "Delete",
        icon: "pi pi-trash",
        disabled: !canDelete.value,
        command: () => {
          deleteConfirmVisible.value = true;
        },
      },
      { separator: true },
      {
        label: "Export Markdown",
        icon: "pi pi-file-export",
        disabled: !canExportMarkdown.value,
        command: () => {
          void exportCurrentMarkdown();
        },
      },
      {
        label: "Export Markdown with attachments",
        icon: "pi pi-file-export",
        visible: hasAttachments.value,
        disabled: !canExportMarkdownWithAttachments.value,
        command: () => {
          void exportCurrentMarkdownWithAttachments();
        },
      },
    ],
  },
  {
    label: "Display",
    icon: "pi pi-desktop",
    items: [
      {
        label: "Show preview pane",
        icon: menuCheckIcon(showPreviewPane.value),
        command: () => {
          showPreviewPane.value = !showPreviewPane.value;
        },
      },
      {
        label: "Preview pane",
        icon: "pi pi-window-maximize",
        items: [
          {
            label: "Right",
            icon: menuCheckIcon(previewPanePosition.value === "right"),
            disabled: !showPreviewPane.value,
            command: () => {
              previewPanePosition.value = "right";
            },
          },
          {
            label: "Bottom",
            icon: menuCheckIcon(previewPanePosition.value === "bottom"),
            disabled: !showPreviewPane.value,
            command: () => {
              previewPanePosition.value = "bottom";
            },
          },
        ],
      },
    ],
  },
]);

/**
 * Connect to the Haven host bridge, apply the host theme, and choose the first
 * writable database as the default target for File/New and File/Open.
 */
onMounted(async () => {
  try {
    const bridge = createMindooDBAppBridge();
    const nextSession = await bridge.connect();
    session.value = nextSession;
    const context = await nextSession.getLaunchContext();
    applyAppTheme(context.theme);
    cleanupTheme = nextSession.onThemeChange((theme) => applyAppTheme(theme));
    currentRuntime.value = context.runtime;
    hostUiPreferences.value = { ...context.uiPreferences };
    cleanupUiPreferences = nextSession.onUiPreferencesChange((uiPreferences) => {
      hostUiPreferences.value = { ...uiPreferences };
    });
    databases.value = context.databases;
    selectedDatabaseId.value = context.preferredDatabaseId
      ?? readableDatabases.value[0]?.id
      ?? context.databases[0]?.id
      ?? "";
    status.value = "Connected. Choose File / New or File / Open.";
  } catch (error) {
    status.value = error instanceof Error ? error.message : String(error);
  }
});

onBeforeUnmount(async () => {
  cleanupTheme?.();
  cleanupUiPreferences?.();
  await session.value?.disconnect();
});

watch(
  () => [markdown.value, currentDatabaseId.value, currentDocument.value?.id, imageResolver.revision.value] as const,
  () => {
    void imageResolver.preloadMarkdownImages(markdown.value);
  },
  { immediate: true },
);

async function renderPreviewMermaidDiagrams() {
  const generation = ++previewRenderGeneration;
  await nextTick();
  if (generation !== previewRenderGeneration || !previewRoot.value) {
    return;
  }
  await renderMermaidPlaceholders(previewRoot.value, {
    idPrefix: "teamedit-preview-mermaid",
  });
}

function applyPreviewImageRatios() {
  if (previewRoot.value) {
    applyImageRatios(previewRoot.value);
  }
}

watch(
  () => [renderedMarkdown.value, showPreviewPane.value] as const,
  () => {
    void renderPreviewMermaidDiagrams();
    // Run after markdown-it has finished writing into v-html so we can
    // measure container width for the ratio computation.
    void nextTick(applyPreviewImageRatios);
  },
  { flush: "post", immediate: true },
);

// The preview reuses object URLs created by the image resolver. When those
// resolve later we re-rerun the scaling so freshly-decoded images pick up
// their saved ratio without waiting for the next markdown change.
watch(
  () => imageResolver.revision.value,
  () => {
    void nextTick(applyPreviewImageRatios);
  },
);

watch(
  [showPreviewPane, previewPanePosition],
  ([nextShowPreviewPane, nextPreviewPanePosition]) => {
    if (typeof localStorage === "undefined") {
      return;
    }
    localStorage.setItem(PREVIEW_PANE_SETTINGS_KEY, JSON.stringify({
      showPreviewPane: nextShowPreviewPane,
      previewPanePosition: nextPreviewPanePosition,
    }));
  },
);

async function openDatabaseById(databaseId: string) {
  if (!session.value || !databaseId) {
    throw new Error("Select a database first.");
  }
  return await session.value.openDatabase(databaseId);
}

/** Refresh the document list shown in the File/Open dialog. */
async function refreshDocuments() {
  const database = await openDatabaseById(selectedDatabaseId.value);
  const result = await database.documents.list({
    status: "existing",
    limit: 100,
  });
  documents.value = [...result.items].sort((left, right) => {
    const subjectComparison = readDocumentSummaryLabel(left).localeCompare(readDocumentSummaryLabel(right), undefined, {
      sensitivity: "base",
      numeric: true,
    });
    return subjectComparison || left.id.localeCompare(right.id);
  });
}

/** Load documents before showing the File/Open picker. */
async function openFileDialog() {
  try {
    await refreshDocuments();
    selectedOpenDocId.value = documents.value[0]?.id ?? "";
    openDialogVisible.value = true;
  } catch (error) {
    status.value = error instanceof Error ? error.message : String(error);
  }
}

/** Create a new empty markdown document. */
async function newFile() {
  try {
    const targetDatabaseInfo = selectedDatabaseInfo.value?.capabilities.includes("create")
      ? selectedDatabaseInfo.value
      : creatableDatabases.value[0];
    if (!targetDatabaseInfo) {
      throw new Error("No writable database is available.");
    }
    selectedDatabaseId.value = targetDatabaseInfo.id;
    const database = await openDatabaseById(targetDatabaseInfo.id);
    const document = await database.documents.create({
      set: {
        subject: "",
        body: "",
      },
    });
    loadDocumentIntoEditor(database, targetDatabaseInfo.id, document);
    status.value = `Created ${document.id}.`;
  } catch (error) {
    status.value = error instanceof Error ? error.message : String(error);
  }
}

/** Open the selected document and initialize the editor/text buffer. */
async function openSelectedDocument() {
  try {
    const databaseId = selectedDatabaseId.value;
    const database = await openDatabaseById(databaseId);
    const document = selectedOpenDocId.value
      ? await database.documents.get(selectedOpenDocId.value)
      : null;
    if (!document) {
      throw new Error("Select a document to open.");
    }
    loadDocumentIntoEditor(database, databaseId, document);
    openDialogVisible.value = false;
    status.value = `Opened ${document.id}.`;
  } catch (error) {
    status.value = error instanceof Error ? error.message : String(error);
  }
}

/** Ask before discarding local edits, otherwise refresh immediately. */
function requestRefreshCurrentDocument() {
  if (isViewingHistorical.value) {
    returnToCurrent();
    return;
  }
  if (!canRefresh.value) {
    status.value = "Open a document before refreshing.";
    return;
  }
  if (hasLocalEdits.value) {
    refreshConfirmVisible.value = true;
    return;
  }
  void refreshCurrentDocument();
}

async function openRevisionDialog() {
  if (!currentDatabase.value || !currentDocument.value || !currentCanBrowseHistory.value) {
    return;
  }
  revisionDialogVisible.value = true;
  revisionLoading.value = true;
  revisionErrorMessage.value = null;
  try {
    revisionEntries.value = await currentDatabase.value.documents.listHistory(currentDocument.value.id);
  } catch (error) {
    revisionErrorMessage.value = error instanceof Error ? error.message : "The revision list could not be loaded.";
  } finally {
    revisionLoading.value = false;
  }
}

async function loadHistoricalRevision(revisionId: MindooDBAppDocumentRevisionId) {
  if (!currentDatabase.value || !currentDocument.value) {
    status.value = "Open a document before loading revisions.";
    return;
  }
  if (revisionEntries.value.find((entry) => entry.revisionId === revisionId)?.isCurrent) {
    revisionDialogVisible.value = false;
    returnToCurrent();
    return;
  }
  try {
    const snapshot = await currentDatabase.value.documents.getAtRevision(currentDocument.value.id, revisionId);
    if (snapshot.state !== "exists" || !snapshot.data) {
      status.value = snapshot.state === "deleted"
        ? "That revision is a deletion marker and cannot be opened in the editor."
        : "That revision is no longer available.";
      return;
    }
    viewingHistoricalSnapshot.value = snapshot;
    textBuffer.value = null;
    savedSubject.value = readHistoricalSubject(snapshot);
    subject.value = savedSubject.value;
    suppressEditorUpdate = true;
    markdown.value = readHistoricalBody(snapshot);
    isDirty.value = false;
    imageResolver.clear();
    revisionDialogVisible.value = false;
    queueMicrotask(() => {
      suppressEditorUpdate = false;
    });
    status.value = `Loaded revision from ${formatRevisionDate(snapshot.timestamp)}.`;
  } catch (error) {
    status.value = error instanceof Error ? error.message : "The revision could not be loaded.";
  }
}

function returnToCurrent() {
  if (!currentDocument.value) {
    return;
  }
  viewingHistoricalSnapshot.value = null;
  imageResolver.clear();
  textBuffer.value = currentDatabase.value
    ? createMindooDBTextBuffer({
        database: currentDatabase.value,
        document: currentDocument.value,
        path: ["body"],
      })
    : null;
  savedSubject.value = readSubject(currentDocument.value);
  subject.value = savedSubject.value;
  suppressEditorUpdate = true;
  markdown.value = textBuffer.value?.value ?? readDocumentBody(currentDocument.value);
  isDirty.value = false;
  queueMicrotask(() => {
    suppressEditorUpdate = false;
  });
  status.value = "Returned to the current version.";
}

/** Re-read the open document from Haven and discard unsaved local edits. */
async function refreshCurrentDocument() {
  if (!currentDatabase.value || !currentDocument.value) {
    status.value = "Open a document before refreshing.";
    return;
  }
  try {
    const document = await currentDatabase.value.documents.get(currentDocument.value.id);
    if (!document) {
      throw new Error("The current document could not be loaded.");
    }
    refreshConfirmVisible.value = false;
    loadDocumentIntoEditor(currentDatabase.value, currentDatabaseId.value, document);
    status.value = "Refreshed from Haven.";
  } catch (error) {
    status.value = error instanceof Error ? error.message : String(error);
  }
}

function copyWithTextareaFallback(value: string) {
  if (typeof document === "undefined") {
    return false;
  }

  const textarea = document.createElement("textarea");
  const activeElement = document.activeElement instanceof HTMLElement ? document.activeElement : null;
  textarea.value = value;
  textarea.setAttribute("readonly", "true");
  textarea.setAttribute("aria-hidden", "true");
  textarea.style.position = "fixed";
  textarea.style.top = "-9999px";
  textarea.style.left = "-9999px";
  textarea.style.opacity = "0";

  document.body.append(textarea);
  textarea.focus();
  textarea.select();
  textarea.setSelectionRange(0, textarea.value.length);

  try {
    return document.execCommand("copy");
  } finally {
    textarea.remove();
    activeElement?.focus();
  }
}

async function copyInfoValue(value: string, label: string) {
  if (!value) {
    return;
  }
  try {
    if (typeof navigator !== "undefined" && navigator.clipboard) {
      await navigator.clipboard.writeText(value);
      copiedInfoLabel.value = `${label} copied.`;
      return;
    }
  } catch {
    // Fall through to the legacy copy path for embedded runtimes.
  }

  copiedInfoLabel.value = copyWithTextareaFallback(value)
    ? `${label} copied.`
    : `${label} could not be copied.`;
}

async function printCurrentDocument() {
  if (!currentDocument.value) {
    status.value = "Open a document before printing.";
    return;
  }

  const printWindow = window.open("", "_blank");
  if (!printWindow) {
    status.value = "Pop-up blocked. Allow pop-ups for this app to print the document.";
    return;
  }

  try {
    status.value = "Preparing print view...";
    await imageResolver.preloadMarkdownImages(markdown.value);
    await renderAndPrintMarkdownWindow(printWindow, {
      title: subject.value || "Untitled document",
      markdown: markdown.value,
      resolveImageUrl: (url) => imageResolver.getCachedImageUrl(url),
    });
    status.value = "Print view opened.";
  } catch (error) {
    printWindow.close();
    status.value = error instanceof Error ? error.message : String(error);
  }
}

/** Export only the markdown text, preserving TeamEdit's stable attachment URLs. */
async function exportCurrentMarkdown() {
  if (!currentDocument.value) {
    status.value = "Open a document before exporting.";
    return;
  }

  try {
    const saved = await exportMarkdownFile(markdown.value, subject.value || "Untitled document");
    status.value = saved ? "Exported markdown file." : "Markdown export cancelled.";
  } catch (error) {
    status.value = error instanceof Error ? error.message : "The markdown export failed.";
  }
}

/**
 * Export a portable ZIP package.
 *
 * The package rewrites `mindoodb-attachment:` markdown image URLs to relative
 * `attachments/...` paths and includes all current MindooDoc attachments. This
 * keeps the exported markdown useful outside Haven while still preserving the
 * simple markdown-only export above for people who want the raw document body.
 */
async function exportCurrentMarkdownWithAttachments() {
  if (!currentDatabase.value || !currentDocument.value) {
    status.value = "Open a document before exporting attachments.";
    return;
  }

  const attachments = activeAttachments.value;
  if (attachments.length === 0) {
    status.value = "This document has no attachments to export.";
    return;
  }

  try {
    status.value = "Preparing markdown export package...";
    const saved = await exportMarkdownPackage({
      database: currentDatabase.value,
      documentId: currentDocument.value.id,
      markdown: markdown.value,
      title: subject.value || "Untitled document",
      attachments,
      revisionId: activeRevisionId.value ?? undefined,
    });
    status.value = saved ? "Exported markdown package." : "Markdown package export cancelled.";
  } catch (error) {
    status.value = error instanceof Error ? error.message : "The markdown package export failed.";
  }
}

/**
 * Flush buffered local text edits to Haven.
 *
 * Haven applies the text patch against the document heads captured by the
 * buffer. If another app or sync source changed the same document meanwhile,
 * the returned canonical body may differ from our local markdown string; in
 * that case we update Milkdown once with the reconciled content.
 */
async function saveFile() {
  if (!textBuffer.value || !currentDatabase.value || !currentDocument.value) {
    status.value = "Open or create a document before saving.";
    return;
  }
  if (!currentCanUpdate.value) {
    status.value = "This application does not have write access to the current document database.";
    return;
  }
  if (isViewingHistorical.value) {
    status.value = "Historical revisions are read-only. Return to the current version before saving.";
    return;
  }
  try {
    let document = currentDocument.value;
    if (subjectDirty.value) {
      document = await currentDatabase.value.documents.update(document.id, {
        set: { subject: subject.value },
      });
      currentDocument.value = document;
    }

    const result = textBuffer.value.dirty
      ? await textBuffer.value.flush()
      : { document, value: markdown.value, reconciled: false };
    currentDocument.value = result.document;
    savedSubject.value = readSubject(result.document);
    subject.value = savedSubject.value;
    if (result.reconciled) {
      suppressEditorUpdate = true;
      markdown.value = result.value;
      queueMicrotask(() => {
        suppressEditorUpdate = false;
      });
      status.value = "Saved and reconciled concurrent edits.";
    } else {
      status.value = "Saved.";
    }
    isDirty.value = false;
  } catch (error) {
    status.value = error instanceof Error ? error.message : String(error);
  }
}

async function deleteCurrentDocument() {
  if (!currentDatabase.value || !currentDocument.value) {
    status.value = "Open a document before deleting.";
    return;
  }
  if (!currentCanDelete.value) {
    status.value = "This application does not have delete access to the current document database.";
    return;
  }
  if (isViewingHistorical.value) {
    status.value = "Historical revisions are read-only. Return to the current version before deleting.";
    return;
  }
  try {
    const deletedDocumentId = currentDocument.value.id;
    await currentDatabase.value.documents.delete(deletedDocumentId);
    deleteConfirmVisible.value = false;
    currentDocument.value = null;
    textBuffer.value = null;
    imageResolver.clear();
    markdown.value = "";
    subject.value = "";
    savedSubject.value = "";
    isDirty.value = false;
    status.value = `Deleted ${deletedDocumentId}.`;
  } catch (error) {
    status.value = error instanceof Error ? error.message : String(error);
  }
}

/**
 * Uploads an image selected inside Milkdown and returns the stable markdown URL
 * that should be stored in the collaborative `body` text.
 */
async function uploadEditorImage(file: File) {
  if (!currentDatabase.value || !currentDocument.value) {
    throw new Error("Open or create a document before inserting images.");
  }
  if (!canManageAttachments.value) {
    throw new Error("This application does not have attachment upload access for the current document database.");
  }

  const attachmentName = createUniqueImageAttachmentName(file.name);
  await uploadFileAttachment(currentDatabase.value, currentDocument.value.id, attachmentName, file);
  const refreshed = await currentDatabase.value.documents.get(currentDocument.value.id);
  if (refreshed) {
    currentDocument.value = refreshed;
  }
  const markdownUrl = createAttachmentMarkdownUrl(attachmentName);
  void imageResolver.resolveImageUrl(markdownUrl);
  status.value = `Uploaded image ${file.name}.`;
  return markdownUrl;
}

async function removeDocumentAttachment(attachmentName: string) {
  await removeAttachment(attachmentName);
  imageResolver.clear();
}

/**
 * Open the attachment picker dialog for the editor's slash menu.
 *
 * The editor awaits the returned promise: resolving with a selection inserts
 * the attachment as either a markdown image (`![alt](url)`) or a markdown link
 * (`[label](url)`); resolving with `null` leaves the document unchanged.
 */
function requestAttachmentInsertFromEditor(): Promise<AttachmentInsertion | null> {
  if (!currentDocument.value) {
    status.value = "Open a document before inserting attachments.";
    return Promise.resolve(null);
  }
  if (!canUseAttachments.value) {
    status.value = "This application does not have attachment access for the current document database.";
    return Promise.resolve(null);
  }
  if (isViewingHistorical.value) {
    status.value = "Return to the current version before inserting attachment links.";
    return Promise.resolve(null);
  }
  // If a previous request is still pending (e.g. user reopened the slash menu
  // before the dialog closed) we cancel it so each call resolves exactly once.
  attachmentPickerResolver?.(null);
  return new Promise<AttachmentInsertion | null>((resolve) => {
    attachmentPickerResolver = resolve;
    attachmentPickerVisible.value = true;
  });
}

function handleAttachmentPickerSelect(selection: AttachmentInsertion) {
  const resolver = attachmentPickerResolver;
  attachmentPickerResolver = null;
  resolver?.(selection);
  status.value = selection.isImage
    ? `Inserted image attachment ${selection.url.replace(/^mindoodb-attachment:/, "")}.`
    : `Inserted link to attachment ${selection.url.replace(/^mindoodb-attachment:/, "")}.`;
}

function handleAttachmentPickerCancel() {
  const resolver = attachmentPickerResolver;
  attachmentPickerResolver = null;
  resolver?.(null);
}

/**
 * Replace the active editor state with a document snapshot.
 *
 * This creates a new `MindooDBTextBuffer` bound to the document's `body` path.
 * The buffer stores the document heads from the snapshot so later saves can be
 * merged correctly with concurrent changes.
 */
function loadDocumentIntoEditor(database: MindooDBAppDatabase, databaseId: string, document: MindooDBAppDocument) {
  imageResolver.clear();
  currentDatabase.value = database;
  currentDatabaseId.value = databaseId;
  currentDocument.value = document;
  viewingHistoricalSnapshot.value = null;
  textBuffer.value = createMindooDBTextBuffer({
    database,
    document,
    path: ["body"],
  });
  subject.value = readSubject(document);
  savedSubject.value = subject.value;
  suppressEditorUpdate = true;
  markdown.value = textBuffer.value.value;
  isDirty.value = false;
  queueMicrotask(() => {
    suppressEditorUpdate = false;
  });
}

/**
 * Receive markdown updates from Milkdown.
 *
 * Milkdown emits full markdown strings rather than raw text splices, so the
 * buffer converts each new value to a minimal text splice with `replaceText()`.
 */
function onEditorUpdate(value: string) {
  if (isViewingHistorical.value) {
    return;
  }
  markdown.value = value;
  if (!suppressEditorUpdate) {
    textBuffer.value?.replaceText(value);
    isDirty.value = textBuffer.value?.dirty ?? false;
  }
}

function readSubject(document: MindooDBAppDocument) {
  const value = document.data.subject;
  return typeof value === "string" ? value : "";
}

function readDocumentBody(document: MindooDBAppDocument) {
  const value = document.data.body;
  return typeof value === "string" ? value : "";
}

function readHistoricalSubject(snapshot: MindooDBAppHistoricalDocument) {
  const value = snapshot.data?.subject;
  return typeof value === "string" ? value : "";
}

function readHistoricalBody(snapshot: MindooDBAppHistoricalDocument) {
  const value = snapshot.data?.body;
  return typeof value === "string" ? value : "";
}

function formatRevisionDate(timestamp: number) {
  return new Intl.DateTimeFormat(undefined, {
    dateStyle: "medium",
    timeStyle: "short",
  }).format(new Date(timestamp));
}

function readDocumentSummaryLabel(document: MindooDBAppDocumentSummary) {
  const value = document.data?.subject;
  return typeof value === "string" && value.trim() ? value : document.id;
}
</script>

<template>
  <main class="app-shell">
    <header
      class="toolbar glass-card"
      :class="{ 'toolbar--ios-multitasking': hostUiPreferences.iosMultitaskingOptimized }"
    >
      <div class="toolbar__leading">
        <span class="toolbar__title">TeamEdit</span>
        <Menubar :model="menuItems" class="toolbar__menubar" />
        <Button
          :icon="isViewingHistorical ? 'pi pi-history' : 'pi pi-refresh'"
          text
          rounded
          severity="secondary"
          class="toolbar__refresh"
          :aria-label="isViewingHistorical ? 'Return to current version' : 'Refresh document'"
          :title="isViewingHistorical ? 'Return to current version' : 'Refresh document'"
          :disabled="!canRefresh"
          @click="requestRefreshCurrentDocument"
        />
      </div>
      <div class="toolbar__meta">
        <button
          v-if="currentCanBrowseHistory && currentDocument"
          class="toolbar__status-badge"
          type="button"
          v-tooltip="statusBadgeTooltip"
          @click="openRevisionDialog"
        >
          {{ statusBadgeLabel }}
        </button>
        <span v-else class="toolbar__status-badge">{{ statusBadgeLabel }}</span>
      </div>
    </header>

    <section class="workspace">
      <section class="document-area">
        <Splitter
          v-if="showPreviewPane"
          :layout="splitterLayout"
          class="editor-preview-splitter"
        >
          <SplitterPanel :size="previewPanePosition === 'bottom' ? 60 : 62" :min-size="25" class="splitter-panel">
            <section class="editor-panel glass-card" :class="{ 'editor-panel--history': isViewingHistorical }">
              <template v-if="currentDocument">
                <div v-if="isViewingHistorical" class="history-banner">
                  <i class="pi pi-history" aria-hidden="true" />
                  <span>You're viewing the version from {{ formatRevisionDate(viewingHistoricalSnapshot?.timestamp ?? 0) }} - read-only.</span>
                  <button type="button" @click="returnToCurrent">Return to current</button>
                </div>
                <label class="field subject-field">
                  <span class="field-label">Title</span>
                  <input
                    v-model="subject"
                    class="native-input subject-input"
                    placeholder="Document title"
                    type="text"
                    :readonly="isViewingHistorical"
                  >
                </label>
                <MilkdownMarkdownEditor
                  :model-value="markdown"
                  :on-image-upload="uploadEditorImage"
                  :resolve-image-url="imageResolver.resolveImageUrl"
                  :request-attachment-insert="requestAttachmentInsertFromEditor"
                  :readonly="isViewingHistorical"
                  @update:model-value="onEditorUpdate"
                />
              </template>
              <div v-else class="empty-state">
                Use File / New or File / Open to start editing a markdown document.
              </div>
            </section>
          </SplitterPanel>

          <SplitterPanel :size="previewPanePosition === 'bottom' ? 40 : 38" :min-size="20" class="splitter-panel">
            <section class="preview-panel glass-card">
              <p class="eyebrow">Preview</p>
              <article ref="previewRoot" class="markdown-preview" v-html="renderedMarkdown" />
            </section>
          </SplitterPanel>
        </Splitter>

        <section v-else class="editor-panel glass-card" :class="{ 'editor-panel--history': isViewingHistorical }">
          <template v-if="currentDocument">
            <div v-if="isViewingHistorical" class="history-banner">
              <i class="pi pi-history" aria-hidden="true" />
              <span>You're viewing the version from {{ formatRevisionDate(viewingHistoricalSnapshot?.timestamp ?? 0) }} - read-only.</span>
              <button type="button" @click="returnToCurrent">Return to current</button>
            </div>
            <label class="field subject-field">
              <span class="field-label">Title</span>
              <input
                v-model="subject"
                class="native-input subject-input"
                placeholder="Document title"
                type="text"
                :readonly="isViewingHistorical"
              >
            </label>
            <MilkdownMarkdownEditor
              :model-value="markdown"
              :on-image-upload="uploadEditorImage"
              :resolve-image-url="imageResolver.resolveImageUrl"
              :request-attachment-insert="requestAttachmentInsertFromEditor"
              :readonly="isViewingHistorical"
              @update:model-value="onEditorUpdate"
            />
          </template>
          <div v-else class="empty-state">
            Use File / New or File / Open to start editing a markdown document.
          </div>
        </section>
      </section>

      <DocumentAttachmentsPanel
        v-if="currentDocument"
        :attachments="activeAttachments"
        :busy-action="attachmentBusyAction"
        :can-manage-attachments="canManageAttachments"
        :can-use-attachments="canUseAttachments"
        :historical="isViewingHistorical"
        :upload-input-key="attachmentUploadInputKey"
        :can-preview-attachment="canPreviewAttachment"
        :format-attachment-size="formatAttachmentSize"
        @upload="uploadAttachments"
        @preview="previewAttachment"
        @download="downloadAttachment"
        @remove="removeDocumentAttachment"
      />
    </section>

    <Dialog v-model:visible="openDialogVisible" modal header="Open Markdown Document" :style="{ width: '32rem' }">
      <div class="dialog-content">
        <label class="field">
          <span class="field-label">Database</span>
          <select v-model="selectedDatabaseId" class="native-input" @change="refreshDocuments">
            <option v-for="database in readableDatabases" :key="database.id" :value="database.id">
              {{ database.title || database.id }}
            </option>
          </select>
        </label>
        <label class="field">
          <span class="field-label">Document</span>
          <select v-model="selectedOpenDocId" class="native-input">
            <option v-for="document in documents" :key="document.id" :value="document.id">
              {{ readDocumentSummaryLabel(document) }}
            </option>
          </select>
        </label>
      </div>
      <template #footer>
        <Button label="Cancel" severity="secondary" @click="openDialogVisible = false" />
        <Button label="Open" :disabled="!selectedOpenDocId" @click="openSelectedDocument" />
      </template>
    </Dialog>

    <Dialog
      v-model:visible="refreshConfirmVisible"
      modal
      header="Discard local edits?"
      :style="{ width: '28rem', maxWidth: '96vw' }"
    >
      <p>
        Refreshing will reload this document from Haven and discard unsaved changes in this window.
      </p>
      <template #footer>
        <Button label="Cancel" text @click="refreshConfirmVisible = false" />
        <Button
          label="Refresh"
          icon="pi pi-refresh"
          severity="warn"
          @click="refreshCurrentDocument"
        />
      </template>
    </Dialog>

    <Dialog
      v-model:visible="deleteConfirmVisible"
      modal
      header="Delete document?"
      :style="{ width: '28rem', maxWidth: '96vw' }"
    >
      <p>
        Delete document <code>{{ currentDocument?.id }}</code>? This removes it from the current database.
      </p>
      <template #footer>
        <Button label="Cancel" text @click="deleteConfirmVisible = false" />
        <Button
          label="Delete"
          icon="pi pi-trash"
          severity="danger"
          :disabled="!canDelete"
          @click="deleteCurrentDocument"
        />
      </template>
    </Dialog>

    <Dialog
      v-model:visible="infoDialogVisible"
      modal
      header="Document info"
      :style="{ width: '32rem', maxWidth: '96vw' }"
      @hide="copiedInfoLabel = null"
    >
      <section class="info-stack">
        <div class="info-row">
          <div class="info-copy">
            <span class="field-label">Database</span>
            <code class="info-value">{{ currentDatabaseLabel }}</code>
          </div>
          <Button
            label="Copy"
            text
            size="small"
            :disabled="!currentDatabaseId"
            @click="copyInfoValue(currentDatabaseId, 'Database id')"
          />
        </div>

        <div class="info-row">
          <div class="info-copy">
            <span class="field-label">Document id</span>
            <code class="info-value">{{ currentDocument?.id ?? "-" }}</code>
          </div>
          <Button
            label="Copy"
            text
            size="small"
            :disabled="!currentDocument?.id"
            @click="copyInfoValue(currentDocument?.id ?? '', 'Document id')"
          />
        </div>
      </section>

      <p v-if="copiedInfoLabel" class="field-hint">{{ copiedInfoLabel }}</p>

      <template #footer>
        <Button label="Close" text @click="infoDialogVisible = false" />
      </template>
    </Dialog>

    <AttachmentPickerDialog
      v-model:visible="attachmentPickerVisible"
      :attachments="activeAttachments"
      :resolve-image-url="imageResolver.resolveImageUrl"
      @select="handleAttachmentPickerSelect"
      @cancel="handleAttachmentPickerCancel"
    />
    <DocumentRevisionDialog
      v-model:visible="revisionDialogVisible"
      :entries="revisionEntries"
      :loading="revisionLoading"
      :error-message="revisionErrorMessage"
      :current-revision-id="currentRevisionId"
      @select="loadHistoricalRevision"
      @cancel="revisionDialogVisible = false"
    />
  </main>
</template>

<style scoped>
.app-shell {
  min-height: 100vh;
  padding: 2px;
  display: grid;
  grid-template-rows: auto minmax(0, 1fr);
  gap: 2px;
}

.toolbar {
  position: sticky;
  top: 2px;
  z-index: 24;
  padding: 1px 0.55rem 1px 4px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 0.45rem;
  border-radius: 0;
  backdrop-filter: blur(26px);
  isolation: isolate;
}

.eyebrow,
.field-label {
  margin: 0;
  color: var(--muted);
  font-size: 0.78rem;
  text-transform: uppercase;
  letter-spacing: 0.08em;
}

.toolbar__leading {
  display: flex;
  align-items: center;
  flex: 1 1 auto;
  min-width: 0;
  gap: 0;
}

:deep(.toolbar__menubar) {
  position: relative;
  flex: 0 0 auto;
  border: 0;
  background: transparent;
  padding: 0 2px;
}

:deep(.toolbar__menubar.p-menubar) {
  min-height: 2rem;
}

.toolbar__refresh {
  flex: 0 0 auto;
  width: 1.9rem;
  height: 1.9rem;
  padding: 0;
}

.toolbar__meta {
  display: flex;
  align-items: center;
  gap: 0.35rem;
  flex: 0 0 auto;
  color: var(--muted);
  font-size: 0.78rem;
}

.toolbar__title {
  flex: 0 0 auto;
  margin: 0;
  font-weight: 700;
  color: var(--muted);
}

.toolbar__status-badge {
  padding: 0.12rem 0.42rem;
  border: 0;
  border-radius: 999px;
  background: rgb(255 255 255 / 0.06);
  color: inherit;
  font: inherit;
}

button.toolbar__status-badge {
  cursor: pointer;
}

button.toolbar__status-badge:hover,
button.toolbar__status-badge:focus-visible {
  background: rgb(255 255 255 / 0.12);
}

:deep(.toolbar__menubar .p-menubar-root-list) {
  gap: 0.15rem;
}

:deep(.toolbar__menubar .p-menubar-item-content) {
  border-radius: 0.6rem;
}

:deep(.toolbar__menubar .p-menubar-item-link) {
  padding: 0.35rem 0.55rem;
  gap: 0.35rem;
}

:deep(.toolbar__menubar .p-menubar-submenu .p-menubar-item-link) {
  padding: 0.5rem 0.75rem;
}

:deep(.toolbar__menubar .menu-check-placeholder) {
  visibility: hidden;
}

@media (max-width: 960px) {
  .toolbar--ios-multitasking .toolbar__title {
    margin-left: 80px;
  }
}

@media (max-width: 960px) {
  :deep(.toolbar__menubar.p-menubar-mobile .p-menubar-root-list) {
    position: absolute;
    top: calc(100% + 0.45rem);
    left: 0;
    z-index: 60;
    min-width: min(18rem, calc(100vw - 1rem));
    max-width: calc(100vw - 1rem);
    padding: 0.45rem;
    border: 1px solid var(--border);
    border-radius: 0.9rem;
    background: var(--bg-elevated);
    box-shadow: var(--shadow);
    backdrop-filter: var(--surface-blur);
  }
}

.workspace {
  min-height: 0;
  display: grid;
  grid-template-rows: minmax(0, 1fr) auto;
  gap: 2px;
}

.editor-panel,
.preview-panel {
  min-height: 0;
  padding: 1rem;
  border-radius: 0;
}

.document-area {
  min-width: 0;
  min-height: 0;
  height: 100%;
}

.editor-preview-splitter {
  min-width: 0;
  min-height: 0;
  height: 100%;
  border: 0;
  border-radius: 0;
  background: transparent;
}

:deep(.editor-preview-splitter .p-splitter-gutter) {
  background: transparent;
}

:deep(.editor-preview-splitter .p-splitter-gutter-handle) {
  background: var(--border);
}

.splitter-panel {
  min-width: 0;
  min-height: 0;
  overflow: hidden;
}

.field {
  display: grid;
  gap: 0.35rem;
}

.native-input {
  width: 100%;
  padding: 0.7rem 0.85rem;
  border: 1px solid var(--border);
  border-radius: 0.85rem;
  background: rgb(255 255 255 / 0.04);
  color: inherit;
}

.empty-state {
  color: var(--muted);
}

.editor-panel {
  height: 100%;
  overflow: auto;
  display: grid;
  grid-template-rows: auto minmax(0, 1fr);
  gap: 1rem;
}

.editor-panel--history {
  grid-template-rows: auto auto minmax(0, 1fr);
}

.history-banner {
  display: flex;
  align-items: center;
  gap: 0.55rem;
  padding: 0.55rem 0.7rem;
  border: 1px solid var(--border);
  border-radius: 0.75rem;
  background: rgb(255 255 255 / 0.04);
  color: var(--muted);
  font-size: 0.86rem;
}

.history-banner button {
  border: 0;
  background: transparent;
  color: var(--accent);
  cursor: pointer;
  font: inherit;
  text-decoration: underline;
  text-underline-offset: 0.15em;
}

.preview-panel {
  height: 100%;
  overflow: auto;
}

.subject-field {
  position: sticky;
  top: 0;
  z-index: 1;
  padding-bottom: 0.75rem;
  background: var(--surface);
}

.subject-input {
  font-size: 1.1rem;
  font-weight: 700;
}

.markdown-preview {
  line-height: 1.65;
}

.markdown-preview :deep(a) {
  color: var(--accent);
}

.markdown-preview :deep(.teamedit-mermaid-figure) {
  margin: 1rem 0;
}

.markdown-preview :deep(.teamedit-mermaid-placeholder) {
  display: grid;
  justify-items: center;
  min-height: 6rem;
  padding: 1rem;
  border: 1px solid var(--border);
  border-radius: var(--radius-md);
  background: rgb(255 255 255 / 0.04);
}

.markdown-preview :deep(.teamedit-mermaid-placeholder pre) {
  display: none;
}

.markdown-preview :deep(.teamedit-mermaid-placeholder svg) {
  max-width: 100%;
  height: auto;
}

.markdown-preview :deep(.teamedit-mermaid-message) {
  margin: 0;
  color: var(--muted);
}

.markdown-preview :deep(.teamedit-mermaid-message--error) {
  color: var(--danger);
}

.empty-state {
  min-height: 20rem;
  display: grid;
  place-items: center;
  text-align: center;
}

.dialog-content {
  display: grid;
  gap: 1rem;
}

.info-stack {
  display: grid;
  gap: 0.85rem;
}

.info-row {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 1rem;
  padding: 0.85rem 0.95rem;
  border: 1px solid rgb(255 255 255 / 0.08);
  border-radius: 0.9rem;
  background: rgb(255 255 255 / 0.03);
}

.info-copy {
  min-width: 0;
  display: grid;
  gap: 0.35rem;
}

.info-value {
  overflow-wrap: anywhere;
  white-space: pre-wrap;
}

.field-hint {
  margin: 0.85rem 0 0;
  color: var(--muted);
  font-size: 0.85rem;
}

@media (max-width: 1100px) {
  .workspace {
    grid-template-columns: 1fr;
  }
}
</style>
