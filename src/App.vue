<script setup lang="ts">
import {
  computed,
  nextTick,
  onBeforeUnmount,
  onMounted,
  ref,
  shallowRef,
  watch,
} from "vue";
import "highlight.js/styles/github.css";
import Button from "primevue/button";
import Dialog from "primevue/dialog";
import Menubar from "primevue/menubar";
import Message from "primevue/message";
import SplitButton from "primevue/splitbutton";
import Splitter from "primevue/splitter";
import SplitterPanel from "primevue/splitterpanel";
import type { MenuItem } from "primevue/menuitem";
import { parseDocx } from "@eigenpal/docx-editor-core/docx";
import type { Document as DocxDocument } from "@eigenpal/docx-editor-core/types/document";
import {
  createMindooDBAppBridge,
  createMindooDBTextBuffer,
  type MindooDBAppDatabase,
  type MindooDBAppDatabaseInfo,
  type MindooDBAppDocument,
  type MindooDBAppDocumentHistoryEntry,
  type MindooDBAppDocumentRevisionId,
  type MindooDBAppHistoricalDocument,
  type MindooDBAppRichTextSpan,
  type MindooDBAppRuntime,
  type MindooDBAppSession,
  type MindooDBAppUiPreferences,
  type MindooDBAppViewNavigator,
  type MindooDBTextBuffer,
} from "mindoodb-app-sdk";

import AttachmentPickerDialog from "@/components/AttachmentPickerDialog.vue";
import DocumentAttachmentsPanel from "@/components/DocumentAttachmentsPanel.vue";
import DocumentRevisionDialog from "@/components/DocumentRevisionDialog.vue";
import MilkdownMarkdownEditor from "@/components/MilkdownMarkdownEditor.vue";
import TagTreeList from "@/components/TagTreeList.vue";
import WordDocumentEditor from "@/editors/word/components/WordDocumentEditor.vue";
import { useAttachmentImageResolver } from "@/composables/useAttachmentImageResolver";
import { useDocumentAttachments } from "@/composables/useDocumentAttachments";
import {
  type AttachmentInsertion,
  createAttachmentMarkdownUrl,
  createUniqueImageAttachmentName,
  uploadFileAttachment,
} from "@/lib/attachmentImages";
import {
  normalizeTags,
  readTags as readDocumentTags,
} from "@/lib/documentTags";
import { exportDocxFile } from "@/lib/exportDocx";
import {
  createExportFileName,
  exportMarkdownFile,
  exportMarkdownPackage,
  saveBlobToDisk,
} from "@/lib/exportMarkdown";
import { applyImageRatios } from "@/lib/imageRatio";
import {
  createMarkdownRenderer,
  normalizeMarkdownForRendering,
} from "@/lib/markdownRendering";
import { renderMermaidPlaceholders } from "@/lib/mermaid";
import { renderAndPrintMarkdownWindow } from "@/lib/printMarkdown";
import { applyAppTheme } from "@/lib/theme";
import { useTeamEditAppUpdate } from "@/pwa/appUpdate";
import {
  ALL_DOCUMENTS_NODE_KEY,
  buildOpenCategoryTree,
  collectNavigatorEntries,
  createOpenViewDefinition,
  dedupeDocumentEntries,
  mapDocumentEntries,
  type OpenCategoryNode,
  type OpenDocumentRow,
  type OpenDocumentTemplateFilter,
} from "@/lib/viewOpen";
import {
  attachCommentsToDocument,
  commentsFromWordDocument,
  createDefaultWordDocument,
  documentToRichTextSpans,
  richTextSpansToDocument,
  summarizeWordDocument,
} from "@/editors/word/lib/richTextSpans";
import {
  loadWordDocumentFromAutomerge,
  loadWordAutomergeSnapshot,
  seedWordRichTextDocument,
  WORD_RICH_TEXT_PATH,
} from "@/editors/word/lib/wordRichTextSession";
import {
  createWordAutomergeHandle,
  type WordAutomergeHandle,
} from "@/editors/word/lib/wordAutomergeHandle";

const PREVIEW_PANE_SETTINGS_KEY = "mindoodb-teamedit-preview-pane";

type PreviewPanePosition = "right" | "bottom";
type DocumentType = "markdown" | "word";
type OpenDocumentTypeFilter = "all" | DocumentType;
type OpenDialogMode = "open" | "template";

interface OpenDocumentSession {
  id: string;
  databaseId: string;
  database: MindooDBAppDatabase;
  documentId: string;
  document: MindooDBAppDocument;
  type: DocumentType;
  textBuffer: MindooDBTextBuffer | null;
  markdown: string;
  wordEditorDocument: DocxDocument | null;
  currentWordDocument: DocxDocument | null;
  subject: string;
  savedSubject: string;
  tags: string[];
  savedTags: string[];
  isTemplate: boolean;
  savedIsTemplate: boolean;
  isDirty: boolean;
}

function readPreviewPaneSettings() {
  if (typeof localStorage === "undefined") {
    return {
      showPreviewPane: true,
      previewPanePosition: "right" as PreviewPanePosition,
    };
  }
  try {
    const parsed = JSON.parse(
      localStorage.getItem(PREVIEW_PANE_SETTINGS_KEY) ?? "{}",
    ) as {
      showPreviewPane?: unknown;
      previewPanePosition?: unknown;
    };
    return {
      showPreviewPane:
        typeof parsed.showPreviewPane === "boolean"
          ? parsed.showPreviewPane
          : true,
      previewPanePosition:
        parsed.previewPanePosition === "bottom"
          ? ("bottom" as const)
          : ("right" as const),
    };
  } catch {
    return {
      showPreviewPane: true,
      previewPanePosition: "right" as PreviewPanePosition,
    };
  }
}

const previewPaneSettings = readPreviewPaneSettings();
const { updateAvailable, updateReloading, reloadForUpdate } =
  useTeamEditAppUpdate();

// Bridge/session state is kept in this root component because TeamEdit is a
// small sample app. Open document sessions keep multiple files available while
// one editor surface is active, matching TeamGrid's Window menu model.
const session = ref<MindooDBAppSession | null>(null);
const currentUserName = ref("User");

/**
 * Convert a canonical (X.500) name such as `cn=Test/o=ACME` to its abbreviated
 * form (`Test/ACME`) by dropping the component type prefixes (`cn=`, `ou=`,
 * `o=`, `c=`, ...). Values without prefixes are passed through unchanged.
 */
function abbreviateUserName(name: string | null | undefined): string {
  if (!name) {
    return "";
  }
  return name
    .split("/")
    .map((component) => component.replace(/^\s*[a-z]+\s*=\s*/i, "").trim())
    .filter((component) => component.length > 0)
    .join("/");
}
const launchTimeTravelDate = ref<number | null>(null);
const databases = ref<MindooDBAppDatabaseInfo[]>([]);
const selectedDatabaseId = ref("");
const currentDatabase = ref<MindooDBAppDatabase | null>(null);
const currentDatabaseId = ref("");
const currentDocument = ref<MindooDBAppDocument | null>(null);
const currentDocumentType = ref<DocumentType>("markdown");
const currentWordDocument = shallowRef<DocxDocument | null>(null);
const wordEditorDocument = shallowRef<DocxDocument | null>(null);
const wordAutomergeHandle = shallowRef<WordAutomergeHandle | null>(null);
let wordAutomergeChangeListener: ((snapshot: { spans: MindooDBAppRichTextSpan[] }) => void) | null = null;
let snapshotActiveSessionTimer: ReturnType<typeof setTimeout> | null = null;
const SNAPSHOT_ACTIVE_SESSION_DEBOUNCE_MS = 300;
const appDocxImportInputRef = shallowRef<HTMLInputElement | null>(null);
const wordEditorRef = shallowRef<{
  openImportDialog: () => void;
  openPageSetup: () => Promise<boolean>;
  print: () => void;
  saveDocx: () => Promise<Blob | null>;
} | null>(null);
const viewingHistoricalSnapshot = ref<MindooDBAppHistoricalDocument | null>(
  null,
);
const currentRuntime = ref<MindooDBAppRuntime>("iframe");
const hostUiPreferences = ref<MindooDBAppUiPreferences>({
  iosMultitaskingOptimized: false,
  reduceMotion: false,
});

// The editor always works with a local markdown string. The text buffer mirrors
// that string, records local edits, and flushes granular text patches on save.
const textBuffer = shallowRef<MindooDBTextBuffer | null>(null);
const markdown = ref("");
const subject = ref("");
const savedSubject = ref("");
const tags = ref<string[]>([]);
const savedTags = ref<string[]>([]);
const isTemplate = ref(false);
const savedIsTemplate = ref(false);
const isDirty = ref(false);
const openDocumentSessions = shallowRef<OpenDocumentSession[]>([]);
const activeDocumentSessionId = ref("");
const status = ref("Connecting to Haven...");
const openDialogVisible = ref(false);
const propertiesDialogVisible = ref(false);
const refreshConfirmVisible = ref(false);
const infoDialogVisible = ref(false);
const deleteConfirmVisible = ref(false);
const closeConfirmVisible = ref(false);
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
let attachmentPickerResolver:
  | ((value: AttachmentInsertion | null) => void)
  | null = null;
const showPreviewPane = ref(previewPaneSettings.showPreviewPane);
const previewPanePosition = ref<PreviewPanePosition>(
  previewPaneSettings.previewPanePosition,
);
const previewRoot = ref<HTMLElement | null>(null);
const selectedOpenDocId = ref("");
const selectedOpenType = ref<OpenDocumentTypeFilter>("all");
const selectedOpenTemplateFilter = ref<OpenDocumentTemplateFilter>("noTemplates");
const openDialogMode = ref<OpenDialogMode>("open");
const selectedOpenCategoryKey = ref(ALL_DOCUMENTS_NODE_KEY);
const openCategoryNodes = ref<OpenCategoryNode[]>([]);
const openDialogDocuments = ref<OpenDocumentRow[]>([]);
const allOpenDialogDocuments = ref<OpenDocumentRow[]>([]);
const openNavigator = ref<MindooDBAppViewNavigator | null>(null);
const copiedInfoLabel = ref<string | null>(null);
const propertiesTitleDraft = ref("");
const propertiesTagsDraft = ref("");
const propertiesIsTemplateDraft = ref(false);
const pendingCloseSessionId = ref("");
let cleanupTheme: (() => void) | null = null;
let cleanupUiPreferences: (() => void) | null = null;
let previewRenderGeneration = 0;

// Programmatic updates from load/reconcile must not be recorded as fresh local
// edits, otherwise opening a document or accepting merged content would mark it
// dirty immediately.
let suppressEditorUpdate = false;

const readableDatabases = computed(() =>
  databases.value.filter((database) => database.capabilities.includes("read")),
);
const creatableDatabases = computed(() =>
  databases.value.filter((database) =>
    database.capabilities.includes("create"),
  ),
);
const selectedDatabaseInfo = computed(
  () =>
    databases.value.find(
      (database) => database.id === selectedDatabaseId.value,
    ) ?? null,
);
const currentDatabaseInfo = computed(
  () =>
    databases.value.find(
      (database) => database.id === currentDatabaseId.value,
    ) ?? null,
);
const currentCanUpdate = computed(
  () => currentDatabaseInfo.value?.capabilities.includes("update") ?? false,
);
const currentCanDelete = computed(
  () => currentDatabaseInfo.value?.capabilities.includes("delete") ?? false,
);
const currentCanBrowseHistory = computed(
  () => currentDatabaseInfo.value?.capabilities.includes("history") ?? false,
);
const isTimeTravelActive = computed(() => launchTimeTravelDate.value != null);
const timeTravelDateLabel = computed(() =>
  launchTimeTravelDate.value == null
    ? ""
    : new Date(launchTimeTravelDate.value).toLocaleString(),
);
const isViewingHistorical = computed(
  () => viewingHistoricalSnapshot.value !== null,
);
const activeDocumentView = computed(
  () => viewingHistoricalSnapshot.value ?? currentDocument.value,
);
const activeRevisionId = computed<MindooDBAppDocumentRevisionId | null>(
  () => viewingHistoricalSnapshot.value?.revisionId ?? null,
);
const canUseAttachments = computed(
  () =>
    currentDatabaseInfo.value?.capabilities.includes("attachments") ?? false,
);
const canManageAttachments = computed(
  () =>
    canUseAttachments.value &&
    currentCanUpdate.value &&
    !isViewingHistorical.value &&
    !isTimeTravelActive.value,
);
const canCreate = computed(
  () => !isTimeTravelActive.value && creatableDatabases.value.length > 0,
);
const subjectDirty = computed(() => subject.value !== savedSubject.value);
const tagsDirty = computed(
  () => JSON.stringify(tags.value) !== JSON.stringify(savedTags.value),
);
const templateDirty = computed(() => isTemplate.value !== savedIsTemplate.value);
const hasLocalEdits = computed(
  () => isDirty.value || subjectDirty.value || tagsDirty.value || templateDirty.value,
);
const openSessions = computed(() =>
  openDocumentSessions.value.map((session) => ({
    id: session.id,
    documentId: session.documentId,
    databaseId: session.databaseId,
    title: session.subject.trim() || "Untitled document",
    type: session.type,
    isActive: session.id === activeDocumentSessionId.value,
    isDirty:
      session.id === activeDocumentSessionId.value
        ? hasLocalEdits.value
        : sessionHasLocalEdits(session),
  })),
);
const isMarkdownDocument = computed(() => currentDocumentType.value === "markdown");
const isWordDocument = computed(() => currentDocumentType.value === "word");
const canSave = computed(() =>
  Boolean(
    !isViewingHistorical.value &&
    !isTimeTravelActive.value &&
    hasLocalEdits.value &&
    currentCanUpdate.value &&
    currentDatabase.value &&
    currentDocument.value,
  ),
);
const canDelete = computed(() =>
  Boolean(
    !isViewingHistorical.value &&
    !isTimeTravelActive.value &&
    currentCanDelete.value &&
    currentDatabase.value &&
    currentDocument.value,
  ),
);
const canRefresh = computed(() =>
  Boolean(currentDatabase.value && currentDocument.value),
);
const canShowInfo = computed(() =>
  Boolean(currentDatabaseId.value && currentDocument.value),
);
const canPrint = computed(() => Boolean(currentDocument.value));
const canImportDocx = computed(() => canCreate.value);
const canOpenWordPageSetup = computed(() =>
  Boolean(currentDocument.value && isWordDocument.value && !editorReadOnly.value),
);
const canExportDocx = computed(() => Boolean(currentDocument.value));
const canExportMarkdown = computed(() => Boolean(currentDocument.value && isMarkdownDocument.value));
const activeAttachments = computed(
  () => activeDocumentView.value?.attachments ?? [],
);
const hasAttachments = computed(() => activeAttachments.value.length > 0);
const canExportMarkdownWithAttachments = computed(() =>
  Boolean(
    currentDatabase.value && currentDocument.value && isMarkdownDocument.value && hasAttachments.value,
  ),
);
const currentRevisionId = computed(() => {
  if (viewingHistoricalSnapshot.value?.revisionId) {
    return viewingHistoricalSnapshot.value.revisionId;
  }
  return (
    revisionEntries.value.find((entry) => entry.isCurrent)?.revisionId ?? null
  );
});
const statusBadgeLabel = computed(() => {
  if (viewingHistoricalSnapshot.value) {
    return `Historical · ${formatRevisionDate(viewingHistoricalSnapshot.value.timestamp)}`;
  }
  if (isTimeTravelActive.value) {
    return `Time travel · ${timeTravelDateLabel.value}`;
  }
  return `Current · ${hasLocalEdits.value ? "Unsaved" : "Saved"}`;
});
const editorReadOnly = computed(
  () => isViewingHistorical.value || isTimeTravelActive.value,
);
const statusBadgeTooltip = computed(() => {
  if (isViewingHistorical.value) {
    return "You're viewing a historical revision. Click to pick a different version or return to the current one.";
  }
  if (isTimeTravelActive.value) {
    return "Time travel mode is active. The whole database is opened read-only.";
  }
  return "You're viewing the current version. Click to browse older revisions.";
});
const splitterLayout = computed(() =>
  previewPanePosition.value === "bottom" ? "vertical" : "horizontal",
);
const documentTitle = computed(
  () => subject.value.trim() || "Untitled document",
);
const pendingCloseSession = computed(
  () =>
    openDocumentSessions.value.find(
      (session) => session.id === pendingCloseSessionId.value,
    ) ?? null,
);
const pendingCloseSessionTitle = computed(
  () => pendingCloseSession.value?.subject.trim() || "Untitled document",
);
const currentDatabaseLabel = computed(() => {
  const info = databases.value.find(
    (database) => database.id === currentDatabaseId.value,
  );
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
  scanAttachment,
  uploadAttachments,
  uploadInputKey: attachmentUploadInputKey,
} = useDocumentAttachments({
  database: currentDatabase,
  document: activeDocumentView,
  revisionId: activeRevisionId,
  onDocumentRefresh(document) {
    currentDocument.value = document;
    snapshotActiveSession();
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
        label: "New Markdown Document",
        icon: "pi pi-file",
        disabled: !canCreate.value,
        command: () => {
          void newMarkdownFile();
        },
      },
      {
        label: "New Word Document",
        icon: "pi pi-file-word",
        disabled: !canCreate.value,
        command: () => {
          void newWordFile();
        },
      },
      {
        label: "New from template...",
        icon: "pi pi-copy",
        disabled: !canCreate.value || readableDatabases.value.length === 0,
        command: () => {
          void openTemplateDialog();
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
      {
        label: "Page Setup",
        icon: "pi pi-file-edit",
        visible: isWordDocument.value,
        disabled: !canOpenWordPageSetup.value,
        command: () => {
          void openWordPageSetup();
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
        label: "Import DOCX",
        icon: "pi pi-file-import",
        disabled: !canImportDocx.value,
        command: () => {
          importCurrentDocx();
        },
      },
      { separator: true },
      {
        label: "Export DOCX",
        icon: "pi pi-file-word",
        disabled: !canExportDocx.value,
        command: () => {
          void exportCurrentDocx();
        },
      },
      {
        label: "Export Markdown",
        icon: "pi pi-file-export",
        visible: isMarkdownDocument.value,
        disabled: !canExportMarkdown.value,
        command: () => {
          void exportCurrentMarkdown();
        },
      },
      {
        label: "Export Markdown with attachments",
        icon: "pi pi-file-export",
        visible: isMarkdownDocument.value && hasAttachments.value,
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
    visible: isMarkdownDocument.value,
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
  {
    label: "Window",
    icon: "pi pi-window-maximize",
    items: [
      ...openSessions.value.map((session) => ({
        label: `${session.isActive ? "\u2713 " : ""}${session.title}${session.isDirty ? " *" : ""}`,
        icon: session.isActive
          ? "pi pi-check"
          : session.type === "word"
            ? "pi pi-file-word"
            : "pi pi-file",
        disabled: session.isActive,
        command: () => switchToOpenSession(session.id),
      })),
      ...(openSessions.value.length > 0 ? [{ separator: true } satisfies MenuItem] : []),
      {
        label: "Close current document",
        icon: "pi pi-times",
        disabled: !currentDocument.value,
        command: () => closeOpenSession(activeDocumentSessionId.value),
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
    currentUserName.value = abbreviateUserName(context.user.username) || "User";
    launchTimeTravelDate.value = context.timeTravelDate ?? null;
    applyAppTheme(context.theme);
    cleanupTheme = nextSession.onThemeChange((theme) => applyAppTheme(theme));
    currentRuntime.value = context.runtime;
    hostUiPreferences.value = { ...context.uiPreferences };
    cleanupUiPreferences = nextSession.onUiPreferencesChange(
      (uiPreferences) => {
        hostUiPreferences.value = { ...uiPreferences };
      },
    );
    databases.value = context.databases;
    selectedDatabaseId.value =
      context.preferredDatabaseId ??
      readableDatabases.value[0]?.id ??
      context.databases[0]?.id ??
      "";
    status.value = "Connected. Choose File / New or File / Open.";
  } catch (error) {
    status.value = error instanceof Error ? error.message : String(error);
  }
});

onBeforeUnmount(async () => {
  if (snapshotActiveSessionTimer) {
    clearTimeout(snapshotActiveSessionTimer);
    snapshotActiveSessionTimer = null;
  }
  stopLivePolling();
  teardownWordAutomergeHandle();
  cleanupTheme?.();
  cleanupUiPreferences?.();
  await session.value?.disconnect();
});

watch(
  () =>
    [
      markdown.value,
      currentDatabaseId.value,
      currentDocument.value?.id,
      imageResolver.revision.value,
    ] as const,
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
    localStorage.setItem(
      PREVIEW_PANE_SETTINGS_KEY,
      JSON.stringify({
        showPreviewPane: nextShowPreviewPane,
        previewPanePosition: nextPreviewPanePosition,
      }),
    );
  },
);

watch(
  () => currentDocument.value,
  () => {
    if (!propertiesDialogVisible.value) {
      resetPropertiesDraft();
    }
  },
  { immediate: true },
);

async function openDatabaseById(databaseId: string) {
  if (!session.value || !databaseId) {
    throw new Error("Select a database first.");
  }
  return await session.value.openDatabase(databaseId);
}

/** Load documents before showing the File/Open picker. */
async function openFileDialog() {
  try {
    openDialogMode.value = "open";
    selectedOpenTemplateFilter.value = "noTemplates";
    await rebuildOpenNavigator();
    selectedOpenCategoryKey.value = ALL_DOCUMENTS_NODE_KEY;
    selectedOpenDocId.value = openDialogDocuments.value[0]?.id ?? "";
    openDialogVisible.value = true;
  } catch (error) {
    status.value = error instanceof Error ? error.message : String(error);
  }
}

async function openTemplateDialog() {
  try {
    openDialogMode.value = "template";
    selectedOpenTemplateFilter.value = "onlyTemplates";
    await rebuildOpenNavigator();
    selectedOpenCategoryKey.value = ALL_DOCUMENTS_NODE_KEY;
    selectedOpenDocId.value = openDialogDocuments.value[0]?.id ?? "";
    openDialogVisible.value = true;
  } catch (error) {
    status.value = error instanceof Error ? error.message : String(error);
  }
}

async function handleOpenDatabaseChange() {
  try {
    const previousDocumentId = selectedOpenDocId.value;
    await rebuildOpenNavigator();
    selectedOpenDocId.value = openDialogDocuments.value.some(
      (document) => document.id === previousDocumentId,
    )
      ? previousDocumentId
      : (openDialogDocuments.value[0]?.id ?? "");
  } catch (error) {
    status.value = error instanceof Error ? error.message : String(error);
  }
}

async function selectOpenCategory(key: string) {
  selectedOpenCategoryKey.value = key;
  await refreshOpenDocumentsForSelectedCategory();
}

async function handleOpenTypeChange() {
  await refreshOpenDocumentsForSelectedCategory();
}

async function handleOpenTemplateFilterChange() {
  try {
    if (openDialogMode.value === "template") {
      selectedOpenTemplateFilter.value = "onlyTemplates";
    }
    await rebuildOpenNavigator();
  } catch (error) {
    status.value = error instanceof Error ? error.message : String(error);
  }
}

async function rebuildOpenNavigator() {
  const databaseInfo = selectedDatabaseInfo.value;
  if (!databaseInfo?.capabilities.includes("views")) {
    throw new Error(
      "This database does not expose the views capability required for categorized Open.",
    );
  }
  await disposeOpenNavigator();
  if (!session.value) {
    throw new Error("Connect to Haven before opening a view.");
  }
  const navigator = await session.value.createViewNavigator({
    databaseIds: [selectedDatabaseId.value],
    definition: createOpenViewDefinition(selectedOpenTemplateFilter.value),
    categorizationStyle: "category_then_document",
    options: {
      includeCategories: true,
      includeDocuments: true,
      hideEmptyCategories: true,
    },
  });
  await navigator.expandAll();
  const entries = await collectNavigatorEntries(navigator);
  const documents = dedupeDocumentEntries(entries);
  const categories = buildOpenCategoryTree(
    entries.filter((entry) => entry.kind === "category"),
    documents.length,
  );
  openNavigator.value = navigator;
  openCategoryNodes.value = categories.roots;
  selectedOpenCategoryKey.value = ALL_DOCUMENTS_NODE_KEY;
  allOpenDialogDocuments.value = documents;
  applyOpenDocumentTypeFilter(documents);
}

async function refreshOpenDocumentsForSelectedCategory() {
  const navigator = openNavigator.value;
  let documents: OpenDocumentRow[];
  if (!navigator || selectedOpenCategoryKey.value === ALL_DOCUMENTS_NODE_KEY) {
    documents = allOpenDialogDocuments.value;
  } else {
    documents = mapDocumentEntries(await navigator.childDocuments(selectedOpenCategoryKey.value));
  }
  applyOpenDocumentTypeFilter(documents);
}

function applyOpenDocumentTypeFilter(documents: OpenDocumentRow[]) {
  openDialogDocuments.value =
    selectedOpenType.value === "all"
      ? documents
      : documents.filter((document) => document.type === selectedOpenType.value);
  if (!openDialogDocuments.value.some((document) => document.id === selectedOpenDocId.value)) {
    selectedOpenDocId.value = openDialogDocuments.value[0]?.id ?? "";
  }
}

async function disposeOpenNavigator() {
  await openNavigator.value?.dispose();
  openNavigator.value = null;
}

function openPropertiesDialog() {
  if (!currentDocument.value || editorReadOnly.value) {
    return;
  }
  resetPropertiesDraft();
  propertiesDialogVisible.value = true;
}

function applyDocumentProperties() {
  if (!currentDocument.value || editorReadOnly.value) {
    return;
  }
  subject.value = propertiesTitleDraft.value.trim();
  // Tags are newline-edited for now so users can paste category paths quickly.
  tags.value = normalizeTags(propertiesTagsDraft.value.split(/\r?\n/));
  isTemplate.value = propertiesIsTemplateDraft.value;
  snapshotActiveSession();
  propertiesDialogVisible.value = false;
}

function resetPropertiesDraft() {
  propertiesTitleDraft.value = subject.value;
  propertiesTagsDraft.value = tags.value.join("\n");
  propertiesIsTemplateDraft.value = isTemplate.value;
}

function createSessionId(databaseId: string, documentId: string) {
  return `${databaseId}:${documentId}`;
}

function tagsEqual(left: string[], right: string[]) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function sessionHasLocalEdits(session: OpenDocumentSession) {
  return (
    session.isDirty ||
    session.subject !== session.savedSubject ||
    !tagsEqual(session.tags, session.savedTags) ||
    session.isTemplate !== session.savedIsTemplate
  );
}

function snapshotActiveSession() {
  if (snapshotActiveSessionTimer) {
    clearTimeout(snapshotActiveSessionTimer);
    snapshotActiveSessionTimer = null;
  }
  if (viewingHistoricalSnapshot.value) {
    return;
  }
  const session = openDocumentSessions.value.find(
    (candidate) => candidate.id === activeDocumentSessionId.value,
  );
  if (!session || !currentDocument.value) {
    return;
  }
  if (currentDatabase.value) {
    session.database = currentDatabase.value;
  }
  session.databaseId = currentDatabaseId.value;
  session.document = currentDocument.value;
  session.documentId = currentDocument.value.id;
  session.type = currentDocumentType.value;
  session.textBuffer = textBuffer.value;
  session.markdown = markdown.value;
  session.wordEditorDocument = wordEditorDocument.value;
  session.currentWordDocument = currentWordDocument.value;
  session.subject = subject.value;
  session.savedSubject = savedSubject.value;
  session.tags = [...tags.value];
  session.savedTags = [...savedTags.value];
  session.isTemplate = isTemplate.value;
  session.savedIsTemplate = savedIsTemplate.value;
  session.isDirty = isDirty.value;
  openDocumentSessions.value = [...openDocumentSessions.value];
}

function scheduleSnapshotActiveSession() {
  if (snapshotActiveSessionTimer) {
    clearTimeout(snapshotActiveSessionTimer);
  }
  snapshotActiveSessionTimer = setTimeout(() => {
    snapshotActiveSessionTimer = null;
    snapshotActiveSession();
  }, SNAPSHOT_ACTIVE_SESSION_DEBOUNCE_MS);
}

function activateSession(session: OpenDocumentSession) {
  activeDocumentSessionId.value = session.id;
  currentDatabase.value = session.database;
  currentDatabaseId.value = session.databaseId;
  currentDocument.value = session.document;
  currentDocumentType.value = session.type;
  viewingHistoricalSnapshot.value = null;
  textBuffer.value = session.textBuffer;
  subject.value = session.subject;
  savedSubject.value = session.savedSubject;
  tags.value = [...session.tags];
  savedTags.value = [...session.savedTags];
  isTemplate.value = session.isTemplate;
  savedIsTemplate.value = session.savedIsTemplate;
  suppressEditorUpdate = true;
  markdown.value = session.markdown;
  wordEditorDocument.value = session.wordEditorDocument;
  currentWordDocument.value = session.currentWordDocument;
  if (session.type === "word" && session.wordEditorDocument) {
    void setupWordAutomergeHandle(session.database, session.document);
  } else {
    teardownWordAutomergeHandle();
  }
  isDirty.value = session.isDirty;
  imageResolver.clear();
  queueMicrotask(() => {
    suppressEditorUpdate = false;
  });
}

function teardownWordAutomergeHandle() {
  const handle = wordAutomergeHandle.value;
  if (handle && wordAutomergeChangeListener) {
    handle.off("change", wordAutomergeChangeListener);
  }
  wordAutomergeHandle.value = null;
  wordAutomergeChangeListener = null;
}

async function setupWordAutomergeHandle(
  database: MindooDBAppDatabase,
  document: MindooDBAppDocument,
  options?: { revisionId?: MindooDBAppDocumentRevisionId },
) {
  teardownWordAutomergeHandle();
  const snapshot = await loadWordAutomergeSnapshot(database, document.id, options);
  const handle = createWordAutomergeHandle({
    database,
    document,
    path: [...WORD_RICH_TEXT_PATH],
    snapshot,
  });
  wordAutomergeChangeListener = (changeSnapshot) => {
    if (isDirty.value || editorReadOnly.value || suppressEditorUpdate) {
      return;
    }
    const comments = currentDocument.value?.data.comments ?? [];
    suppressEditorUpdate = true;
    const reloaded = richTextSpansToDocument(changeSnapshot.spans, comments);
    wordEditorDocument.value = reloaded;
    currentWordDocument.value = reloaded;
    queueMicrotask(() => {
      suppressEditorUpdate = false;
    });
  };
  handle.on("change", wordAutomergeChangeListener);
  wordAutomergeHandle.value = handle;
}

function clearActiveDocumentState() {
  activeDocumentSessionId.value = "";
  currentDatabase.value = null;
  currentDatabaseId.value = "";
  currentDocument.value = null;
  currentDocumentType.value = "markdown";
  viewingHistoricalSnapshot.value = null;
  textBuffer.value = null;
  markdown.value = "";
  wordEditorDocument.value = null;
  currentWordDocument.value = null;
  teardownWordAutomergeHandle();
  subject.value = "";
  savedSubject.value = "";
  tags.value = [];
  savedTags.value = [];
  isTemplate.value = false;
  savedIsTemplate.value = false;
  isDirty.value = false;
  imageResolver.clear();
}

function switchToOpenSession(sessionId: string) {
  snapshotActiveSession();
  const session = openDocumentSessions.value.find(
    (candidate) => candidate.id === sessionId,
  );
  if (!session) {
    status.value = "That document window is no longer open.";
    return;
  }
  activateSession(session);
  status.value = `Switched to ${session.subject.trim() || session.documentId}.`;
}

function closeOpenSession(sessionId: string) {
  if (!sessionId) {
    return;
  }
  snapshotActiveSession();
  const session = openDocumentSessions.value.find(
    (candidate) => candidate.id === sessionId,
  );
  if (!session) {
    return;
  }
  if (sessionHasLocalEdits(session)) {
    pendingCloseSessionId.value = sessionId;
    closeConfirmVisible.value = true;
    return;
  }
  removeOpenSession(sessionId);
}

function confirmCloseOpenSession() {
  const sessionId = pendingCloseSessionId.value;
  closeConfirmVisible.value = false;
  pendingCloseSessionId.value = "";
  if (sessionId) {
    removeOpenSession(sessionId);
  }
}

function removeOpenSession(sessionId: string) {
  const wasActive = activeDocumentSessionId.value === sessionId;
  openDocumentSessions.value = openDocumentSessions.value.filter(
    (candidate) => candidate.id !== sessionId,
  );
  if (!wasActive) {
    status.value = "Closed document window.";
    return;
  }
  const nextSession = openDocumentSessions.value[0] ?? null;
  if (nextSession) {
    activateSession(nextSession);
  } else {
    clearActiveDocumentState();
  }
  status.value = "Closed document window.";
}

/** Create a new empty markdown document. */
async function newMarkdownFile() {
  try {
    const targetDatabaseInfo =
      selectedDatabaseInfo.value?.capabilities.includes("create")
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
        tags: [],
        istemplate: false,
        type: "markdown",
        form: "teamedit",
        body: "",
      },
    });
    await loadDocumentIntoEditor(database, targetDatabaseInfo.id, document);
    status.value = `Created ${document.id}.`;
  } catch (error) {
    status.value = error instanceof Error ? error.message : String(error);
  }
}

/** Backwards-compatible default action for toolbar buttons. */
async function newFile() {
  await newMarkdownFile();
}

/** Create a new empty Word document shell. */
async function newWordFile() {
  try {
    const targetDatabaseInfo =
      selectedDatabaseInfo.value?.capabilities.includes("create")
        ? selectedDatabaseInfo.value
        : creatableDatabases.value[0];
    if (!targetDatabaseInfo) {
      throw new Error("No writable database is available.");
    }
    selectedDatabaseId.value = targetDatabaseInfo.id;
    const database = await openDatabaseById(targetDatabaseInfo.id);
    const defaultWordDocument = createDefaultWordDocument();
    const document = await database.documents.create({
      set: {
        subject: "",
        tags: [],
        istemplate: false,
        type: "word",
        form: "teamedit",
        comments: [],
      },
    });
    const seeded = await seedWordRichTextDocument(database, document, defaultWordDocument);
    await loadDocumentIntoEditor(database, targetDatabaseInfo.id, seeded);
    status.value = `Created Word document ${document.id}.`;
  } catch (error) {
    status.value = error instanceof Error ? error.message : String(error);
  }
}

async function createDocumentFromTemplate(
  templateDatabase: MindooDBAppDatabase,
  templateDatabaseId: string,
  templateDocumentId: string,
) {
  const templateDocument = await templateDatabase.documents.get(templateDocumentId);
  if (!templateDocument) {
    throw new Error("The selected template could not be loaded.");
  }

  const type = readDocumentType(templateDocument);
  const needsUpdate = type === "word";
  const targetDatabaseInfo =
    selectedDatabaseInfo.value?.capabilities.includes("create") &&
    (!needsUpdate || selectedDatabaseInfo.value.capabilities.includes("update"))
      ? selectedDatabaseInfo.value
      : databases.value.find((database) =>
          database.capabilities.includes("create") &&
          (!needsUpdate || database.capabilities.includes("update")),
        );
  if (!targetDatabaseInfo) {
    throw new Error(
      needsUpdate
        ? "No writable database is available for creating a Word document from a template."
        : "No writable database is available for creating from a template.",
    );
  }

  selectedDatabaseId.value = targetDatabaseInfo.id;
  const targetDatabase = targetDatabaseInfo.id === templateDatabaseId
    ? templateDatabase
    : await openDatabaseById(targetDatabaseInfo.id);
  const templateData = cloneDocumentData(templateDocument.data);
  const templateWordDocument = type === "word"
    ? (await loadWordDocumentFromAutomerge(templateDatabase, templateDocument)).wordDocument
    : null;
  const created = await targetDatabase.documents.create({
    set: {
      ...templateData,
      subject: `Copy of ${readSubject(templateDocument) || "Untitled document"}`,
      tags: readTags(templateDocument),
      istemplate: false,
      type,
      form: "teamedit",
      ...(type === "word"
        ? {
            comments: cloneJsonValue(templateDocument.data.comments) ?? [],
          }
        : {
            body: readDocumentBody(templateDocument),
          }),
    },
  });
  const openedDocument = type === "word" && templateWordDocument
    ? await seedWordRichTextDocument(targetDatabase, created, templateWordDocument)
    : created;

  await loadDocumentIntoEditor(targetDatabase, targetDatabaseInfo.id, openedDocument);
  status.value = `Created ${created.id} from template.`;
}

/** Open the selected document and initialize the editor/text buffer. */
async function openSelectedDocument() {
  try {
    const databaseId = selectedDatabaseId.value;
    const database = await openDatabaseById(databaseId);
    if (!selectedOpenDocId.value) {
      throw new Error(
        openDialogMode.value === "template"
          ? "Select a template to create from."
          : "Select a document to open.",
      );
    }
    if (openDialogMode.value === "template") {
      await createDocumentFromTemplate(database, databaseId, selectedOpenDocId.value);
    } else {
      const document = await database.documents.get(selectedOpenDocId.value);
      if (!document) {
        throw new Error("Select a document to open.");
      }
      await loadDocumentIntoEditor(database, databaseId, document);
      status.value = `Opened ${document.id}.`;
    }
    openDialogVisible.value = false;
    await disposeOpenNavigator();
  } catch (error) {
    status.value = error instanceof Error ? error.message : String(error);
  }
}

/* --- Live collaboration -------------------------------------------------- */

/**
 * How often the open editor re-checks Haven for changes made by other
 * devices/users while auto-refresh is enabled. Remote changes are merged via
 * the same reconcile paths the save flow uses (text buffer / Automerge
 * handle), so a poll tick never clobbers local work: it is skipped entirely
 * while there are unsaved local edits.
 */
const LIVE_POLL_INTERVAL_MS = 4000;
const AUTO_REFRESH_STORAGE_KEY = "mindoodb-teamedit-auto-refresh";
let livePollTimer: ReturnType<typeof setInterval> | null = null;
let livePolling = false;

/**
 * Opt-in live collaboration, toggled via the refresh split button's dropdown.
 * Off by default; persisted per device so the choice survives reloads.
 */
const autoRefreshEnabled = ref(readAutoRefreshPreference());

const refreshMenuItems = computed<MenuItem[]>(() => [
  {
    label: "Auto-refresh",
    icon: autoRefreshEnabled.value ? "pi pi-check" : "pi pi-circle",
    command: () => {
      autoRefreshEnabled.value = !autoRefreshEnabled.value;
    },
  },
]);

function readAutoRefreshPreference(): boolean {
  if (typeof localStorage === "undefined") {
    return false;
  }
  return localStorage.getItem(AUTO_REFRESH_STORAGE_KEY) === "1";
}

function startLivePolling() {
  stopLivePolling();
  livePollTimer = setInterval(() => {
    void pollCurrentDocumentForRemoteChanges();
  }, LIVE_POLL_INTERVAL_MS);
}

function stopLivePolling() {
  if (livePollTimer) {
    clearInterval(livePollTimer);
    livePollTimer = null;
  }
}

watch(
  autoRefreshEnabled,
  (enabled) => {
    if (typeof localStorage !== "undefined") {
      localStorage.setItem(AUTO_REFRESH_STORAGE_KEY, enabled ? "1" : "0");
    }
    if (enabled) {
      startLivePolling();
    } else {
      stopLivePolling();
    }
  },
  { immediate: true },
);

/**
 * Live-collaboration poll: pull the merged server state for the open document
 * into the editor between saves. Skips whenever auto-refresh is disabled, the
 * user has unsaved local edits, a historical revision or time travel view is
 * active, or the tab is hidden. Best-effort: failures are swallowed and
 * retried on the next tick.
 */
async function pollCurrentDocumentForRemoteChanges(): Promise<void> {
  if (!autoRefreshEnabled.value || livePolling) {
    return;
  }
  if (typeof document !== "undefined" && document.hidden) {
    return;
  }
  if (!currentDatabase.value || !currentDocument.value) {
    return;
  }
  if (hasLocalEdits.value || editorReadOnly.value) {
    return;
  }
  livePolling = true;
  try {
    if (isWordDocument.value) {
      await pollWordRemote();
    } else {
      await pollMarkdownRemote();
    }
  } catch {
    // Ignore; live polling is best-effort and retries next interval.
  } finally {
    livePolling = false;
  }
}

/** Adopt merged markdown text from Haven while the editor is idle. */
async function pollMarkdownRemote(): Promise<void> {
  const buffer = textBuffer.value;
  const database = currentDatabase.value;
  const documentId = currentDocument.value?.id;
  if (!buffer || buffer.dirty || !database || !documentId) {
    return;
  }
  const fresh = await database.documents.get(documentId);
  if (!fresh) {
    return;
  }
  const changed = buffer.reconcile(fresh);
  currentDocument.value = fresh;
  adoptRemoteMetadata(fresh);
  if (changed) {
    suppressEditorUpdate = true;
    markdown.value = buffer.value;
    queueMicrotask(() => {
      suppressEditorUpdate = false;
    });
  }
}

/**
 * Reload the Word Automerge replica from Haven while the editor is idle. The
 * handle's change listener rebuilds the visible editor document only when the
 * merged spans actually differ.
 */
async function pollWordRemote(): Promise<void> {
  const handle = wordAutomergeHandle.value;
  const database = currentDatabase.value;
  const documentId = currentDocument.value?.id;
  if (!handle || handle.dirty || !database || !documentId) {
    return;
  }
  const fresh = await database.documents.get(documentId);
  if (!fresh) {
    return;
  }
  // Update the document first so the handle's change listener picks up the
  // latest comments alongside the reconciled spans.
  currentDocument.value = fresh;
  adoptRemoteMetadata(fresh);
  handle.syncDocument(fresh);
  const snapshot = await handle.refresh();
  handle.reconcile(fresh, snapshot);
}

/**
 * Mirror remotely changed metadata (subject/tags/template flag) into the
 * editor. Only called from the poll, which already guarantees there are no
 * local edits, so overwriting the drafts is safe.
 */
function adoptRemoteMetadata(document: MindooDBAppDocument) {
  savedSubject.value = readSubject(document);
  subject.value = savedSubject.value;
  savedTags.value = readTags(document);
  tags.value = [...savedTags.value];
  savedIsTemplate.value = readIsTemplate(document);
  isTemplate.value = savedIsTemplate.value;
}

/** Ask before discarding local edits, otherwise refresh immediately. */
function requestRefreshCurrentDocument() {
  if (isViewingHistorical.value) {
    void returnToCurrent();
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
  if (
    !currentDatabase.value ||
    !currentDocument.value ||
    !currentCanBrowseHistory.value
  ) {
    return;
  }
  revisionDialogVisible.value = true;
  revisionLoading.value = true;
  revisionErrorMessage.value = null;
  try {
    revisionEntries.value = await currentDatabase.value.documents.listHistory(
      currentDocument.value.id,
    );
  } catch (error) {
    revisionErrorMessage.value =
      error instanceof Error
        ? error.message
        : "The revision list could not be loaded.";
  } finally {
    revisionLoading.value = false;
  }
}

async function loadHistoricalRevision(
  revisionId: MindooDBAppDocumentRevisionId,
) {
  if (!currentDatabase.value || !currentDocument.value) {
    status.value = "Open a document before loading revisions.";
    return;
  }
  if (
    revisionEntries.value.find((entry) => entry.revisionId === revisionId)
      ?.isCurrent
  ) {
    revisionDialogVisible.value = false;
    await returnToCurrent();
    return;
  }
  try {
    const snapshot = await currentDatabase.value.documents.getAtRevision(
      currentDocument.value.id,
      revisionId,
    );
    if (snapshot.state !== "exists" || !snapshot.data) {
      status.value =
        snapshot.state === "deleted"
          ? "That revision is a deletion marker and cannot be opened in the editor."
          : "That revision is no longer available.";
      return;
    }
    viewingHistoricalSnapshot.value = snapshot;
    textBuffer.value = null;
    teardownWordAutomergeHandle();
    savedSubject.value = readHistoricalSubject(snapshot);
    subject.value = savedSubject.value;
    savedTags.value = readHistoricalTags(snapshot);
    tags.value = [...savedTags.value];
    savedIsTemplate.value = readHistoricalIsTemplate(snapshot);
    isTemplate.value = savedIsTemplate.value;
    suppressEditorUpdate = true;
    markdown.value = readHistoricalBody(snapshot);
    wordEditorDocument.value = null;
    currentWordDocument.value = null;
    if (isWordDocument.value && currentDatabase.value && currentDocument.value) {
      const { wordDocument } = await loadWordDocumentFromAutomerge(
        currentDatabase.value,
        {
          ...currentDocument.value,
          data: snapshot.data,
        },
        { revisionId },
      );
      wordEditorDocument.value = wordDocument;
      currentWordDocument.value = wordDocument;
    }
    isDirty.value = false;
    imageResolver.clear();
    revisionDialogVisible.value = false;
    queueMicrotask(() => {
      suppressEditorUpdate = false;
    });
    status.value = `Loaded revision from ${formatRevisionDate(snapshot.timestamp)}.`;
  } catch (error) {
    status.value =
      error instanceof Error
        ? error.message
        : "The revision could not be loaded.";
  }
}

async function returnToCurrent() {
  if (!currentDatabase.value || !currentDocument.value) {
    return;
  }
  const current = await currentDatabase.value.documents.get(currentDocument.value.id);
  if (!current) {
    status.value = "The current document is no longer available.";
    return;
  }

  let nextWordDocument: DocxDocument | null = null;
  if (readDocumentType(current) === "word") {
    ({ wordDocument: nextWordDocument } = await loadWordDocumentFromAutomerge(
      currentDatabase.value,
      current,
    ));
  }

  currentDocument.value = current;
  currentDocumentType.value = readDocumentType(current);
  viewingHistoricalSnapshot.value = null;
  imageResolver.clear();
  textBuffer.value = currentDocumentType.value === "markdown"
    ? createMindooDBTextBuffer({
        database: currentDatabase.value,
        document: current,
        path: ["body"],
      })
    : null;
  savedSubject.value = readSubject(current);
  subject.value = savedSubject.value;
  savedTags.value = readTags(current);
  tags.value = [...savedTags.value];
  savedIsTemplate.value = readIsTemplate(current);
  isTemplate.value = savedIsTemplate.value;
  suppressEditorUpdate = true;
  markdown.value =
    textBuffer.value?.value ?? readDocumentBody(current);
  wordEditorDocument.value = nextWordDocument;
  currentWordDocument.value = nextWordDocument;
  if (nextWordDocument) {
    void setupWordAutomergeHandle(
      currentDatabase.value,
      current,
    );
  } else {
    teardownWordAutomergeHandle();
  }
  isDirty.value = false;
  queueMicrotask(() => {
    suppressEditorUpdate = false;
  });
  snapshotActiveSession();
  status.value = "Returned to the current version.";
}

/** Re-read the open document from Haven and discard unsaved local edits. */
async function refreshCurrentDocument() {
  if (!currentDatabase.value || !currentDocument.value) {
    status.value = "Open a document before refreshing.";
    return;
  }
  try {
    const document = await currentDatabase.value.documents.get(
      currentDocument.value.id,
    );
    if (!document) {
      throw new Error("The current document could not be loaded.");
    }
    refreshConfirmVisible.value = false;
    await loadDocumentIntoEditor(
      currentDatabase.value,
      currentDatabaseId.value,
      document,
    );
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
  const activeElement =
    document.activeElement instanceof HTMLElement
      ? document.activeElement
      : null;
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

  if (isWordDocument.value) {
    const editor = wordEditorRef.value;
    if (!editor) {
      status.value = "The Word editor is not ready yet.";
      return;
    }
    status.value = "Opening Word print dialog...";
    editor.print();
    status.value = "Word print dialog opened.";
    return;
  }

  const printWindow = window.open("", "_blank");
  if (!printWindow) {
    status.value =
      "Pop-up blocked. Allow pop-ups for this app to print the document.";
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

function importCurrentDocx() {
  if (!canImportDocx.value) {
    status.value = "No writable database is available for DOCX import.";
    return;
  }
  appDocxImportInputRef.value?.click();
}

async function handleAppDocxImport(event: Event) {
  const input = event.target as HTMLInputElement;
  const file = input.files?.[0] ?? null;
  input.value = "";
  if (!file) {
    return;
  }
  try {
    status.value = `Importing ${file.name}...`;
    const parsed = await parseDocx(await file.arrayBuffer());
    attachCommentsToDocument(parsed, commentsFromWordDocument(parsed));
    await importParsedDocxAsNewDocument(file, parsed);
  } catch (error) {
    status.value = error instanceof Error ? error.message : String(error);
  }
}

async function importParsedDocxAsNewDocument(file: File, document: DocxDocument) {
  const targetDatabaseInfo =
    selectedDatabaseInfo.value?.capabilities.includes("create") &&
    selectedDatabaseInfo.value.capabilities.includes("update")
      ? selectedDatabaseInfo.value
      : databases.value.find((database) =>
          database.capabilities.includes("create") &&
          database.capabilities.includes("update"),
        );
  if (!targetDatabaseInfo) {
    throw new Error("No writable database is available for DOCX import.");
  }

  selectedDatabaseId.value = targetDatabaseInfo.id;
  const database = await openDatabaseById(targetDatabaseInfo.id);
  const comments = commentsFromWordDocument(document);
  const imported = await database.documents.create({
    set: {
      subject: createImportedDocxTitle(file),
      tags: [],
      istemplate: false,
      type: "word",
      form: "teamedit",
      comments,
    },
  });
  const seeded = await seedWordRichTextDocument(database, imported, document);
  await loadDocumentIntoEditor(database, targetDatabaseInfo.id, seeded);
  status.value = `Imported ${file.name} as a new Word document.`;
}

function createImportedDocxTitle(file: File) {
  return file.name.replace(/\.docx$/i, "").trim() || "Imported Word document";
}

async function openWordPageSetup() {
  if (!currentDocument.value || !isWordDocument.value) {
    status.value = "Open a Word document before changing page setup.";
    return;
  }
  if (editorReadOnly.value) {
    status.value = "Return to the current version before changing page setup.";
    return;
  }
  const editor = wordEditorRef.value;
  if (!editor) {
    status.value = "The Word editor is not ready yet.";
    return;
  }
  const opened = await editor.openPageSetup();
  status.value = opened
    ? "Opened Word page setup."
    : "The Word page setup dialog is not ready yet.";
}

/** Export only the markdown text, preserving TeamEdit's stable attachment URLs. */
async function exportCurrentMarkdown() {
  if (!currentDocument.value) {
    status.value = "Open a document before exporting.";
    return;
  }

  try {
    const saved = await exportMarkdownFile(
      markdown.value,
      subject.value || "Untitled document",
    );
    status.value = saved
      ? "Exported markdown file."
      : "Markdown export cancelled.";
  } catch (error) {
    status.value =
      error instanceof Error ? error.message : "The markdown export failed.";
  }
}

/** Export a Word document with embedded renderable images and Mermaid diagrams. */
async function exportCurrentDocx() {
  if (!currentDocument.value) {
    status.value = "Open a document before exporting.";
    return;
  }

  if (isWordDocument.value) {
    const editor = wordEditorRef.value;
    if (!editor) {
      status.value = "The Word editor is not ready yet.";
      return;
    }
    status.value = "Preparing Word DOCX export...";
    const blob = await editor.saveDocx();
    if (!blob) {
      status.value = "The Word document could not be exported yet.";
      return;
    }
    const saved = await saveBlobToDisk(
      blob,
      createExportFileName(subject.value || "Untitled document", "docx"),
    );
    status.value = saved ? "Exported Word DOCX file." : "DOCX export cancelled.";
    return;
  }

  try {
    status.value = "Preparing DOCX export...";
    await imageResolver.preloadMarkdownImages(markdown.value);
    const saved = await exportDocxFile({
      markdown: markdown.value,
      title: subject.value || "Untitled document",
      attachments: activeAttachments.value,
      resolveImageUrl: (url) => imageResolver.getCachedImageUrl(url),
    });
    status.value = saved ? "Exported DOCX file." : "DOCX export cancelled.";
  } catch (error) {
    status.value =
      error instanceof Error ? error.message : "The DOCX export failed.";
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
    status.value = saved
      ? "Exported markdown package."
      : "Markdown package export cancelled.";
  } catch (error) {
    status.value =
      error instanceof Error
        ? error.message
        : "The markdown package export failed.";
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
  if (!currentDatabase.value || !currentDocument.value) {
    status.value = "Open or create a document before saving.";
    return;
  }
  if (!currentCanUpdate.value) {
    status.value =
      "This application does not have write access to the current document database.";
    return;
  }
  if (editorReadOnly.value) {
    status.value = isTimeTravelActive.value
      ? "Time travel mode is read-only."
      : "Historical revisions are read-only. Return to the current version before saving.";
    return;
  }
  try {
    let document = currentDocument.value;
    const hasMetadataChanges = subjectDirty.value || tagsDirty.value || templateDirty.value;
    const documentMetadataSet = {
      subject: subject.value,
      tags: tags.value,
      istemplate: isTemplate.value,
      type: currentDocumentType.value,
    };
    if (!isWordDocument.value && hasMetadataChanges) {
      document = await currentDatabase.value.documents.update(document.id, {
        set: documentMetadataSet,
      });
      currentDocument.value = document;
    }

    if (isWordDocument.value) {
      let wordReconciled = false;
      let wordBodyFlushed = false;
      const wordComments = currentWordDocument.value
        ? commentsFromWordDocument(currentWordDocument.value)
        : [];
      const hasCommentChanges = JSON.stringify(wordComments) !== JSON.stringify(document.data.comments ?? []);

      if (isDirty.value && wordAutomergeHandle.value && currentWordDocument.value) {
        const result = await wordAutomergeHandle.value.flush({
          spans: documentToRichTextSpans(currentWordDocument.value),
        });
        document = result.document;
        currentDocument.value = document;
        wordReconciled = result.reconciled;
        wordBodyFlushed = true;
        suppressEditorUpdate = true;
        const reloaded = richTextSpansToDocument(
          result.spans,
          cloneJsonValue(document.data.comments ?? []),
        );
        wordEditorDocument.value = reloaded;
        currentWordDocument.value = reloaded;
        queueMicrotask(() => {
          suppressEditorUpdate = false;
        });
      }

      if (hasMetadataChanges || hasCommentChanges) {
        document = await currentDatabase.value.documents.update(document.id, {
          set: {
            ...(hasMetadataChanges
              ? {
                  subject: subject.value,
                  tags: tags.value,
                  istemplate: isTemplate.value,
                  type: currentDocumentType.value,
                }
              : {}),
            ...(hasCommentChanges ? { comments: wordComments } : {}),
          },
        });
        currentDocument.value = document;
      }

      if (
        wordBodyFlushed
        && !hasMetadataChanges
        && !hasCommentChanges
        && wordAutomergeHandle.value
      ) {
        wordAutomergeHandle.value.syncDocument(document);
      } else {
        await setupWordAutomergeHandle(
          currentDatabase.value,
          document,
        );
      }
      savedSubject.value = readSubject(document);
      subject.value = savedSubject.value;
      savedTags.value = readTags(document);
      tags.value = [...savedTags.value];
      savedIsTemplate.value = readIsTemplate(document);
      isTemplate.value = savedIsTemplate.value;
      isDirty.value = false;
      snapshotActiveSession();
      status.value = wordReconciled
        ? "Saved Word document and reconciled concurrent edits."
        : "Saved Word document.";
      return;
    }

    const result = textBuffer.value?.dirty
      ? await textBuffer.value.flush()
      : { document, value: markdown.value, reconciled: false };
    currentDocument.value = result.document;
    savedSubject.value = readSubject(result.document);
    subject.value = savedSubject.value;
    savedTags.value = readTags(result.document);
    tags.value = [...savedTags.value];
    savedIsTemplate.value = readIsTemplate(result.document);
    isTemplate.value = savedIsTemplate.value;
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
    snapshotActiveSession();
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
    status.value =
      "This application does not have delete access to the current document database.";
    return;
  }
  if (editorReadOnly.value) {
    status.value = isTimeTravelActive.value
      ? "Time travel mode is read-only."
      : "Historical revisions are read-only. Return to the current version before deleting.";
    return;
  }
  try {
    const deletedDocumentId = currentDocument.value.id;
    const deletedSessionId = activeDocumentSessionId.value;
    await currentDatabase.value.documents.delete(deletedDocumentId);
    deleteConfirmVisible.value = false;
    openDocumentSessions.value = openDocumentSessions.value.filter(
      (session) => session.id !== deletedSessionId,
    );
    imageResolver.clear();
    const nextSession = openDocumentSessions.value[0] ?? null;
    if (nextSession) {
      activateSession(nextSession);
    } else {
      clearActiveDocumentState();
    }
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
    throw new Error(
      "This application does not have attachment upload access for the current document database.",
    );
  }

  const attachmentName = createUniqueImageAttachmentName(file.name);
  await uploadFileAttachment(
    currentDatabase.value,
    currentDocument.value.id,
    attachmentName,
    file,
  );
  const refreshed = await currentDatabase.value.documents.get(
    currentDocument.value.id,
  );
  if (refreshed) {
    currentDocument.value = refreshed;
    snapshotActiveSession();
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
    status.value =
      "This application does not have attachment access for the current document database.";
    return Promise.resolve(null);
  }
  if (editorReadOnly.value) {
    status.value = isTimeTravelActive.value
      ? "Time travel mode is read-only."
      : "Return to the current version before inserting attachment links.";
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
 * Create a clean in-memory window session from a Haven document snapshot.
 *
 * Markdown sessions keep a text buffer so pending text patches survive window
 * switches. Word sessions keep the editor model that will be restored on
 * activation.
 */
async function createDocumentSession(
  database: MindooDBAppDatabase,
  databaseId: string,
  document: MindooDBAppDocument,
): Promise<OpenDocumentSession> {
  const type = readDocumentType(document);
  const sessionSubject = readSubject(document);
  const sessionTags = readTags(document);
  const sessionIsTemplate = readIsTemplate(document);
  const sessionTextBuffer = type === "markdown"
    ? createMindooDBTextBuffer({
        database,
        document,
        path: ["body"],
      })
    : null;
  let sessionWordDocument: DocxDocument | null = null;

  if (type === "word") {
    const loaded = await loadWordDocumentFromAutomerge(database, document);
    sessionWordDocument = loaded.wordDocument;
    console.log("[Word load] Editor document model", {
      docId: document.id,
      spanCount: loaded.snapshot.spans.length,
      ...summarizeWordDocument(sessionWordDocument),
    });
  }

  return {
    id: createSessionId(databaseId, document.id),
    databaseId,
    database,
    documentId: document.id,
    document,
    type,
    textBuffer: sessionTextBuffer,
    markdown: sessionTextBuffer?.value ?? "",
    wordEditorDocument: sessionWordDocument,
    currentWordDocument: sessionWordDocument,
    subject: sessionSubject,
    savedSubject: sessionSubject,
    tags: sessionTags,
    savedTags: [...sessionTags],
    isTemplate: sessionIsTemplate,
    savedIsTemplate: sessionIsTemplate,
    isDirty: false,
  };
}

/**
 * Register a document as an open TeamEdit window and make it active.
 */
async function loadDocumentIntoEditor(
  database: MindooDBAppDatabase,
  databaseId: string,
  document: MindooDBAppDocument,
) {
  snapshotActiveSession();
  const nextSession = await createDocumentSession(database, databaseId, document);
  const existingIndex = openDocumentSessions.value.findIndex(
    (candidate) => candidate.id === nextSession.id,
  );
  if (existingIndex >= 0) {
    openDocumentSessions.value = openDocumentSessions.value.map((candidate) =>
      candidate.id === nextSession.id ? nextSession : candidate,
    );
  } else {
    openDocumentSessions.value = [...openDocumentSessions.value, nextSession];
  }
  activateSession(nextSession);
}

/**
 * Receive markdown updates from Milkdown.
 *
 * Milkdown emits full markdown strings rather than raw text splices, so the
 * buffer converts each new value to a minimal text splice with `replaceText()`.
 */
function onEditorUpdate(value: string) {
  if (editorReadOnly.value) {
    return;
  }
  markdown.value = value;
  if (!suppressEditorUpdate) {
    textBuffer.value?.replaceText(value);
    isDirty.value = textBuffer.value?.dirty ?? false;
    snapshotActiveSession();
  }
}

function onWordEditorChange(document: DocxDocument) {
  if (editorReadOnly.value) {
    return;
  }
  currentWordDocument.value = document;
  if (!suppressEditorUpdate) {
    isDirty.value = true;
    scheduleSnapshotActiveSession();
  }
}

async function onWordDocxImport(file: File, document: DocxDocument) {
  try {
    await importParsedDocxAsNewDocument(file, document);
  } catch (error) {
    status.value = error instanceof Error ? error.message : String(error);
  }
}

function readSubject(document: MindooDBAppDocument) {
  const value = document.data.subject;
  return typeof value === "string" ? value : "";
}

function readIsTemplate(document: MindooDBAppDocument) {
  return document.data.istemplate === true;
}

function readDocumentType(document: MindooDBAppDocument): DocumentType {
  return document.data.type === "word" ? "word" : "markdown";
}

function readTags(document: MindooDBAppDocument) {
  return readDocumentTags(document.data);
}

function readDocumentBody(document: MindooDBAppDocument) {
  const value = document.data.body;
  return typeof value === "string" ? value : "";
}

function cloneDocumentData(data: Record<string, unknown>) {
  return cloneJsonValue(data) as Record<string, unknown>;
}

function cloneJsonValue<T>(value: T): T {
  if (value === undefined) {
    return value;
  }
  if (typeof structuredClone === "function") {
    return structuredClone(value);
  }
  return JSON.parse(JSON.stringify(value)) as T;
}

function readHistoricalSubject(snapshot: MindooDBAppHistoricalDocument) {
  const value = snapshot.data?.subject;
  return typeof value === "string" ? value : "";
}

function readHistoricalBody(snapshot: MindooDBAppHistoricalDocument) {
  const value = snapshot.data?.body;
  return typeof value === "string" ? value : "";
}

function readHistoricalTags(snapshot: MindooDBAppHistoricalDocument) {
  return readDocumentTags(snapshot.data);
}

function readHistoricalIsTemplate(snapshot: MindooDBAppHistoricalDocument) {
  return snapshot.data?.istemplate === true;
}

function formatRevisionDate(timestamp: number) {
  return new Intl.DateTimeFormat(undefined, {
    dateStyle: "medium",
    timeStyle: "short",
  }).format(new Date(timestamp));
}
</script>

<template>
  <main class="app-shell">
    <input
      ref="appDocxImportInputRef"
      class="app-shell__hidden-input"
      type="file"
      accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      @change="handleAppDocxImport"
    />
    <Message
      v-if="updateAvailable"
      severity="warn"
      :closable="false"
      class="app-update-banner"
    >
      <div class="app-update-banner__content">
        <div class="app-update-banner__copy">
          <strong>New version available</strong>
          <p>
            Reload TeamEdit to switch to the latest version and refresh
            offline assets.
          </p>
        </div>
        <Button
          label="Reload now"
          size="small"
          :loading="updateReloading"
          @click="reloadForUpdate"
        />
      </div>
    </Message>

    <header
      class="toolbar glass-card"
      :class="{
        'toolbar--ios-multitasking': hostUiPreferences.iosMultitaskingOptimized,
      }"
    >
      <div class="toolbar__leading">
        <span class="toolbar__title">TeamEdit</span>
        <Menubar :model="menuItems" class="toolbar__menubar" />
        <Button
          icon="pi pi-save"
          text
          rounded
          severity="secondary"
          class="toolbar__icon-button"
          aria-label="Save document"
          title="Save document"
          :disabled="!canSave"
          @click="saveFile"
        />
        <SplitButton
          :icon="isViewingHistorical ? 'pi pi-history' : 'pi pi-refresh'"
          text
          severity="secondary"
          class="toolbar__refresh-split"
          :model="refreshMenuItems"
          :disabled="!canRefresh"
          :button-props="{
            'aria-label': isViewingHistorical
              ? 'Return to current version'
              : 'Refresh document',
            title: isViewingHistorical
              ? 'Return to current version'
              : 'Refresh document',
          }"
          :menu-button-props="{
            'aria-label': 'Refresh options',
            title: 'Refresh options',
          }"
          @click="requestRefreshCurrentDocument"
        />
      </div>
      <button
        v-if="currentDocument"
        class="toolbar__document-title"
        type="button"
        :disabled="editorReadOnly"
        @click="openPropertiesDialog"
      >
        {{ documentTitle }}
      </button>
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
      <template v-if="currentDocument">
        <section class="document-area">
          <Splitter
            v-if="isMarkdownDocument && showPreviewPane"
            :layout="splitterLayout"
            class="editor-preview-splitter"
          >
            <SplitterPanel
              :size="previewPanePosition === 'bottom' ? 60 : 62"
              :min-size="25"
              class="splitter-panel"
            >
              <section
                class="editor-panel glass-card"
                :class="{ 'editor-panel--history': editorReadOnly }"
              >
                <div v-if="isTimeTravelActive" class="history-banner">
                  <i class="pi pi-clock" aria-hidden="true" />
                  <span
                    >Time travel mode is active as of
                    {{ timeTravelDateLabel }} - read-only.</span
                  >
                </div>
                <div v-if="isViewingHistorical" class="history-banner">
                  <i class="pi pi-history" aria-hidden="true" />
                  <span
                    >You're viewing the version from
                    {{
                      formatRevisionDate(
                        viewingHistoricalSnapshot?.timestamp ?? 0,
                      )
                    }}
                    - read-only.</span
                  >
                  <button type="button" @click="returnToCurrent">
                    Return to current
                  </button>
                </div>
                <MilkdownMarkdownEditor
                  :key="activeDocumentSessionId"
                  :model-value="markdown"
                  :on-image-upload="uploadEditorImage"
                  :resolve-image-url="imageResolver.resolveImageUrl"
                  :request-attachment-insert="requestAttachmentInsertFromEditor"
                  :readonly="editorReadOnly"
                  @update:model-value="onEditorUpdate"
                />
              </section>
            </SplitterPanel>

            <SplitterPanel
              :size="previewPanePosition === 'bottom' ? 40 : 38"
              :min-size="20"
              class="splitter-panel"
            >
              <section class="preview-panel glass-card">
                <p class="eyebrow">Preview</p>
                <article
                  ref="previewRoot"
                  class="markdown-preview"
                  v-html="renderedMarkdown"
                />
              </section>
            </SplitterPanel>
          </Splitter>

          <section
            v-else-if="isMarkdownDocument"
            class="editor-panel glass-card"
            :class="{ 'editor-panel--history': editorReadOnly }"
          >
            <div v-if="isTimeTravelActive" class="history-banner">
              <i class="pi pi-clock" aria-hidden="true" />
              <span
                >Time travel mode is active as of {{ timeTravelDateLabel }} -
                read-only.</span
              >
            </div>
            <div v-if="isViewingHistorical" class="history-banner">
              <i class="pi pi-history" aria-hidden="true" />
              <span
                >You're viewing the version from
                {{
                  formatRevisionDate(viewingHistoricalSnapshot?.timestamp ?? 0)
                }}
                - read-only.</span
              >
              <button type="button" @click="returnToCurrent">
                Return to current
              </button>
            </div>
            <MilkdownMarkdownEditor
              :key="activeDocumentSessionId"
              :model-value="markdown"
              :on-image-upload="uploadEditorImage"
              :resolve-image-url="imageResolver.resolveImageUrl"
              :request-attachment-insert="requestAttachmentInsertFromEditor"
              :readonly="editorReadOnly"
              @update:model-value="onEditorUpdate"
            />
          </section>
          <section
            v-else
            class="editor-panel glass-card"
            :class="{ 'editor-panel--history': editorReadOnly }"
          >
            <div v-if="isTimeTravelActive" class="history-banner">
              <i class="pi pi-clock" aria-hidden="true" />
              <span
                >Time travel mode is active as of {{ timeTravelDateLabel }} -
                read-only.</span
              >
            </div>
            <div v-if="isViewingHistorical" class="history-banner">
              <i class="pi pi-history" aria-hidden="true" />
              <span
                >You're viewing the version from
                {{
                  formatRevisionDate(viewingHistoricalSnapshot?.timestamp ?? 0)
                }}
                - read-only.</span
              >
              <button type="button" @click="returnToCurrent">
                Return to current
              </button>
            </div>
            <WordDocumentEditor
              :key="activeDocumentSessionId"
              ref="wordEditorRef"
              :author="currentUserName"
              :initial-document="wordEditorDocument"
              :readonly="editorReadOnly"
              :title="subject"
              @change="onWordEditorChange"
              @import="onWordDocxImport"
            />
          </section>
        </section>

        <DocumentAttachmentsPanel
          :attachments="activeAttachments"
          :busy-action="attachmentBusyAction"
          :can-manage-attachments="canManageAttachments"
          :can-use-attachments="canUseAttachments"
          :historical="editorReadOnly"
          :upload-input-key="attachmentUploadInputKey"
          :can-preview-attachment="canPreviewAttachment"
          :format-attachment-size="formatAttachmentSize"
          @upload="uploadAttachments"
          @scan="scanAttachment"
          @preview="previewAttachment"
          @download="downloadAttachment"
          @remove="removeDocumentAttachment"
        />
      </template>
      <section v-else class="empty-state">
        <h1>Collaborative text documents</h1>
        <p>
          Create markdown notes or Word documents. Markdown uses the existing
          text patch flow; Word documents use the rich-text bridge exposed by
          Haven through the app SDK.
        </p>
        <div class="empty-state__actions">
          <Button
            label="New Markdown document"
            icon="pi pi-file-plus"
            :disabled="!canCreate"
            @click="newMarkdownFile"
          />
          <Button
            label="New Word document"
            icon="pi pi-file-word"
            :disabled="!canCreate"
            @click="newWordFile"
          />
          <Button
            label="Open document"
            icon="pi pi-folder-open"
            severity="secondary"
            :disabled="readableDatabases.length === 0"
            @click="openFileDialog"
          />
          <Button
            label="New from template..."
            icon="pi pi-copy"
            severity="secondary"
            :disabled="!canCreate || readableDatabases.length === 0"
            @click="openTemplateDialog"
          />
        </div>
      </section>
    </section>

    <Dialog
      v-model:visible="openDialogVisible"
      modal
      :header="openDialogMode === 'template' ? 'New from template' : 'Open Document'"
      :style="{ width: '58rem', maxWidth: '96vw' }"
      @hide="disposeOpenNavigator"
    >
      <div class="dialog-content">
        <label class="field">
          <span class="field-label">Database</span>
          <select
            v-model="selectedDatabaseId"
            class="native-input"
            @change="handleOpenDatabaseChange"
          >
            <option
              v-for="database in readableDatabases"
              :key="database.id"
              :value="database.id"
            >
              {{ database.title || database.id }}
            </option>
          </select>
        </label>
        <label class="field">
          <span class="field-label">Document type</span>
          <select
            v-model="selectedOpenType"
            class="native-input"
            @change="handleOpenTypeChange"
          >
            <option value="all">All documents</option>
            <option value="markdown">Markdown documents</option>
            <option value="word">Word documents</option>
          </select>
        </label>
        <label class="field">
          <span class="field-label">Template type</span>
          <select
            v-model="selectedOpenTemplateFilter"
            class="native-input"
            :disabled="openDialogMode === 'template'"
            @change="handleOpenTemplateFilterChange"
          >
            <option value="all">All documents</option>
            <option value="noTemplates">No templates</option>
            <option value="onlyTemplates">Only templates</option>
          </select>
        </label>
        <div class="open-dialog__browser">
          <aside class="open-dialog__tree" aria-label="Document tags">
            <TagTreeList
              :nodes="openCategoryNodes"
              :selected-key="selectedOpenCategoryKey"
              @select="selectOpenCategory"
            />
          </aside>
          <div class="document-list">
            <button
              v-for="document in openDialogDocuments"
              :key="document.id"
              class="document-row"
              :class="{
                'document-row--selected': document.id === selectedOpenDocId,
              }"
              type="button"
              @click="selectedOpenDocId = document.id"
              @dblclick="openSelectedDocument"
            >
              <strong>{{ document.title }}</strong>
              <small>{{ document.type === "word" ? "Word" : "Markdown" }} · {{ document.detail }}</small>
              <small>{{ document.id }}</small>
            </button>
            <p
              v-if="openDialogDocuments.length === 0"
              class="document-list__empty"
            >
              {{
                openDialogMode === "template"
                  ? "No template documents in this category."
                  : "No documents in this category."
              }}
            </p>
          </div>
        </div>
      </div>
      <template #footer>
        <Button
          label="Cancel"
          severity="secondary"
          @click="openDialogVisible = false"
        />
        <Button
          :label="openDialogMode === 'template' ? 'Create' : 'Open'"
          :disabled="!selectedOpenDocId"
          @click="openSelectedDocument"
        />
      </template>
    </Dialog>

    <Dialog
      v-model:visible="propertiesDialogVisible"
      modal
      header="Document properties"
      :style="{ width: '34rem', maxWidth: '96vw' }"
    >
      <div class="dialog-content">
        <label class="field">
          <span class="field-label">Title</span>
          <input
            v-model="propertiesTitleDraft"
            class="native-input"
            type="text"
            autocomplete="off"
            placeholder="Document title"
          />
        </label>
        <label class="field">
          <span class="field-label">Tags</span>
          <textarea
            v-model="propertiesTagsDraft"
            class="native-input native-input--textarea"
            rows="6"
            placeholder="Work\Planning&#10;Customer\ABC"
          />
        </label>
        <p class="field-hint">
          Enter one tag per line. Use a backslash to create hierarchy, for
          example <code>Work\Planning</code>.
        </p>
        <label class="properties-dialog__checkbox">
          <input v-model="propertiesIsTemplateDraft" type="checkbox" />
          <span>Use this document as a template</span>
        </label>
      </div>
      <template #footer>
        <Button
          label="Cancel"
          text
          @click="
            propertiesDialogVisible = false;
            resetPropertiesDraft();
          "
        />
        <Button
          label="Apply"
          icon="pi pi-check"
          :disabled="editorReadOnly"
          @click="applyDocumentProperties"
        />
      </template>
    </Dialog>

    <Dialog
      v-model:visible="refreshConfirmVisible"
      modal
      header="Discard local edits?"
      :style="{ width: '28rem', maxWidth: '96vw' }"
    >
      <p>
        Refreshing will reload this document from Haven and discard unsaved
        changes in this window.
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
        Delete document <code>{{ currentDocument?.id }}</code
        >? This removes it from the current database.
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
      v-model:visible="closeConfirmVisible"
      modal
      header="Discard unsaved changes?"
      :style="{ width: '28rem', maxWidth: '96vw' }"
      @hide="pendingCloseSessionId = ''"
    >
      <p>
        Close <strong>{{ pendingCloseSessionTitle }}</strong> and discard its
        unsaved changes in this TeamEdit window?
      </p>
      <template #footer>
        <Button label="Cancel" text @click="closeConfirmVisible = false" />
        <Button
          label="Discard and close"
          icon="pi pi-times"
          severity="warn"
          @click="confirmCloseOpenSession"
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

.app-shell__hidden-input {
  display: none;
}

.app-update-banner {
  left: 50%;
  max-width: min(44rem, calc(100vw - 1.5rem));
  position: fixed;
  top: 0.75rem;
  transform: translateX(-50%);
  width: calc(100vw - 1.5rem);
  z-index: 2400;
}

.app-update-banner__content {
  align-items: center;
  display: flex;
  gap: 1rem;
  justify-content: space-between;
}

.app-update-banner__copy {
  min-width: 0;
}

.app-update-banner__copy p {
  margin: 0.2rem 0 0;
}

.toolbar {
  position: sticky;
  top: 2px;
  z-index: 101;
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

.toolbar__icon-button {
  flex: 0 0 auto;
  width: 1.9rem;
  height: 1.9rem;
  padding: 0;
}

/* Icon-only split button: square refresh action + narrow dropdown arrow. */
.toolbar__refresh-split {
  flex: 0 0 auto;
}

:deep(.toolbar__refresh-split .p-splitbutton-button) {
  width: 1.9rem;
  height: 1.9rem;
  padding: 0;
}

:deep(.toolbar__refresh-split .p-splitbutton-dropdown) {
  width: 1.3rem;
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

.toolbar__document-title {
  flex: 0 1 min(32rem, 36vw);
  max-width: min(32rem, 36vw);
  padding: 0.22rem 0.75rem;
  border: 1px solid transparent;
  border-radius: 999px;
  background: transparent;
  color: var(--text);
  cursor: pointer;
  font-weight: 700;
  overflow: hidden;
  text-align: right;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.toolbar__document-title:hover:not(:disabled),
.toolbar__document-title:focus-visible {
  border-color: var(--border-strong);
  background: rgb(212 160 23 / 0.12);
}

.toolbar__document-title:disabled {
  cursor: default;
  opacity: 0.7;
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

  .toolbar__document-title {
    max-width: 14rem;
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

.native-input--textarea {
  min-height: 8rem;
  resize: vertical;
}

.editor-panel {
  height: 100%;
  overflow: auto;
  display: grid;
  grid-template-rows: minmax(0, 1fr);
  gap: 1rem;
}

.editor-panel--history {
  grid-template-rows: auto minmax(0, 1fr);
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

.markdown-preview {
  line-height: 1.65;
}

.markdown-preview :deep(pre) {
  overflow-x: auto;
  padding: 0.9rem 1rem;
  border: 1px solid var(--border);
  border-radius: var(--radius-md);
  background: rgb(255 255 255 / 0.04);
}

.markdown-preview :deep(code) {
  font-family:
    "JetBrains Mono", ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
  font-size: 0.92em;
}

.markdown-preview :deep(:not(pre) > code) {
  padding: 0.15rem 0.35rem;
  border-radius: 0.35rem;
  background: rgb(255 255 255 / 0.07);
}

.markdown-preview :deep(pre code.hljs) {
  padding: 0;
  background: transparent;
}

.markdown-preview :deep(a) {
  color: var(--accent);
}

.markdown-preview :deep(.teamedit-heading) {
  scroll-margin-top: 1.5rem;
}

.markdown-preview :deep(mark) {
  padding: 0.05rem 0.25rem;
  border-radius: 0.3rem;
  background: rgb(212 160 23 / 0.32);
  color: var(--text);
}

.markdown-preview :deep(sub),
.markdown-preview :deep(sup) {
  line-height: 0;
}

.markdown-preview :deep(.teamedit-align--left) {
  text-align: left;
}

.markdown-preview :deep(.teamedit-align--center) {
  text-align: center;
}

.markdown-preview :deep(.teamedit-align--right) {
  text-align: right;
}

.markdown-preview :deep(.teamedit-align ul),
.markdown-preview :deep(.teamedit-align ol) {
  display: inline-block;
  text-align: left;
}

.markdown-preview :deep(.teamedit-task-list) {
  padding-left: 1.35rem;
}

.markdown-preview :deep(.teamedit-task-list-item) {
  list-style: none;
}

.markdown-preview :deep(.teamedit-task-list-checkbox) {
  width: 1rem;
  height: 1rem;
  margin: 0 0.5rem 0 -1.35rem;
  vertical-align: -0.15em;
  accent-color: var(--accent);
}

.markdown-preview :deep(.footnotes) {
  margin-top: 2rem;
  padding-top: 1rem;
  border-top: 1px solid var(--border);
  color: var(--muted);
  font-size: 0.92em;
}

.markdown-preview :deep(.footnotes ol) {
  padding-left: 1.35rem;
}

.markdown-preview :deep(.footnotes li:target) {
  color: var(--text);
}

.markdown-preview :deep(.footnote-ref),
.markdown-preview :deep(.footnote-backref) {
  font-weight: 600;
  text-decoration: none;
}

.markdown-preview :deep(.teamedit-callout) {
  margin: 1rem 0;
  padding: 0.9rem 1rem;
  border: 1px solid var(--border);
  border-left: 4px solid var(--accent);
  border-radius: var(--radius-md);
  background: rgb(255 255 255 / 0.04);
}

.markdown-preview :deep(.teamedit-callout--tip) {
  border-left-color: #34d399;
}

.markdown-preview :deep(.teamedit-callout--warning),
.markdown-preview :deep(.teamedit-callout--caution) {
  border-left-color: #f59e0b;
}

.markdown-preview :deep(.teamedit-callout--danger) {
  border-left-color: #f87171;
}

.markdown-preview :deep(.teamedit-callout-title) {
  margin: 0 0 0.5rem;
  font-weight: 700;
}

.markdown-preview :deep(.teamedit-callout-body > :first-child) {
  margin-top: 0;
}

.markdown-preview :deep(.teamedit-callout-body > :last-child) {
  margin-bottom: 0;
}

.markdown-preview :deep(table) {
  width: 100%;
  min-width: max(100%, 28rem);
  table-layout: fixed;
  border-collapse: collapse;
  border: 1px solid var(--border);
}

.markdown-preview :deep(th),
.markdown-preview :deep(td) {
  min-width: 7rem;
  height: 2rem;
  padding: 4px 16px;
  border: 1px solid var(--border);
  text-align: left;
  vertical-align: top;
  overflow-wrap: anywhere;
}

.markdown-preview :deep(th) {
  background: rgb(255 255 255 / 0.04);
  font-weight: 600;
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
  margin: auto;
  max-width: 36rem;
  display: grid;
  gap: 0.85rem;
  padding: 1.25rem 1.5rem;
  text-align: center;
}

.empty-state h1 {
  margin: 0;
  font-family: var(--font-title);
  font-size: clamp(1.5rem, 3.5vw, 2.4rem);
  line-height: 1.15;
}

.empty-state p {
  margin: 0;
  color: var(--muted);
  line-height: 1.55;
}

.empty-state__actions {
  display: flex;
  justify-content: center;
  gap: 0.55rem;
  flex-wrap: wrap;
}

.dialog-content {
  display: grid;
  gap: 1rem;
}

.open-dialog__browser {
  min-height: 22rem;
  display: grid;
  grid-template-columns: minmax(12rem, 0.8fr) minmax(0, 1.4fr);
  gap: 0.8rem;
}

.open-dialog__tree {
  max-height: 24rem;
  overflow: auto;
  padding: 0.35rem;
  border: 1px solid var(--border);
  border-radius: 0.85rem;
  background: rgb(255 255 255 / 0.025);
}

.document-list {
  max-height: 24rem;
  overflow: auto;
  display: grid;
  align-content: start;
  gap: 0.45rem;
}

.document-row {
  display: grid;
  gap: 0.2rem;
  padding: 0.75rem 0.85rem;
  border: 1px solid var(--border);
  border-radius: 0.85rem;
  background: rgb(255 255 255 / 0.03);
  color: inherit;
  cursor: pointer;
  text-align: left;
}

.document-row--selected,
.document-row:hover {
  border-color: var(--accent);
  background: rgb(255 255 255 / 0.07);
}

.document-row small,
.document-list__empty {
  color: var(--muted);
}

.document-list__empty {
  margin: 0;
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

.field-hint code {
  color: var(--accent);
}

.properties-dialog__checkbox {
  display: inline-flex;
  align-items: center;
  gap: 0.5rem;
  color: var(--text);
}

@media (max-width: 1100px) {
  .workspace {
    grid-template-columns: 1fr;
  }
}

@media (max-width: 720px) {
  .app-update-banner__content {
    align-items: flex-start;
    flex-direction: column;
  }

  .open-dialog__browser {
    grid-template-columns: 1fr;
  }
}
</style>
