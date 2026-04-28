<script setup lang="ts">
import { computed, onBeforeUnmount, onMounted, ref } from "vue";
import Button from "primevue/button";
import Dialog from "primevue/dialog";
import Menubar from "primevue/menubar";
import Splitter from "primevue/splitter";
import SplitterPanel from "primevue/splitterpanel";
import MarkdownIt from "markdown-it";
import type { MenuItem } from "primevue/menuitem";
import {
  createMindooDBAppBridge,
  createMindooDBTextBuffer,
  type MindooDBAppDatabase,
  type MindooDBAppDatabaseInfo,
  type MindooDBAppDocument,
  type MindooDBAppDocumentSummary,
  type MindooDBAppSession,
  type MindooDBAppUiPreferences,
  type MindooDBTextBuffer,
} from "mindoodb-app-sdk";

import MilkdownMarkdownEditor from "@/components/MilkdownMarkdownEditor.vue";
import { applyAppTheme } from "@/lib/theme";

const markdownIt = new MarkdownIt({
  html: false,
  linkify: true,
  typographer: true,
});

// Bridge/session state is kept in this root component because TeamEdit is a
// small sample app with one active document at a time.
const session = ref<MindooDBAppSession | null>(null);
const databases = ref<MindooDBAppDatabaseInfo[]>([]);
const selectedDatabaseId = ref("");
const documents = ref<MindooDBAppDocumentSummary[]>([]);
const currentDatabase = ref<MindooDBAppDatabase | null>(null);
const currentDatabaseId = ref("");
const currentDocument = ref<MindooDBAppDocument | null>(null);
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
const showPreviewPane = ref(true);
const previewPanePosition = ref<"right" | "bottom">("right");
const selectedOpenDocId = ref("");
const copiedInfoLabel = ref<string | null>(null);
let cleanupTheme: (() => void) | null = null;
let cleanupUiPreferences: (() => void) | null = null;

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
const canCreate = computed(() => creatableDatabases.value.length > 0);
const subjectDirty = computed(() => subject.value !== savedSubject.value);
const hasLocalEdits = computed(() => isDirty.value || subjectDirty.value);
const canSave = computed(() => Boolean(hasLocalEdits.value && currentCanUpdate.value && currentDatabase.value && currentDocument.value));
const canDelete = computed(() => Boolean(currentCanDelete.value && currentDatabase.value && currentDocument.value));
const canRefresh = computed(() => Boolean(currentDatabase.value && currentDocument.value));
const canShowInfo = computed(() => Boolean(currentDatabaseId.value && currentDocument.value));
const renderedMarkdown = computed(() => markdownIt.render(markdown.value || ""));
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
        label: "Info",
        icon: "pi pi-info-circle",
        disabled: !canShowInfo.value,
        command: () => {
          copiedInfoLabel.value = null;
          infoDialogVisible.value = true;
        },
      },
      { separator: true },
      {
        label: "Save",
        icon: "pi pi-save",
        disabled: !canSave.value,
        command: () => {
          void saveFile();
        },
      },
      {
        label: "Delete",
        icon: "pi pi-trash",
        disabled: !canDelete.value,
        command: () => {
          deleteConfirmVisible.value = true;
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

/** Create a new markdown document with a starter `body` field. */
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
        subject: "Untitled Markdown",
        body: "# Untitled Markdown\n\nStart writing together.",
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
  try {
    const deletedDocumentId = currentDocument.value.id;
    await currentDatabase.value.documents.delete(deletedDocumentId);
    deleteConfirmVisible.value = false;
    currentDocument.value = null;
    textBuffer.value = null;
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
 * Replace the active editor state with a document snapshot.
 *
 * This creates a new `MindooDBTextBuffer` bound to the document's `body` path.
 * The buffer stores the document heads from the snapshot so later saves can be
 * merged correctly with concurrent changes.
 */
function loadDocumentIntoEditor(database: MindooDBAppDatabase, databaseId: string, document: MindooDBAppDocument) {
  currentDatabase.value = database;
  currentDatabaseId.value = databaseId;
  currentDocument.value = document;
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
          icon="pi pi-refresh"
          text
          rounded
          severity="secondary"
          class="toolbar__refresh"
          aria-label="Refresh document"
          title="Refresh document"
          :disabled="!canRefresh"
          @click="requestRefreshCurrentDocument"
        />
      </div>
      <div class="toolbar__meta">
        <span class="toolbar__dirty">{{ hasLocalEdits ? "Unsaved" : "Saved" }}</span>
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
            <section class="editor-panel glass-card">
              <template v-if="currentDocument">
                <label class="field subject-field">
                  <span class="field-label">Title</span>
                  <input
                    v-model="subject"
                    class="native-input subject-input"
                    placeholder="Document title"
                    type="text"
                  >
                </label>
                <MilkdownMarkdownEditor
                  :model-value="markdown"
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
              <article class="markdown-preview" v-html="renderedMarkdown" />
            </section>
          </SplitterPanel>
        </Splitter>

        <section v-else class="editor-panel glass-card">
          <template v-if="currentDocument">
            <label class="field subject-field">
              <span class="field-label">Title</span>
              <input
                v-model="subject"
                class="native-input subject-input"
                placeholder="Document title"
                type="text"
              >
            </label>
            <MilkdownMarkdownEditor
              :model-value="markdown"
              @update:model-value="onEditorUpdate"
            />
          </template>
          <div v-else class="empty-state">
            Use File / New or File / Open to start editing a markdown document.
          </div>
        </section>
      </section>
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
  </main>
</template>

<style scoped>
.app-shell {
  min-height: 100vh;
  padding: 2px;
  display: grid;
  grid-template-rows: auto 1fr;
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
  border-radius: 0.6rem;
  backdrop-filter: blur(26px);
  isolation: isolate;
}

.toolbar--ios-multitasking :deep(.p-menubar-button) {
  margin-left: 80px;
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
  gap: 0.15rem;
}

:deep(.toolbar__menubar) {
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
  margin: 0 0.35rem 0 0.3rem;
  font-weight: 700;
  color: var(--muted);
}

.toolbar__dirty {
  padding: 0.12rem 0.42rem;
  border-radius: 999px;
  background: rgb(255 255 255 / 0.06);
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

@media (max-width: 759px) {
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
  display: block;
}

.editor-panel,
.preview-panel {
  min-height: 0;
  padding: 1rem;
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
