<script setup lang="ts">
import { nextTick, shallowRef, watch } from "vue";
import { DocxEditor } from "@eigenpal/docx-editor-vue";
import "@eigenpal/docx-editor-vue/styles.css";
import type { DocxEditorRef } from "@eigenpal/docx-editor-vue";
import { parseDocx } from "@eigenpal/docx-editor-core/docx";
import type { Document as DocxDocument } from "@eigenpal/docx-editor-core/types/document";
import {
  attachCommentsToDocument,
  commentsFromWordDocument,
  createDefaultWordDocument,
  summarizeWordDocument,
} from "@/editors/word/lib/richTextSpans";

const DOCX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const DOCX_FILE_MENU_SELECTOR = ".menu-bar > .docx-popover:first-child";
const DOCX_FILE_MENU_TRIGGER_SELECTOR = `${DOCX_FILE_MENU_SELECTOR} .docx-menu-dropdown__trigger`;
const DOCX_FILE_MENU_ITEM_SELECTOR = `${DOCX_FILE_MENU_SELECTOR} .docx-menu-dropdown__item`;
const DOCX_PAGE_SETUP_ITEM_INDEX = 2;

const props = defineProps<{
  initialDocument?: DocxDocument | null;
  author?: string;
  readonly?: boolean;
  title?: string;
}>();

const emit = defineEmits<{
  change: [document: DocxDocument];
  import: [file: File, document: DocxDocument];
}>();

const editorRootRef = shallowRef<HTMLElement | null>(null);
const editorRef = shallowRef<DocxEditorRef | null>(null);
const importInputRef = shallowRef<HTMLInputElement | null>(null);
const documentModel = shallowRef<DocxDocument>(createDefaultWordDocument());
const importError = shallowRef("");
const editorInstanceKey = shallowRef(0);

watch(
  () => props.initialDocument,
  async (document) => {
    const nextDocument = document ?? createDefaultWordDocument();
    attachCommentsToDocument(nextDocument, commentsFromWordDocument(nextDocument));
    documentModel.value = nextDocument;
    editorInstanceKey.value += 1;
  },
  { immediate: true },
);

async function importDocx(event: Event) {
  const input = event.target as HTMLInputElement;
  const file = input.files?.[0] ?? null;
  input.value = "";
  if (!file) {
    return;
  }
  try {
    importError.value = "";
    const parsed = await parseDocx(await file.arrayBuffer());
    attachCommentsToDocument(parsed, commentsFromWordDocument(parsed));
    documentModel.value = parsed;
    editorInstanceKey.value += 1;
    emit("import", file, parsed);
  } catch (error) {
    importError.value = error instanceof Error ? error.message : String(error);
  }
}

function openImportDialog() {
  if (props.readonly) {
    return;
  }
  importInputRef.value?.click();
}

async function handleChange(document: unknown) {
  await nextTick();
  const nextDocument = readCurrentEditorDocument(document);
  applyCommentAuthor(nextDocument);
  console.log("[Word editor] change", summarizeWordDocument(nextDocument));
  emit("change", nextDocument);
}

async function handleUpdateDocument(document: unknown) {
  if (!document) {
    console.log("[Word editor] update:document", { document: null });
    return;
  }
  await nextTick();
  const nextDocument = readCurrentEditorDocument(document);
  applyCommentAuthor(nextDocument);
  console.log("[Word editor] update:document", summarizeWordDocument(nextDocument));
  emit("change", nextDocument);
}

async function handleReady() {
  await nextTick();
  const document = editorRef.value?.getDocument?.() as DocxDocument | null | undefined;
  console.log("[Word editor] ready", document ? summarizeWordDocument(document) : { document: null });
}

function readCurrentEditorDocument(fallback: unknown) {
  return (editorRef.value?.getDocument?.() as DocxDocument | null | undefined)
    ?? (fallback as DocxDocument);
}

// The Vue DOCX editor currently creates UI comments with the hard-coded
// author "User"; replace only that default so imported or existing authors
// remain intact while new TeamEdit comments use the SDK launch username.
function applyCommentAuthor(document: DocxDocument) {
  const author = props.author?.trim();
  if (!author) {
    return;
  }
  const comments = commentsFromWordDocument(document);
  let changed = false;
  const authoredComments = comments.map((comment) => {
    if (!comment || typeof comment !== "object") {
      return comment;
    }
    const currentAuthor = (comment as { author?: unknown }).author;
    if (currentAuthor && currentAuthor !== "User") {
      return comment;
    }
    changed = true;
    return {
      ...comment,
      author,
    };
  });
  if (changed) {
    attachCommentsToDocument(document, authoredComments);
  }
}

async function saveDocx() {
  const document = editorRef.value?.getDocument?.() as DocxDocument | null | undefined;
  if (document) {
    applyCommentAuthor(document);
  }
  const saved = await editorRef.value?.save?.() as ArrayBuffer | Blob | null | undefined;
  if (!saved) {
    return null;
  }
  return saved instanceof Blob
    ? saved
    : new Blob([saved], { type: DOCX_MIME_TYPE });
}

async function openPageSetup() {
  if (props.readonly) {
    return false;
  }

  // The Vue DOCX editor has no public page-setup ref API yet. Its dialog is
  // only opened through the built-in File menu, so TeamEdit triggers that
  // existing menu item from here while keeping file open/save centralized in
  // the app-level menu.
  const root = editorRootRef.value;
  const trigger = root?.querySelector<HTMLButtonElement>(DOCX_FILE_MENU_TRIGGER_SELECTOR);
  if (!trigger) {
    return false;
  }
  trigger.click();
  await nextTick();
  const items = root?.querySelectorAll<HTMLButtonElement>(DOCX_FILE_MENU_ITEM_SELECTOR);
  const pageSetupItem = items?.[DOCX_PAGE_SETUP_ITEM_INDEX] ?? null;
  if (!pageSetupItem) {
    return false;
  }
  pageSetupItem.click();
  return true;
}

function print() {
  const pagesElement = editorRootRef.value?.querySelector(".paged-editor__pages");
  if (!pagesElement) {
    window.print();
    return;
  }

  const printWindow = window.open("", "_blank");
  if (!printWindow) {
    window.print();
    return;
  }

  const pagesClone = pagesElement.cloneNode(true) as HTMLElement;
  pagesClone.style.cssText = "display: block; margin: 0; padding: 0;";
  for (const page of pagesClone.querySelectorAll(".layout-page")) {
    const element = page as HTMLElement;
    element.style.boxShadow = "none";
    element.style.margin = "0";
  }

  printWindow.document.write(`<!DOCTYPE html>
<html>
<head>
  <title>${escapeHtml(props.title || "Print")}</title>
  <style>
${collectFontFaceRules().join("\n")}
* { margin: 0; padding: 0; box-sizing: border-box; }
body { background: white; }
.layout-page { break-after: page; page-break-after: always; }
.layout-page:last-child { break-after: auto; page-break-after: auto; }
img { max-width: 100%; page-break-inside: avoid; }
table, tr { page-break-inside: avoid; }
@page { margin: 0; size: auto; }
  </style>
</head>
<body>${pagesClone.outerHTML}</body>
</html>`);
  printWindow.document.close();

  printWindow.onload = () => {
    printWindow.print();
    printWindow.close();
  };

  setTimeout(() => {
    if (!printWindow.closed) {
      printWindow.print();
      printWindow.close();
    }
  }, 1000);
}

function collectFontFaceRules() {
  const rules: string[] = [];
  for (const sheet of Array.from(document.styleSheets)) {
    try {
      for (const rule of Array.from(sheet.cssRules)) {
        if (rule instanceof CSSFontFaceRule) {
          rules.push(rule.cssText);
        }
      }
    } catch {
      // Cross-origin stylesheets cannot be inspected.
    }
  }
  return rules;
}

function escapeHtml(value: string) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

defineExpose({
  openImportDialog,
  openPageSetup,
  print,
  saveDocx,
});
</script>

<template>
  <section ref="editorRootRef" class="word-editor">
    <input
      ref="importInputRef"
      class="word-editor__hidden-input"
      type="file"
      accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      :disabled="readonly"
      @change="importDocx"
    />
    <p v-if="importError" class="word-editor__error">{{ importError }}</p>
    <DocxEditor
      :key="editorInstanceKey"
      ref="editorRef"
      :document="documentModel"
      :mode="readonly ? 'viewing' : 'editing'"
      :read-only="readonly"
      :show-toolbar="!readonly"
      @change="handleChange"
      @update:document="handleUpdateDocument"
      @ready="handleReady"
    />
  </section>
</template>

<style scoped>
.word-editor {
  display: grid;
  min-height: 100%;
  gap: 0.75rem;
}

.word-editor__hidden-input {
  display: none;
}

.word-editor :deep(.doc-name) {
  display: none;
}

/* The Vue DOCX editor only exposes a whole-title-bar toggle today, not a
   partial menu API. TeamEdit owns file open/save/export, so hide just the
   embedded File dropdown while keeping Format/Insert/Help available. */
.word-editor :deep(.menu-bar > .docx-popover:first-child) {
  display: none;
}

.word-editor__error {
  color: var(--muted);
  font-size: 0.85rem;
}

.word-editor__error {
  color: var(--danger);
}
</style>
