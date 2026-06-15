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
const isHydrating = shallowRef(false);
let hydrationGeneration = 0;

watch(
  () => props.initialDocument,
  async (document) => {
    const generation = ++hydrationGeneration;
    isHydrating.value = true;
    const nextDocument = cloneWordDocument(document ?? createDefaultWordDocument());
    attachCommentsToDocument(nextDocument, commentsFromWordDocument(nextDocument));
    documentModel.value = nextDocument;
    editorInstanceKey.value += 1;
    await nextTick();
    window.setTimeout(() => completeHydration(generation), 250);
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

async function handleUpdateDocument(document: unknown) {
  if (!document) {
    return;
  }
  await nextTick();
  // While hydrating, the editor echoes its parsed document back via
  // `update:document` (docx-editor >= 1.5.0). Assigning that echo to
  // `documentModel` would mutate the `:document` prop and re-trigger the
  // editor's loader, which echoes again with a fresh object reference -
  // an unbounded microtask loop that freezes the tab. Drop the echo instead.
  if (isHydrating.value) {
    return;
  }
  const nextDocument = readCurrentEditorDocument(document);
  emit("change", nextDocument);
}

async function handleReady() {
  await nextTick();
  await nextTick();
  window.setTimeout(() => completeHydration(hydrationGeneration), 0);
}

function readCurrentEditorDocument(fallback: unknown) {
  return (editorRef.value?.getDocument?.() as DocxDocument | null | undefined)
    ?? (fallback as DocxDocument);
}

function completeHydration(generation: number) {
  if (generation === hydrationGeneration) {
    isHydrating.value = false;
  }
}

function cloneWordDocument(document: DocxDocument) {
  return structuredClone(document) as DocxDocument;
}

async function saveDocx() {
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
      :author="author"
      :mode="readonly ? 'viewing' : 'editing'"
      :read-only="readonly"
      :show-toolbar="!readonly"
      @update:document="handleUpdateDocument"
      @change="handleUpdateDocument"
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

.word-editor :deep(.menu-bar > .docx-popover:first-child) {
  display: none;
}

.word-editor__error {
  color: var(--danger);
  font-size: 0.85rem;
}
</style>
