<script setup lang="ts">
import { nextTick, shallowRef, watch } from "vue";
import { DocxEditor, createEmptyDocument } from "@eigenpal/docx-editor-vue";
import "@eigenpal/docx-editor-vue/styles.css";
import type { DocxEditorRef } from "@eigenpal/docx-editor-vue";
import { parseDocx } from "@eigenpal/docx-editor-core/docx";
import type { Document as DocxDocument } from "@eigenpal/docx-editor-core/types/document";
import {
  attachCommentsToDocument,
  commentsFromWordDocument,
  summarizeWordDocument,
} from "@/editors/word/lib/richTextSpans";

const props = defineProps<{
  initialDocument?: DocxDocument | null;
  readonly?: boolean;
  title?: string;
}>();

const emit = defineEmits<{
  change: [document: DocxDocument];
  import: [file: File, document: DocxDocument];
}>();

const editorRef = shallowRef<DocxEditorRef | null>(null);
const documentModel = shallowRef<DocxDocument>(createEmptyDocument());
const importError = shallowRef("");
const editorInstanceKey = shallowRef(0);

watch(
  () => props.initialDocument,
  async (document) => {
    const nextDocument = document ?? createEmptyDocument();
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

async function handleChange(document: unknown) {
  await nextTick();
  const nextDocument = readCurrentEditorDocument(document);
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
</script>

<template>
  <section class="word-editor">
    <div class="word-editor__toolbar">
      <label class="word-editor__import">
        Import DOCX
        <input
          type="file"
          accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          :disabled="readonly"
          @change="importDocx"
        />
      </label>
      <span class="word-editor__note">
        Rich-text storage is proxied through Haven via the SDK; this app does not bundle Automerge.
      </span>
    </div>
    <p v-if="importError" class="word-editor__error">{{ importError }}</p>
    <DocxEditor
      :key="editorInstanceKey"
      ref="editorRef"
      :document="documentModel"
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

.word-editor__toolbar {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 1rem;
}

.word-editor__import {
  display: inline-flex;
  cursor: pointer;
  align-items: center;
  border: 1px solid rgb(255 255 255 / 0.18);
  border-radius: 999px;
  padding: 0.45rem 0.8rem;
  font-weight: 700;
}

.word-editor__import input {
  display: none;
}

.word-editor__note,
.word-editor__error {
  color: var(--muted);
  font-size: 0.85rem;
}

.word-editor__error {
  color: var(--danger);
}
</style>
