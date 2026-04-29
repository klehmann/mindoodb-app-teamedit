<script setup lang="ts">
import { computed, nextTick, ref, watch } from "vue";
import Button from "primevue/button";
import Checkbox from "primevue/checkbox";
import Dialog from "primevue/dialog";
import type { MindooDBAppAttachmentInfo } from "mindoodb-app-sdk";

import {
  type AttachmentInsertion,
  createAttachmentMarkdownUrl,
  formatAttachmentSize,
} from "@/lib/attachmentImages";

/**
 * Picker dialog used by the editor's slash menu to insert an image or link
 * pointing at an existing document attachment.
 *
 * The dialog only renders thumbnails for image attachments and resolves them
 * lazily through the same `resolveImageUrl` resolver used by the editor and
 * preview pane, so we never duplicate attachment downloads.
 */

interface AttachmentRow {
  attachment: MindooDBAppAttachmentInfo;
  isImage: boolean;
  thumbnailUrl: string | null;
  thumbnailFailed: boolean;
}

const props = defineProps<{
  visible: boolean;
  attachments: MindooDBAppAttachmentInfo[];
  resolveImageUrl: (url: string) => Promise<string> | string;
}>();

const emit = defineEmits<{
  "update:visible": [value: boolean];
  select: [selection: AttachmentInsertion];
  cancel: [];
}>();

const imagesOnly = ref(true);
const selectedAttachmentName = ref<string | null>(null);
const altInputDirty = ref(false);
const altText = ref("");
const rows = ref<AttachmentRow[]>([]);

function isImageMimeType(mimeType: string) {
  return mimeType.toLowerCase().startsWith("image/");
}

function deriveDefaultAlt(fileName: string) {
  const base = fileName.split("/").pop() ?? fileName;
  const dotIndex = base.lastIndexOf(".");
  return dotIndex > 0 ? base.slice(0, dotIndex) : base;
}

const filteredRows = computed(() => {
  if (!imagesOnly.value) {
    return rows.value;
  }
  return rows.value.filter((row) => row.isImage);
});

const selectedRow = computed(() => {
  if (!selectedAttachmentName.value) {
    return null;
  }
  return rows.value.find((row) => row.attachment.fileName === selectedAttachmentName.value) ?? null;
});

const canInsert = computed(() => Boolean(selectedRow.value));

const altInputLabel = computed(() => {
  if (!selectedRow.value) {
    return "Alt text";
  }
  return selectedRow.value.isImage ? "Alt text" : "Link label";
});

function rebuildRows() {
  // Sort so that images bubble up and within each group filenames sort naturally.
  const sorted = [...props.attachments].sort((left, right) => {
    const leftIsImage = isImageMimeType(left.mimeType) ? 0 : 1;
    const rightIsImage = isImageMimeType(right.mimeType) ? 0 : 1;
    if (leftIsImage !== rightIsImage) {
      return leftIsImage - rightIsImage;
    }
    return left.fileName.localeCompare(right.fileName, undefined, {
      sensitivity: "base",
      numeric: true,
    });
  });

  rows.value = sorted.map((attachment) => ({
    attachment,
    isImage: isImageMimeType(attachment.mimeType),
    thumbnailUrl: null,
    thumbnailFailed: false,
  }));
}

async function loadThumbnails() {
  // Each row keeps its own load state; we resolve images sequentially on open
  // to keep the visible items priority-ordered without flooding the bridge.
  for (const row of rows.value) {
    if (!row.isImage || row.thumbnailUrl || row.thumbnailFailed) {
      continue;
    }
    try {
      const resolved = await props.resolveImageUrl(createAttachmentMarkdownUrl(row.attachment.fileName));
      row.thumbnailUrl = typeof resolved === "string" ? resolved : null;
    } catch {
      row.thumbnailFailed = true;
    }
  }
}

function selectFirstAvailable() {
  const first = filteredRows.value[0] ?? rows.value[0] ?? null;
  selectedAttachmentName.value = first?.attachment.fileName ?? null;
  altInputDirty.value = false;
  altText.value = first ? deriveDefaultAlt(first.attachment.fileName) : "";
}

function selectRow(row: AttachmentRow) {
  if (selectedAttachmentName.value === row.attachment.fileName) {
    return;
  }
  selectedAttachmentName.value = row.attachment.fileName;
  if (!altInputDirty.value) {
    altText.value = deriveDefaultAlt(row.attachment.fileName);
  }
}

function handleAltInput(event: Event) {
  const input = event.target as HTMLInputElement;
  altInputDirty.value = true;
  altText.value = input.value;
}

watch(
  () => props.visible,
  async (visible) => {
    if (!visible) {
      return;
    }
    rebuildRows();
    selectFirstAvailable();
    await nextTick();
    void loadThumbnails();
  },
  { immediate: true },
);

watch(
  () => props.attachments,
  () => {
    if (!props.visible) {
      return;
    }
    rebuildRows();
    if (!selectedRow.value) {
      selectFirstAvailable();
    } else {
      void loadThumbnails();
    }
  },
);

watch(imagesOnly, () => {
  if (!props.visible) {
    return;
  }
  // When toggling the filter, fall back to the first visible row so the user
  // is never left with a hidden but selected row driving the alt text field.
  if (!filteredRows.value.some((row) => row.attachment.fileName === selectedAttachmentName.value)) {
    selectFirstAvailable();
  }
});

function close() {
  emit("update:visible", false);
}

function cancel() {
  emit("cancel");
  close();
}

function confirm() {
  const row = selectedRow.value;
  if (!row) {
    return;
  }
  emit("select", {
    url: createAttachmentMarkdownUrl(row.attachment.fileName),
    alt: altText.value.trim(),
    isImage: row.isImage,
  });
  close();
}

function handleVisibleChange(value: boolean) {
  if (!value) {
    cancel();
    return;
  }
  emit("update:visible", value);
}
</script>

<template>
  <Dialog
    :visible="visible"
    modal
    header="Insert attachment"
    :style="{ width: '34rem', maxWidth: '96vw' }"
    @update:visible="handleVisibleChange"
  >
    <div class="picker">
      <div class="picker__filter">
        <label class="picker__filter-control">
          <Checkbox v-model="imagesOnly" :binary="true" inputId="attachment-picker-images-only" />
          <span>Images only</span>
        </label>
        <span class="picker__count">
          {{ filteredRows.length }} of {{ rows.length }} attachment{{ rows.length === 1 ? "" : "s" }}
        </span>
      </div>

      <ul v-if="filteredRows.length" class="picker__list" role="listbox">
        <li
          v-for="row in filteredRows"
          :key="row.attachment.attachmentId"
          class="picker__row"
          :class="{ 'picker__row--selected': selectedAttachmentName === row.attachment.fileName }"
          role="option"
          :aria-selected="selectedAttachmentName === row.attachment.fileName"
          tabindex="0"
          @click="selectRow(row)"
          @keydown.enter.prevent="selectRow(row)"
          @keydown.space.prevent="selectRow(row)"
          @dblclick="selectRow(row); confirm()"
        >
          <div class="picker__thumb" :class="{ 'picker__thumb--icon': !row.thumbnailUrl }">
            <img
              v-if="row.isImage && row.thumbnailUrl"
              :src="row.thumbnailUrl"
              :alt="row.attachment.fileName"
            >
            <i v-else-if="row.isImage" class="pi pi-image" aria-hidden="true" />
            <i v-else class="pi pi-paperclip" aria-hidden="true" />
          </div>
          <div class="picker__details">
            <strong>{{ row.attachment.fileName }}</strong>
            <span>{{ row.attachment.mimeType }} · {{ formatAttachmentSize(row.attachment.size) }}</span>
          </div>
        </li>
      </ul>
      <p v-else class="picker__empty">
        <template v-if="imagesOnly && rows.length">
          This document has attachments but no image attachments.
          Untick "Images only" to insert a file link instead.
        </template>
        <template v-else>
          This document has no attachments yet. Upload one in the Attachments
          panel and try again.
        </template>
      </p>

      <label v-if="canInsert" class="picker__alt">
        <span class="field-label">{{ altInputLabel }}</span>
        <input
          class="native-input"
          type="text"
          :value="altText"
          :placeholder="selectedRow?.attachment.fileName ?? ''"
          @input="handleAltInput"
        >
      </label>
    </div>

    <template #footer>
      <Button label="Cancel" text @click="cancel" />
      <Button
        label="Insert"
        icon="pi pi-check"
        :disabled="!canInsert"
        @click="confirm"
      />
    </template>
  </Dialog>
</template>

<style scoped>
.picker {
  display: grid;
  gap: 0.85rem;
}

.picker__filter {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 0.75rem;
  font-size: 0.85rem;
  color: var(--muted);
}

.picker__filter-control {
  display: inline-flex;
  align-items: center;
  gap: 0.5rem;
  cursor: pointer;
}

.picker__count {
  font-variant-numeric: tabular-nums;
}

.picker__list {
  margin: 0;
  padding: 0.25rem;
  list-style: none;
  max-height: 22rem;
  overflow: auto;
  display: grid;
  gap: 0.4rem;
  border: 1px solid var(--border);
  border-radius: 0.85rem;
  background: rgb(255 255 255 / 0.02);
}

.picker__row {
  display: grid;
  grid-template-columns: 3.25rem minmax(0, 1fr);
  align-items: center;
  gap: 0.75rem;
  padding: 0.5rem 0.65rem;
  border: 1px solid transparent;
  border-radius: 0.7rem;
  cursor: pointer;
  outline: none;
  transition: background-color 120ms ease, border-color 120ms ease;
}

.picker__row:hover {
  background: rgb(255 255 255 / 0.04);
}

.picker__row:focus-visible {
  border-color: var(--p-primary-color, #5298f7);
}

.picker__row--selected {
  border-color: var(--p-primary-color, #5298f7);
  background: color-mix(in srgb, var(--p-primary-color, #5298f7) 12%, transparent);
}

.picker__thumb {
  display: grid;
  place-items: center;
  width: 3.25rem;
  height: 3.25rem;
  border-radius: 0.6rem;
  overflow: hidden;
  background: rgb(255 255 255 / 0.04);
}

.picker__thumb img {
  width: 100%;
  height: 100%;
  object-fit: cover;
}

.picker__thumb--icon i {
  font-size: 1.4rem;
  color: var(--muted);
}

.picker__details {
  min-width: 0;
  display: grid;
  gap: 0.15rem;
}

.picker__details strong {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.picker__details span {
  color: var(--muted);
  font-size: 0.8rem;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.picker__empty {
  margin: 0;
  padding: 1.25rem 1rem;
  color: var(--muted);
  text-align: center;
  border: 1px dashed var(--border);
  border-radius: 0.85rem;
  background: rgb(255 255 255 / 0.02);
}

.picker__alt {
  display: grid;
  gap: 0.35rem;
}

.field-label {
  margin: 0;
  color: var(--muted);
  font-size: 0.78rem;
  text-transform: uppercase;
  letter-spacing: 0.08em;
}

.native-input {
  width: 100%;
  padding: 0.65rem 0.8rem;
  border: 1px solid var(--border);
  border-radius: 0.85rem;
  background: rgb(255 255 255 / 0.04);
  color: inherit;
  font: inherit;
}
</style>
