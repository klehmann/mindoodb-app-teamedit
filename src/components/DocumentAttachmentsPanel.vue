<script setup lang="ts">
import { ref } from "vue";
import Button from "primevue/button";
import Dialog from "primevue/dialog";
import type { MindooDBAppAttachmentInfo } from "mindoodb-app-sdk";

const props = defineProps<{
  attachments: MindooDBAppAttachmentInfo[];
  busyAction: string | null;
  canManageAttachments: boolean;
  canUseAttachments: boolean;
  historical?: boolean;
  uploadInputKey: number;
  canPreviewAttachment: (fileName: string, mimeType: string) => unknown;
  formatAttachmentSize: (size: number) => string;
}>();

const emit = defineEmits<{
  upload: [files: FileList | readonly File[] | null];
  preview: [attachmentName: string];
  download: [attachmentName: string];
  remove: [attachmentName: string];
}>();

const dragActive = ref(false);
const dragDepth = ref(0);
const pendingRemoval = ref<MindooDBAppAttachmentInfo | null>(null);

function dataTransferHasFiles(event: DragEvent) {
  if (event.dataTransfer?.types) {
    return Array.from(event.dataTransfer.types).includes("Files");
  }
  return false;
}

function canAcceptDrop() {
  return props.canManageAttachments && !props.busyAction;
}

function handleUpload(event: Event) {
  const input = event.target as HTMLInputElement;
  emit("upload", input.files);
}

function handleDragEnter(event: DragEvent) {
  if (!dataTransferHasFiles(event) || !canAcceptDrop()) return;
  event.preventDefault();
  dragDepth.value += 1;
  dragActive.value = true;
}

function handleDragOver(event: DragEvent) {
  if (!dataTransferHasFiles(event)) return;
  if (!canAcceptDrop()) {
    if (event.dataTransfer) event.dataTransfer.dropEffect = "none";
    return;
  }
  event.preventDefault();
  if (event.dataTransfer) event.dataTransfer.dropEffect = "copy";
}

function handleDragLeave(event: DragEvent) {
  if (!dataTransferHasFiles(event)) return;
  dragDepth.value = Math.max(0, dragDepth.value - 1);
  if (dragDepth.value === 0) {
    dragActive.value = false;
  }
}

function handleDrop(event: DragEvent) {
  dragDepth.value = 0;
  dragActive.value = false;
  if (!dataTransferHasFiles(event) || !canAcceptDrop()) return;
  event.preventDefault();
  emit("upload", event.dataTransfer?.files ?? null);
}

function requestRemoval(attachment: MindooDBAppAttachmentInfo) {
  if (!props.canManageAttachments || props.busyAction) {
    return;
  }
  pendingRemoval.value = attachment;
}

function confirmRemoval() {
  const attachment = pendingRemoval.value;
  pendingRemoval.value = null;
  if (attachment) {
    emit("remove", attachment.fileName);
  }
}

function updateRemovalDialogVisible(visible: boolean) {
  if (!visible) {
    pendingRemoval.value = null;
  }
}
</script>

<template>
  <section
    class="attachments-panel glass-card"
    :class="{ 'attachments-panel--drag-active': dragActive && canManageAttachments }"
    @dragenter="handleDragEnter"
    @dragover="handleDragOver"
    @dragleave="handleDragLeave"
    @drop="handleDrop"
  >
    <div class="attachments-panel__header">
      <div>
        <p class="attachments-panel__label">Attachments</p>
        <p class="attachments-panel__hint">
          <template v-if="!canUseAttachments">
            Attachment actions are unavailable for this database binding.
          </template>
          <template v-else-if="canManageAttachments && !busyAction">
            Drop files here or use Upload.
          </template>
          <template v-else-if="busyAction">
            {{ busyAction }}...
          </template>
          <template v-else-if="historical">
            Read-only - viewing a historical revision. You can still preview or download files.
          </template>
          <template v-else>
            This app can preview and download attachments, but cannot upload or remove them.
          </template>
        </p>
      </div>

      <label class="upload-button" :class="{ 'upload-button--disabled': !canManageAttachments || Boolean(busyAction) }">
        <input
          :key="uploadInputKey"
          class="sr-only"
          type="file"
          multiple
          :disabled="!canManageAttachments || Boolean(busyAction)"
          @change="handleUpload"
        >
        <span>Upload</span>
      </label>
    </div>

    <div v-if="attachments.length" class="attachment-list">
      <article
        v-for="attachment in attachments"
        :key="attachment.attachmentId"
        class="attachment-item"
      >
        <div class="attachment-item__details">
          <strong>{{ attachment.fileName }}</strong>
          <p>{{ attachment.mimeType }} · {{ formatAttachmentSize(attachment.size) }}</p>
        </div>
        <div class="attachment-item__actions">
          <Button
            icon="pi pi-eye"
            rounded
            text
            severity="secondary"
            :aria-label="`Preview ${attachment.fileName}`"
            :title="`Preview ${attachment.fileName}`"
            :disabled="!canUseAttachments || !canPreviewAttachment(attachment.fileName, attachment.mimeType) || Boolean(busyAction)"
            @click="emit('preview', attachment.fileName)"
          />
          <Button
            icon="pi pi-download"
            rounded
            text
            severity="secondary"
            :aria-label="`Download ${attachment.fileName}`"
            :title="`Download ${attachment.fileName}`"
            :disabled="!canUseAttachments || Boolean(busyAction)"
            @click="emit('download', attachment.fileName)"
          />
          <Button
            icon="pi pi-trash"
            rounded
            text
            severity="danger"
            :aria-label="`Remove ${attachment.fileName}`"
            :title="`Remove ${attachment.fileName}`"
            :disabled="!canManageAttachments || Boolean(busyAction)"
            @click="requestRemoval(attachment)"
          />
        </div>
      </article>
    </div>
    <p v-else class="attachments-panel__empty">No attachments are stored for this document.</p>

    <Dialog
      :visible="Boolean(pendingRemoval)"
      modal
      header="Remove attachment?"
      :style="{ width: '28rem', maxWidth: '96vw' }"
      @update:visible="updateRemovalDialogVisible"
    >
      <p>
        Remove attachment <code>{{ pendingRemoval?.fileName }}</code> from this document?
      </p>
      <template #footer>
        <Button label="Cancel" text @click="pendingRemoval = null" />
        <Button
          label="Remove"
          icon="pi pi-trash"
          severity="danger"
          :disabled="!canManageAttachments || Boolean(busyAction)"
          @click="confirmRemoval"
        />
      </template>
    </Dialog>
  </section>
</template>

<style scoped>
.attachments-panel {
  min-height: 8.5rem;
  max-height: 14rem;
  padding: 0.75rem 1rem;
  display: grid;
  grid-template-rows: auto minmax(0, 1fr);
  gap: 0.6rem;
  border: 2px dashed transparent;
  border-radius: 0;
  transition: border-color 120ms ease, background-color 120ms ease;
}

.attachments-panel--drag-active {
  border-color: var(--p-primary-color, #5298f7);
  background-color: color-mix(in srgb, var(--p-primary-color, #5298f7) 10%, transparent);
}

.attachments-panel--drag-active * {
  pointer-events: none;
}

.attachments-panel__header,
.attachment-item {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 0.75rem;
}

.attachments-panel__label,
.attachments-panel__hint,
.attachments-panel__empty,
.attachment-item__details p {
  margin: 0;
}

.attachments-panel__label {
  color: var(--muted);
  font-size: 0.78rem;
  text-transform: uppercase;
  letter-spacing: 0.08em;
}

.attachments-panel__hint,
.attachments-panel__empty,
.attachment-item__details p {
  color: var(--muted);
  font-size: 0.82rem;
}

.upload-button {
  flex: 0 0 auto;
  padding: 0.4rem 0.75rem;
  border: 1px solid var(--border);
  border-radius: 999px;
  cursor: pointer;
  font-size: 0.85rem;
}

.upload-button--disabled {
  cursor: not-allowed;
  opacity: 0.45;
}

.sr-only {
  position: absolute;
  width: 1px;
  height: 1px;
  padding: 0;
  margin: -1px;
  overflow: hidden;
  clip: rect(0, 0, 0, 0);
  white-space: nowrap;
  border: 0;
}

.attachment-list {
  min-height: 0;
  overflow: auto;
  display: grid;
  gap: 0.45rem;
}

.attachment-item {
  padding: 0.45rem 0.55rem;
  border: 1px solid rgb(255 255 255 / 0.08);
  background: rgb(255 255 255 / 0.03);
}

.attachment-item__details {
  min-width: 0;
}

.attachment-item__details strong,
.attachment-item__details p {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.attachment-item__actions {
  flex: 0 0 auto;
  display: flex;
  align-items: center;
  gap: 0.2rem;
}
</style>
