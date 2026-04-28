<script setup lang="ts">
import { nextTick, onBeforeUnmount, onMounted, ref, watch } from "vue";

const props = defineProps<{
  modelValue: string;
  onImageUpload?: (file: File) => Promise<string>;
  resolveImageUrl?: (url: string) => Promise<string> | string;
}>();

const emit = defineEmits<{
  "update:modelValue": [value: string];
}>();

const root = ref<HTMLElement | null>(null);
let editor: any = null;
let applyingExternalValue = false;
let editorGeneration = 0;

onMounted(createEditor);

async function createEditor() {
  await nextTick();
  if (!root.value) {
    return;
  }
  const generation = ++editorGeneration;
  const { Crepe } = await import("@milkdown/crepe");
  if (!root.value || generation !== editorGeneration) {
    return;
  }
  editor = new Crepe({
    root: root.value,
    defaultValue: props.modelValue,
    featureConfigs: {
      [Crepe.Feature.ImageBlock]: {
        // Crepe stores the returned URL in markdown. TeamEdit returns stable
        // MindooDB attachment URLs here, not temporary object URLs.
        onUpload: props.onImageUpload,
        proxyDomURL: props.resolveImageUrl,
      },
    },
  });
  editor.on((listener: any) => {
    listener.markdownUpdated((_ctx: unknown, markdown: string) => {
      if (!applyingExternalValue) {
        emit("update:modelValue", markdown);
      }
    });
  });
  await editor.create();
}

async function recreateEditor() {
  applyingExternalValue = true;
  const previousEditor = editor;
  editor = null;
  editorGeneration += 1;
  try {
    await previousEditor?.destroy?.();
    if (root.value) {
      root.value.replaceChildren();
    }
    await createEditor();
  } finally {
    await nextTick();
    applyingExternalValue = false;
  }
}

watch(
  () => props.modelValue,
  async (value) => {
    if (!editor || value === editor.getMarkdown?.()) {
      return;
    }
    await recreateEditor();
  },
);

onBeforeUnmount(async () => {
  editorGeneration += 1;
  editor?.destroy?.();
  editor = null;
});
</script>

<template>
  <div ref="root" class="milkdown-host" />
</template>

<style scoped>
.milkdown-host {
  min-height: 100%;
}

:deep(.milkdown) {
  min-height: 100%;
  border-radius: var(--radius-md);
}
</style>
