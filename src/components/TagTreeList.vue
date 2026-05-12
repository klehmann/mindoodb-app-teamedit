<script setup lang="ts">
import type { OpenCategoryNode } from "@/lib/viewOpen";

defineOptions({ name: "TagTreeList" });

defineProps<{
  nodes: OpenCategoryNode[];
  selectedKey: string;
}>();

defineEmits<{
  select: [key: string];
}>();
</script>

<template>
  <ul class="tag-tree-list">
    <li v-for="node in nodes" :key="node.key" class="tag-tree-list__item">
      <button
        type="button"
        class="tag-tree-list__button"
        :class="{ 'tag-tree-list__button--selected': node.key === selectedKey }"
        @click="$emit('select', node.key)"
      >
        <span>{{ node.label }}</span>
        <small>{{ node.count }}</small>
      </button>
      <TagTreeList
        v-if="node.children.length > 0"
        :nodes="node.children"
        :selected-key="selectedKey"
        @select="$emit('select', $event)"
      />
    </li>
  </ul>
</template>

<style scoped>
.tag-tree-list {
  display: grid;
  gap: 0.2rem;
  margin: 0;
  padding: 0;
  list-style: none;
}

.tag-tree-list .tag-tree-list {
  margin-left: 0.8rem;
  padding-left: 0.55rem;
  border-left: 1px solid var(--border);
}

.tag-tree-list__button {
  width: 100%;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 0.5rem;
  padding: 0.45rem 0.55rem;
  border: 1px solid transparent;
  border-radius: 0.65rem;
  background: transparent;
  color: var(--text);
  cursor: pointer;
  text-align: left;
}

.tag-tree-list__button:hover,
.tag-tree-list__button--selected {
  border-color: var(--border-strong);
  background: rgb(212 160 23 / 0.12);
}

.tag-tree-list__button small {
  color: var(--muted);
}
</style>
