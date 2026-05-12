<script setup lang="ts">
import { commandsCtx, editorViewCtx } from "@milkdown/kit/core";
import { clearTextInCurrentBlockCommand } from "@milkdown/kit/preset/commonmark";
import { insert } from "@milkdown/kit/utils";
import { nextTick, onBeforeUnmount, onMounted, ref, watch } from "vue";

import type { AttachmentInsertion } from "@/lib/attachmentImages";
import { escapeHtml } from "@/lib/markdownRendering";
import { isMermaidLanguage, MERMAID_PLACEHOLDER_CLASS, renderMermaidSvg } from "@/lib/mermaid";

const props = defineProps<{
  modelValue: string;
  onImageUpload?: (file: File) => Promise<string>;
  resolveImageUrl?: (url: string) => Promise<string> | string;
  /**
   * Called by the slash menu's "Attachment" entry. The host opens a picker UI
   * with the document's existing attachments; resolving with `null` cancels
   * the insertion.
   */
  requestAttachmentInsert?: () => Promise<AttachmentInsertion | null>;
  readonly?: boolean;
}>();

const emit = defineEmits<{
  "update:modelValue": [value: string];
}>();

const root = ref<HTMLElement | null>(null);
let editor: any = null;
let applyingExternalValue = false;
let editorGeneration = 0;

const MERMAID_TEMPLATE = "```mermaid\nflowchart TD\n  A[Start] --> B[Next step]\n```";
const CALLOUT_TEMPLATE = ":::note Note\nWrite callout content here.\n:::";
const diagramIcon = `
  <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24">
    <path d="M4 4h6v6H4V4Zm10 0h6v6h-6V4ZM4 14h6v6H4v-6Zm10 0h6v6h-6v-6ZM8 10h2v2h4v2h2v-2h2v-2h-4V8h-2v2H8v2Z"/>
  </svg>
`;
const calloutIcon = `
  <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24">
    <path d="M4 4h16v16H4V4Zm2 2v12h12V6H6Zm2 2h8v2H8V8Zm0 4h5v2H8v-2Z"/>
  </svg>
`;
const attachmentIcon = `
  <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24">
    <path d="M16.5 6.5v9.75a4.75 4.75 0 0 1-9.5 0V5.5a3.25 3.25 0 0 1 6.5 0v10.25a1.75 1.75 0 0 1-3.5 0V7H11v8.75a.75.75 0 0 0 1.5 0V5.5a1.75 1.75 0 0 0-3.5 0v10.75a3.25 3.25 0 0 0 6.5 0V6.5h1Z"/>
  </svg>
`;
const highlightIcon = `
  <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24">
    <path d="M5 18h14v2H5v-2Zm2.8-2 7.7-7.7-2.8-2.8L5 13.2V16h2.8ZM14.1 4.1l1.4-1.4a1 1 0 0 1 1.4 0l1.4 1.4a1 1 0 0 1 0 1.4l-1.4 1.4-2.8-2.8Z"/>
  </svg>
`;
const subscriptIcon = `
  <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24">
    <path d="m5 6 4 5-4 5h2.5l2.7-3.5L13 16h2.5l-4-5 4-5H13l-2.8 3.5L7.5 6H5Zm11 9h3c.6 0 1 .4 1 1v1c0 .6-.4 1-1 1h-2v1h3v2h-5v-3c0-.6.4-1 1-1h2v-1h-3v-1Z"/>
  </svg>
`;
const superscriptIcon = `
  <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24">
    <path d="m5 8 4 5-4 5h2.5l2.7-3.5L13 18h2.5l-4-5 4-5H13l-2.8 3.5L7.5 8H5Zm11-5h3c.6 0 1 .4 1 1v1c0 .6-.4 1-1 1h-2v1h3v2h-5V6c0-.6.4-1 1-1h2V4h-3V3Z"/>
  </svg>
`;

function escapeMarkdownAlt(value: string) {
  // Markdown image alt and link labels are bracketed; escape the closing brace
  // and backslashes so user-provided text never breaks the surrounding syntax.
  return value.replace(/\\/g, "\\\\").replace(/]/g, "\\]");
}

function escapeMarkdownUrl(value: string) {
  // Parentheses around the URL also need escaping. We expect the
  // mindoodb-attachment scheme so this is mostly defensive.
  return value.replace(/\\/g, "\\\\").replace(/\)/g, "\\)");
}

function buildAttachmentMarkdown(result: AttachmentInsertion) {
  const alt = escapeMarkdownAlt(result.alt || "");
  const url = escapeMarkdownUrl(result.url);
  return result.isImage ? `![${alt}](${url})` : `[${alt || url}](${url})`;
}

function wrapSelectionWithMarkdown(ctx: any, before: string, after = before) {
  const view = ctx.get(editorViewCtx);
  const { state } = view;
  const { from, to } = state.selection;
  const selectedText = state.doc.textBetween(from, to, "\n");
  if (!selectedText) {
    return;
  }

  view.dispatch(
    state.tr
      .insertText(after, to)
      .insertText(before, from)
      .scrollIntoView(),
  );
  view.focus();
}

function createMermaidMessageHtml(message: string, tone: "muted" | "error" = "muted") {
  return [
    `<div class="${MERMAID_PLACEHOLDER_CLASS}">`,
    `<p class="teamedit-mermaid-message teamedit-mermaid-message--${tone}">${escapeHtml(message)}</p>`,
    `</div>`,
  ].join("");
}

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
      [Crepe.Feature.CodeMirror]: {
        renderPreview(language: string, content: string, applyPreview: (value: null | string | HTMLElement) => void) {
          if (!isMermaidLanguage(language)) {
            return null;
          }
          applyPreview(createMermaidMessageHtml("Loading Mermaid preview..."));
          void renderMermaidSvg(content, {
            idPrefix: "teamedit-editor-mermaid",
            convertLabelsToSvgText: true,
          })
            .then((svg) => {
              applyPreview(`<div class="${MERMAID_PLACEHOLDER_CLASS}">${svg}</div>`);
            })
            .catch((error) => {
              applyPreview(createMermaidMessageHtml(
                error instanceof Error ? error.message : "Unable to render Mermaid diagram.",
                "error",
              ));
            });
          return undefined;
        },
      },
      [Crepe.Feature.Toolbar]: {
        buildToolbar(builder: any) {
          builder.addGroup("teamedit-format", "TeamEdit").addItem("highlight", {
            icon: highlightIcon,
            active: () => false,
            onRun(ctx: any) {
              wrapSelectionWithMarkdown(ctx, "==");
            },
          }).addItem("subscript", {
            icon: subscriptIcon,
            active: () => false,
            onRun(ctx: any) {
              wrapSelectionWithMarkdown(ctx, "~");
            },
          }).addItem("superscript", {
            icon: superscriptIcon,
            active: () => false,
            onRun(ctx: any) {
              wrapSelectionWithMarkdown(ctx, "^");
            },
          });
        },
      },
      [Crepe.Feature.BlockEdit]: {
        buildMenu(builder: any) {
          builder.getGroup("advanced").addItem("callout", {
            label: "Callout",
            icon: calloutIcon,
            onRun(ctx: any) {
              ctx.get(commandsCtx).call(clearTextInCurrentBlockCommand.key);
              insert(CALLOUT_TEMPLATE)(ctx);
            },
          });
          builder.addGroup("diagram", "Diagram").addItem("mermaid", {
            label: "Diagram",
            icon: diagramIcon,
            onRun(ctx: any) {
              ctx.get(commandsCtx).call(clearTextInCurrentBlockCommand.key);
              insert(MERMAID_TEMPLATE)(ctx);
            },
          });
          // The "Attachment" entry lets the user reference an existing document
          // attachment without re-uploading it. The picker dialog lives in the
          // host so it can read the document's attachments and resolve image
          // thumbnails with the same resolver the editor uses.
          builder.addGroup("attachment", "Attachment").addItem("attachment", {
            label: "Attachment",
            icon: attachmentIcon,
            onRun: (ctx: any) => {
              ctx.get(commandsCtx).call(clearTextInCurrentBlockCommand.key);
              const requestInsert = props.requestAttachmentInsert;
              if (!requestInsert) {
                return;
              }
              void Promise.resolve(requestInsert()).then((result) => {
                if (!result) {
                  return;
                }
                insert(buildAttachmentMarkdown(result))(ctx);
              });
            },
          });
        },
      },
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
  editor.setReadonly?.(props.readonly ?? false);
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

watch(
  () => props.readonly,
  (value) => {
    editor?.setReadonly?.(value ?? false);
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
  height: 100%;
  min-height: 100%;
  overflow: auto;
  border-radius: var(--radius-md);
  background: var(--crepe-color-background);
  color: var(--crepe-color-on-background);
  font-family: var(--crepe-font-default);
}

:deep(.milkdown .ProseMirror) {
  min-height: 100%;
  padding: clamp(1.5rem, 6vw, 3.75rem) clamp(1rem, 10vw, 7.5rem);
  outline: none;
}

:deep(.milkdown .ProseMirror ::selection) {
  background: var(--crepe-color-selected);
}

:deep(.milkdown .ProseMirror h1),
:deep(.milkdown .ProseMirror h2),
:deep(.milkdown .ProseMirror h3),
:deep(.milkdown .ProseMirror h4),
:deep(.milkdown .ProseMirror h5),
:deep(.milkdown .ProseMirror h6) {
  padding: 2px 0;
  color: var(--crepe-color-on-background);
  font-family: var(--crepe-font-title);
  font-weight: 400;
  letter-spacing: -0.02em;
}

:deep(.milkdown .ProseMirror h1) {
  margin-top: 32px;
  font-size: 42px;
  line-height: 50px;
}

:deep(.milkdown .ProseMirror h2) {
  margin-top: 28px;
  font-size: 36px;
  line-height: 44px;
}

:deep(.milkdown .ProseMirror h3) {
  margin-top: 24px;
  font-size: 32px;
  line-height: 40px;
}

:deep(.milkdown .ProseMirror h4) {
  margin-top: 20px;
  font-size: 28px;
  line-height: 36px;
}

:deep(.milkdown .ProseMirror h5) {
  margin-top: 16px;
  font-size: 24px;
  line-height: 32px;
}

:deep(.milkdown .ProseMirror h6) {
  margin-top: 16px;
  font-size: 18px;
  font-weight: 700;
  line-height: 28px;
}

:deep(.milkdown .ProseMirror p) {
  padding: 4px 0;
  font-size: 16px;
  line-height: 24px;
}

:deep(.milkdown .ProseMirror a) {
  color: var(--crepe-color-primary);
  text-decoration: underline;
  text-underline-offset: 0.18em;
}

:deep(.milkdown .ProseMirror code) {
  display: inline-block;
  padding: 0 0.2em;
  border-radius: 4px;
  background: var(--crepe-color-inline-area);
  color: var(--crepe-color-inline-code);
  font-family: var(--crepe-font-code);
  font-size: 87.5%;
  line-height: 1.4286;
}

:deep(.milkdown .ProseMirror pre) {
  margin: 4px 0;
  padding: 10px;
  border-radius: 8px;
  background: var(--crepe-color-inline-area);
}

:deep(.milkdown .ProseMirror pre code) {
  display: block;
  padding: 0;
  background: transparent;
  color: var(--crepe-color-on-background);
}

:deep(.milkdown .ProseMirror blockquote) {
  position: relative;
  margin: 4px 0;
  padding: 0 0 0 40px;
  color: var(--crepe-color-on-surface-variant);
}

:deep(.milkdown .ProseMirror blockquote::before) {
  content: "";
  position: absolute;
  top: 4px;
  bottom: 4px;
  left: 0;
  width: 4px;
  border-radius: 999px;
  background: var(--crepe-color-selected);
}

:deep(.milkdown .ProseMirror hr) {
  height: 13px;
  padding: 6px 0;
  border: 0;
  background-color: color-mix(in srgb, var(--crepe-color-outline), transparent 80%);
  background-clip: content-box;
}

:deep(.milkdown .ProseMirror img) {
  max-width: 100%;
  vertical-align: bottom;
}

:deep(.milkdown .milkdown-table-block td),
:deep(.milkdown .milkdown-table-block th) {
  border-color: color-mix(in srgb, var(--crepe-color-outline), transparent 70%);
  padding: 6px 16px;
}

:deep(.milkdown .milkdown-toolbar),
:deep(.milkdown .milkdown-slash-menu),
:deep(.milkdown .milkdown-link-preview > .link-preview),
:deep(.milkdown .milkdown-link-edit > .link-edit) {
  border: 1px solid color-mix(in srgb, var(--crepe-color-outline), transparent 78%);
}

@media (max-width: 720px) {
  :deep(.milkdown .ProseMirror) {
    padding: 1.25rem 1rem;
  }

  :deep(.milkdown .ProseMirror h1) {
    font-size: 34px;
    line-height: 42px;
  }

  :deep(.milkdown .ProseMirror h2) {
    font-size: 30px;
    line-height: 38px;
  }

  :deep(.milkdown .ProseMirror h3) {
    font-size: 26px;
    line-height: 34px;
  }
}

:deep(.teamedit-mermaid-placeholder) {
  display: grid;
  justify-items: center;
  min-height: 6rem;
  padding: 1rem;
  border: 1px solid var(--border);
  border-radius: var(--radius-md);
  background: rgb(255 255 255 / 0.04);
}

:deep(.teamedit-mermaid-placeholder pre) {
  display: none;
}

:deep(.teamedit-mermaid-placeholder svg) {
  max-width: 100%;
  height: auto;
}

:deep(.teamedit-mermaid-message) {
  margin: 0;
  color: var(--muted);
}

:deep(.teamedit-mermaid-message--error) {
  color: var(--danger);
}
</style>
