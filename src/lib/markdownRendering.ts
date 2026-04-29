import MarkdownIt from "markdown-it";

import { isMermaidLanguage, MERMAID_PLACEHOLDER_CLASS } from "@/lib/mermaid";

export const MARKDOWN_CONTENT_CLASS = "teamedit-markdown-content";

type RenderRule = NonNullable<MarkdownIt["renderer"]["rules"][string]>;

export interface MarkdownRenderOptions {
  resolveImageUrl?: (url: string) => string;
}

export function escapeHtml(value: string) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll("\"", "&quot;");
}

export function createMermaidPlaceholderHtml(source: string) {
  return [
    `<figure class="teamedit-mermaid-figure">`,
    `<div class="${MERMAID_PLACEHOLDER_CLASS}" aria-label="Mermaid diagram">`,
    `<pre><code>${escapeHtml(source)}</code></pre>`,
    `</div>`,
    `</figure>`,
  ].join("");
}

export function normalizeMarkdownForRendering(markdown: string) {
  let inFence = false;
  let fenceMarker = "";

  return markdown
    .split(/\r?\n/)
    .map((line) => {
      const fenceMatch = line.match(/^\s*(```+|~~~+)/);
      if (fenceMatch) {
        const marker = fenceMatch[1][0];
        if (!inFence) {
          inFence = true;
          fenceMarker = marker;
        } else if (marker === fenceMarker) {
          inFence = false;
          fenceMarker = "";
        }
      }

      if (!inFence && /^\s*<br\s*\/?>\s*$/i.test(line)) {
        return "";
      }
      return line;
    })
    .join("\n");
}

export function createMarkdownRenderer(options: MarkdownRenderOptions = {}) {
  const renderer = new MarkdownIt({
    html: false,
    linkify: true,
    typographer: true,
  });

  const defaultImageRenderer: RenderRule = renderer.renderer.rules.image
    ?? ((tokens, index, renderOptions, _env, self) => self.renderToken(tokens, index, renderOptions));
  renderer.renderer.rules.image = (tokens, index, renderOptions, env, self) => {
    const token = tokens[index];
    const srcIndex = token.attrIndex("src");
    if (srcIndex >= 0 && token.attrs?.[srcIndex] && options.resolveImageUrl) {
      const src = token.attrs[srcIndex][1];
      token.attrs[srcIndex][1] = options.resolveImageUrl(src);
    }
    return defaultImageRenderer(tokens, index, renderOptions, env, self);
  };

  const defaultFenceRenderer: RenderRule = renderer.renderer.rules.fence
    ?? ((tokens, index, renderOptions, _env, self) => self.renderToken(tokens, index, renderOptions));
  renderer.renderer.rules.fence = (tokens, index, renderOptions, env, self) => {
    const token = tokens[index];
    const language = token.info.trim().split(/\s+/, 1)[0] ?? "";
    if (isMermaidLanguage(language)) {
      return createMermaidPlaceholderHtml(token.content);
    }
    return defaultFenceRenderer(tokens, index, renderOptions, env, self);
  };

  return renderer;
}

export function renderMarkdownFragment(markdown: string, options: MarkdownRenderOptions = {}) {
  return createMarkdownRenderer(options).render(normalizeMarkdownForRendering(markdown || ""));
}
