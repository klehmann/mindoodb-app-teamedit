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

  // We replace markdown-it's default image rule entirely instead of decorating
  // it because the default writes `alt` from `renderInlineAsText(children)`
  // every render, which would clobber any rewrites we made above.
  renderer.renderer.rules.image = (tokens, index, renderOptions, _env, self) => {
    const token = tokens[index];
    const srcAttr = token.attrIndex("src");
    let src = srcAttr >= 0 && token.attrs?.[srcAttr] ? token.attrs[srcAttr][1] : "";
    if (src && options.resolveImageUrl) {
      src = options.resolveImageUrl(src);
    }
    const titleAttr = token.attrIndex("title");
    const titleValue = titleAttr >= 0 && token.attrs?.[titleAttr] ? token.attrs[titleAttr][1] : "";
    const sourceAlt = self.renderInlineAsText(token.children ?? [], renderOptions, _env);

    // Crepe's image-block schema repurposes the markdown image fields:
    //   ![<ratio>](<src> "<caption>")
    // - `alt` carries the resize ratio (e.g. "1.00", "0.50")
    // - `title` carries the visible caption
    // We translate that here so the rendered <img> stays accessible: the
    // numeric alt becomes a `data-teamedit-image-ratio` attribute that the
    // runtime image-scaling hooks consume, and the caption (title) takes
    // over as the real alt so screen readers and broken-image fallbacks
    // read the caption instead of a nonsense number.
    const ratioCandidate = Number.parseFloat(sourceAlt);
    const altLooksLikeRatio = sourceAlt.length > 0 && Number.isFinite(ratioCandidate);
    const finalAlt = altLooksLikeRatio ? titleValue.trim() : sourceAlt;
    const ratioAttr = altLooksLikeRatio
      && Number.isFinite(ratioCandidate)
      && ratioCandidate > 0
      && Math.abs(ratioCandidate - 1) > 1e-3
      ? ratioCandidate.toFixed(2)
      : null;

    const attrs: string[] = [];
    if (src) {
      attrs.push(`src="${escapeHtml(src)}"`);
    }
    attrs.push(`alt="${escapeHtml(finalAlt)}"`);
    if (titleValue) {
      attrs.push(`title="${escapeHtml(titleValue)}"`);
    }
    if (ratioAttr) {
      attrs.push(`data-teamedit-image-ratio="${ratioAttr}"`);
    }
    return `<img ${attrs.join(" ")}>`;
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
