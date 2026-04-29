import MarkdownIt from "markdown-it";
import hljs from "highlight.js/lib/common";
import katex from "katex";
import markdownItFootnote from "markdown-it-footnote";

import { isMermaidLanguage, MERMAID_PLACEHOLDER_CLASS } from "@/lib/mermaid";

export const MARKDOWN_CONTENT_CLASS = "teamedit-markdown-content";

type StateCore = Parameters<MarkdownIt["core"]["ruler"]["push"]>[1] extends (state: infer State, ...args: never[]) => unknown ? State : never;
type StateBlock = Parameters<MarkdownIt["block"]["ruler"]["push"]>[1] extends (state: infer State, ...args: never[]) => unknown ? State : never;
type StateInline = Parameters<MarkdownIt["inline"]["ruler"]["push"]>[1] extends (state: infer State, ...args: never[]) => unknown ? State : never;

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

function renderLatex(source: string, displayMode: boolean) {
  return katex.renderToString(source, {
    throwOnError: false,
    displayMode,
  });
}

function normalizeHighlightLanguage(language: string) {
  const normalized = language.trim().toLowerCase();
  if (normalized === "md") {
    return "markdown";
  }
  return normalized;
}

function renderHighlightedCode(source: string, language: string) {
  const normalizedLanguage = normalizeHighlightLanguage(language);
  const languageClass = normalizedLanguage ? ` language-${escapeHtml(normalizedLanguage)}` : "";
  let code = escapeHtml(source);

  if (normalizedLanguage && hljs.getLanguage(normalizedLanguage)) {
    try {
      code = hljs.highlight(source, {
        language: normalizedLanguage,
        ignoreIllegals: true,
      }).value;
    } catch {
      code = escapeHtml(source);
    }
  }

  return `<pre><code class="hljs${languageClass}">${code}</code></pre>\n`;
}

function attrJoinUnique(token: StateCore["tokens"][number], name: string, value: string) {
  const existing = token.attrGet(name)?.split(/\s+/).filter(Boolean) ?? [];
  if (!existing.includes(value)) {
    token.attrJoin(name, value);
  }
}

function taskListRule(state: StateCore) {
  const listStack: number[] = [];

  for (let index = 0; index < state.tokens.length; index += 1) {
    const token = state.tokens[index];
    if (token.type === "bullet_list_open" || token.type === "ordered_list_open") {
      listStack.push(index);
      continue;
    }
    if (token.type === "bullet_list_close" || token.type === "ordered_list_close") {
      listStack.pop();
      continue;
    }
    if (token.type !== "list_item_open") {
      continue;
    }

    const inlineToken = state.tokens.slice(index + 1).find((candidate) =>
      candidate.type === "inline" && candidate.level > token.level);
    const match = inlineToken?.content.match(/^\[( |x|X)\]\s+/);
    if (!inlineToken || !match) {
      continue;
    }

    const checked = match[1].toLowerCase() === "x";
    const markerLength = match[0].length;
    inlineToken.content = inlineToken.content.slice(markerLength);
    inlineToken.children?.some((child) => {
      if (child.type !== "text") {
        return false;
      }
      child.content = child.content.slice(markerLength);
      return true;
    });

    attrJoinUnique(token, "class", "teamedit-task-list-item");
    token.attrSet("data-teamedit-task-checked", checked ? "true" : "false");
    const parentListIndex = listStack.at(-1);
    if (parentListIndex !== undefined) {
      attrJoinUnique(state.tokens[parentListIndex], "class", "teamedit-task-list");
    }
  }
}

function mathBlockRule(state: StateBlock, startLine: number, endLine: number, silent: boolean) {
  const start = state.bMarks[startLine] + state.tShift[startLine];
  const max = state.eMarks[startLine];
  const firstLine = state.src.slice(start, max);
  if (!firstLine.startsWith("$$")) {
    return false;
  }

  let content = firstLine.slice(2);
  let nextLine = startLine;
  const sameLineEnd = content.indexOf("$$");
  if (sameLineEnd >= 0) {
    content = content.slice(0, sameLineEnd);
  } else {
    const lines: string[] = [content];
    for (nextLine = startLine + 1; nextLine < endLine; nextLine += 1) {
      const lineStart = state.bMarks[nextLine] + state.tShift[nextLine];
      const lineEnd = state.eMarks[nextLine];
      const line = state.src.slice(lineStart, lineEnd);
      const endIndex = line.indexOf("$$");
      if (endIndex >= 0) {
        lines.push(line.slice(0, endIndex));
        break;
      }
      lines.push(line);
    }
    if (nextLine >= endLine) {
      return false;
    }
    content = lines.join("\n");
  }

  if (silent) {
    return true;
  }

  const token = state.push("math_block", "div", 0);
  token.block = true;
  token.content = content.trim();
  token.map = [startLine, nextLine + 1];
  state.line = nextLine + 1;
  return true;
}

function findInlineMathEnd(source: string, start: number) {
  for (let index = start; index < source.length; index += 1) {
    if (source[index] === "$" && source[index - 1] !== "\\" && source[index + 1] !== "$") {
      return index;
    }
  }
  return -1;
}

function mathInlineRule(state: StateInline, silent: boolean) {
  const start = state.pos;
  const source = state.src;
  if (source[start] !== "$" || source[start + 1] === "$" || source[start - 1] === "\\") {
    return false;
  }

  const end = findInlineMathEnd(source, start + 1);
  if (end < 0) {
    return false;
  }

  const content = source.slice(start + 1, end);
  if (!content.trim()) {
    return false;
  }

  if (!silent) {
    const token = state.push("math_inline", "span", 0);
    token.content = content;
  }
  state.pos = end + 1;
  return true;
}

function alignDivBlockRule(state: StateBlock, startLine: number, endLine: number, silent: boolean) {
  const start = state.bMarks[startLine] + state.tShift[startLine];
  const max = state.eMarks[startLine];
  const firstLine = state.src.slice(start, max);
  const match = firstLine.match(/^<div\s+align=(?:"(left|center|right)"|'(left|center|right)'|(left|center|right))\s*>\s*$/i);
  if (!match) {
    return false;
  }
  const alignment = (match[1] ?? match[2] ?? match[3]).toLowerCase();

  let nextLine = startLine + 1;
  for (; nextLine < endLine; nextLine += 1) {
    const lineStart = state.bMarks[nextLine] + state.tShift[nextLine];
    const lineEnd = state.eMarks[nextLine];
    const line = state.src.slice(lineStart, lineEnd);
    if (/^<\/div>\s*$/i.test(line)) {
      break;
    }
  }
  if (nextLine >= endLine) {
    return false;
  }

  if (silent) {
    return true;
  }

  const token = state.push("teamedit_align_center", "div", 0);
  token.block = true;
  token.content = state.getLines(startLine + 1, nextLine, state.blkIndent, false);
  token.attrSet("data-teamedit-align", alignment);
  token.map = [startLine, nextLine + 1];
  state.line = nextLine + 1;
  return true;
}

function normalizeTableBreakMarkers(line: string) {
  if (!line.includes("|") || !/<br\s*\/?>/i.test(line)) {
    return line;
  }

  return line
    .split("|")
    .map((cell) => /^\s*<br\s*\/?>\s*$/i.test(cell) ? "" : cell)
    .join("|");
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
      return inFence ? line : normalizeTableBreakMarkers(line);
    })
    .join("\n");
}

export function createMarkdownRenderer(options: MarkdownRenderOptions = {}) {
  const renderer = new MarkdownIt({
    html: false,
    linkify: true,
    typographer: true,
  });
  renderer.use(markdownItFootnote);

  renderer.block.ruler.before("paragraph", "teamedit_align_center", alignDivBlockRule, {
    alt: ["paragraph", "reference", "blockquote", "list"],
  });
  renderer.block.ruler.before("fence", "math_block", mathBlockRule, {
    alt: ["paragraph", "reference", "blockquote", "list"],
  });
  renderer.inline.ruler.after("escape", "math_inline", mathInlineRule);
  renderer.core.ruler.after("inline", "teamedit_task_lists", taskListRule);

  renderer.renderer.rules.math_block = (tokens, index) =>
    `<div class="katex-display">${renderLatex(tokens[index].content, true)}</div>\n`;
  renderer.renderer.rules.math_inline = (tokens, index) => renderLatex(tokens[index].content, false);
  renderer.renderer.rules.teamedit_align_center = (tokens, index, renderOptions, env) => {
    const alignment = tokens[index].attrGet("data-teamedit-align") ?? "center";
    return `<div class="teamedit-align teamedit-align--${alignment}">\n${renderer.render(normalizeMarkdownForRendering(tokens[index].content), env)}\n</div>\n`;
  };
  renderer.renderer.rules.list_item_open = (tokens, index, renderOptions, env, self) => {
    const token = tokens[index];
    const checked = token.attrGet("data-teamedit-task-checked");
    const checkbox = checked
      ? `<input class="teamedit-task-list-checkbox" type="checkbox" disabled${checked === "true" ? " checked" : ""}>`
      : "";
    return `${self.renderToken(tokens, index, renderOptions)}${checkbox}`;
  };

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

  renderer.renderer.rules.fence = (tokens, index) => {
    const token = tokens[index];
    const language = token.info.trim().split(/\s+/, 1)[0] ?? "";
    if (isMermaidLanguage(language)) {
      return createMermaidPlaceholderHtml(token.content);
    }
    return renderHighlightedCode(token.content, language);
  };

  return renderer;
}

export function renderMarkdownFragment(markdown: string, options: MarkdownRenderOptions = {}) {
  return createMarkdownRenderer(options).render(normalizeMarkdownForRendering(markdown || ""));
}
