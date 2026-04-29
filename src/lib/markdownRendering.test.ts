import { describe, expect, it } from "vitest";

import {
  createMermaidPlaceholderHtml,
  escapeHtml,
  normalizeMarkdownForRendering,
  renderMarkdownFragment,
} from "./markdownRendering";
import { MERMAID_PLACEHOLDER_CLASS } from "./mermaid";

describe("markdownRendering", () => {
  it("renders Mermaid fences as hydratable placeholders", () => {
    const html = renderMarkdownFragment("```mermaid\nflowchart TD\n  A --> B\n```");

    expect(html).toContain(MERMAID_PLACEHOLDER_CLASS);
    expect(html).toContain("flowchart TD");
    expect(html).not.toContain("<pre><code class=\"language-mermaid\">");
  });

  it("preserves non-Mermaid code fences", () => {
    const html = renderMarkdownFragment("```ts\nconst value = 1;\n```");

    expect(html).toContain("<pre><code class=\"language-ts\">");
    expect(html).toContain("const value = 1;");
    expect(html).not.toContain(MERMAID_PLACEHOLDER_CLASS);
  });

  it("escapes Mermaid source before embedding it in placeholders", () => {
    const html = createMermaidPlaceholderHtml("flowchart TD\n  A[<script>] --> B");

    expect(html).toContain("&lt;script&gt;");
    expect(html).not.toContain("<script>");
  });

  it("rewrites image sources with the supplied resolver", () => {
    const html = renderMarkdownFragment("![alt](mindoodb-attachment:image.png)", {
      resolveImageUrl: (url) => `blob:${url}`,
    });

    expect(html).toContain("src=\"blob:mindoodb-attachment:image.png\"");
  });

  it("normalizes standalone Milkdown br markers before rendering", () => {
    const html = renderMarkdownFragment("Start writing together.\n<br />\n<br>\n\n| a | b |\n| - | - |");

    expect(html).not.toContain("&lt;br");
    expect(html).toContain("Start writing together.");
    expect(html).toContain("<table>");
  });

  it("does not normalize br markers inside fenced code blocks", () => {
    const markdown = "```html\n<br />\n```";

    expect(normalizeMarkdownForRendering(markdown)).toBe(markdown);
    expect(renderMarkdownFragment(markdown)).toContain("&lt;br /&gt;");
  });

  it("escapes document strings for HTML contexts", () => {
    expect(escapeHtml("<Team & Edit \"Doc\">")).toBe("&lt;Team &amp; Edit &quot;Doc&quot;&gt;");
  });
});
