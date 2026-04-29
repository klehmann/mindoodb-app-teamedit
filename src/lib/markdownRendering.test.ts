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

  it("normalizes Milkdown br markers inside table cells before rendering", () => {
    const html = renderMarkdownFragment([
      "| a | <br /> | <br> |",
      "| - | - | - |",
      "| <br /> | value | <br/> |",
    ].join("\n"));

    expect(html).toContain("<table>");
    expect(html).toContain("<td>value</td>");
    expect(html).not.toContain("&lt;br");
  });

  it("does not normalize br markers inside fenced code blocks", () => {
    const markdown = "```html\n<br />\n```";

    expect(normalizeMarkdownForRendering(markdown)).toBe(markdown);
    expect(renderMarkdownFragment(markdown)).toContain("&lt;br /&gt;");
  });

  it("escapes document strings for HTML contexts", () => {
    expect(escapeHtml("<Team & Edit \"Doc\">")).toBe("&lt;Team &amp; Edit &quot;Doc&quot;&gt;");
  });

  it("translates Crepe ratio-encoded image alt text into a data attribute", () => {
    const html = renderMarkdownFragment("![0.50](mindoodb-attachment:image.png \"Team photo\")");

    expect(html).toContain("data-teamedit-image-ratio=\"0.50\"");
    expect(html).toContain("alt=\"Team photo\"");
    expect(html).not.toContain("alt=\"0.50\"");
  });

  it("strips numeric image alt text without a caption", () => {
    const html = renderMarkdownFragment("![0.75](mindoodb-attachment:image.png)");

    expect(html).toContain("data-teamedit-image-ratio=\"0.75\"");
    expect(html).toContain("alt=\"\"");
  });

  it("does not add a ratio attribute when the alt is the implicit ratio of 1", () => {
    const html = renderMarkdownFragment("![1.00](mindoodb-attachment:image.png \"Team photo\")");

    expect(html).not.toContain("data-teamedit-image-ratio");
    expect(html).toContain("alt=\"Team photo\"");
  });

  it("leaves non-numeric alt text untouched", () => {
    const html = renderMarkdownFragment("![A nice photo](mindoodb-attachment:image.png)");

    expect(html).not.toContain("data-teamedit-image-ratio");
    expect(html).toContain("alt=\"A nice photo\"");
  });
});
