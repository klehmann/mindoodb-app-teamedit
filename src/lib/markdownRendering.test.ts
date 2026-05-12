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

    expect(html).toContain("<pre><code class=\"hljs language-ts\">");
    expect(html).toContain("hljs-keyword");
    expect(html).toContain("value =");
    expect(html).not.toContain(MERMAID_PLACEHOLDER_CLASS);
  });

  it("syntax highlights supported code fences", () => {
    const html = renderMarkdownFragment("```markdown\n# Heading\n\nText with **strong** content.\n```");

    expect(html).toContain("<pre><code class=\"hljs language-markdown\">");
    expect(html).toContain("hljs-section");
    expect(html).toContain("hljs-strong");
  });

  it("renders inline LaTeX with KaTeX", () => {
    const html = renderMarkdownFragment("Energy: $E = mc^2$.");

    expect(html).toContain("class=\"katex\"");
    expect(html).toContain("E");
    expect(html).not.toContain("$E = mc^2$");
  });

  it("renders display LaTeX blocks with KaTeX", () => {
    const html = renderMarkdownFragment("$$\n\\int_{-\\infty}^{\\infty} e^{-x^2} dx = \\sqrt{\\pi}\n$$");

    expect(html).toContain("class=\"katex-display\"");
    expect(html).toContain("∞");
    expect(html).not.toContain("$$");
  });

  it("renders markdown task list items as disabled checkboxes", () => {
    const html = renderMarkdownFragment([
      "* Task List",
      "  * [ ] To-do item",
      "  * [x] Completed item",
    ].join("\n"));

    expect(html).toContain("<ul class=\"teamedit-task-list\">");
    expect(html).toContain("class=\"teamedit-task-list-item\" data-teamedit-task-checked=\"false\"");
    expect(html).toContain("class=\"teamedit-task-list-checkbox\" type=\"checkbox\" disabled>");
    expect(html).toContain("class=\"teamedit-task-list-item\" data-teamedit-task-checked=\"true\"");
    expect(html).toContain("class=\"teamedit-task-list-checkbox\" type=\"checkbox\" disabled checked>");
    expect(html).toContain("To-do item");
    expect(html).toContain("Completed item");
    expect(html).not.toContain("[ ] To-do item");
    expect(html).not.toContain("[x] Completed item");
  });

  it("renders GFM tables with alignment", () => {
    const html = renderMarkdownFragment([
      "| Left | Center | Right |",
      "| :--- | :----: | ----: |",
      "| a | b | c |",
    ].join("\n"));

    expect(html).toContain("<table>");
    expect(html).toContain("<th style=\"text-align:left\">Left</th>");
    expect(html).toContain("<th style=\"text-align:center\">Center</th>");
    expect(html).toContain("<th style=\"text-align:right\">Right</th>");
    expect(html).toContain("<td style=\"text-align:right\">c</td>");
  });

  it("renders GFM strikethrough and autolinks", () => {
    const html = renderMarkdownFragment("Visit https://mindoo.com and ~~remove this~~.");

    expect(html).toContain("<a href=\"https://mindoo.com\">https://mindoo.com</a>");
    expect(html).toContain("<del>remove this</del>");
    expect(html).not.toContain("~~remove this~~");
  });

  it("renders footnote references and definitions", () => {
    const html = renderMarkdownFragment([
      "This editor is a game-changer[^1].",
      "",
      "[^1]: Built by [DigitalPro](https://digitalpro.dev).",
    ].join("\n"));

    expect(html).toContain("class=\"footnote-ref\"");
    expect(html).toContain("href=\"#fn1\"");
    expect(html).toContain("<hr class=\"footnotes-sep\">");
    expect(html).toContain("<section class=\"footnotes\">");
    expect(html).toContain("Built by");
    expect(html).toContain("href=\"https://digitalpro.dev\"");
    expect(html).toContain("class=\"footnote-backref\"");
    expect(html).not.toContain("[^1]");
  });

  it("renders safe center-aligned div wrappers without enabling arbitrary HTML", () => {
    const html = renderMarkdownFragment([
      "<div align=\"center\">",
      "",
      "### Centered Title",
      "",
      "* **Centered** list item",
      "",
      "<script>alert('xss')</script>",
      "",
      "</div>",
    ].join("\n"));

    expect(html).toContain("<div class=\"teamedit-align teamedit-align--center\">");
    expect(html).toContain("<h3 id=\"centered-title\" class=\"teamedit-heading\">Centered Title</h3>");
    expect(html).toContain("<strong>Centered</strong> list item");
    expect(html).toContain("&lt;script&gt;alert");
    expect(html).not.toContain("<script>");
    expect(html).not.toContain("align=\"center\"");
  });

  it("renders safe left and right aligned div wrappers", () => {
    const html = renderMarkdownFragment([
      "<div align=\"left\">",
      "Left aligned",
      "</div>",
      "",
      "<div align=\"right\">",
      "**Right aligned**",
      "</div>",
    ].join("\n"));

    expect(html).toContain("<div class=\"teamedit-align teamedit-align--left\">");
    expect(html).toContain("Left aligned");
    expect(html).toContain("<div class=\"teamedit-align teamedit-align--right\">");
    expect(html).toContain("<strong>Right aligned</strong>");
    expect(html).not.toContain("align=\"left\"");
    expect(html).not.toContain("align=\"right\"");
  });

  it("adds stable, duplicate-aware heading anchors", () => {
    const html = renderMarkdownFragment([
      "## Roadmap",
      "",
      "### Roadmap",
      "",
      "## Roadmap!",
    ].join("\n"));

    expect(html).toContain("<h2 id=\"roadmap\" class=\"teamedit-heading\">Roadmap</h2>");
    expect(html).toContain("<h3 id=\"roadmap-1\" class=\"teamedit-heading\">Roadmap</h3>");
    expect(html).toContain("<h2 id=\"roadmap-2\" class=\"teamedit-heading\">Roadmap!</h2>");
  });

  it("renders highlight, subscript, and superscript inline formatting", () => {
    const html = renderMarkdownFragment("Use ==highlight==, H~2~O, and x^2^.");

    expect(html).toContain("<mark>highlight</mark>");
    expect(html).toContain("H<sub>2</sub>O");
    expect(html).toContain("x<sup>2</sup>");
  });

  it("normalizes Crepe-escaped formatting markers before rendering", () => {
    const html = renderMarkdownFragment("Use \\=\\=highlight\\=\\=, H\\~2\\~O, and x\\^2\\^.");

    expect(html).toContain("<mark>highlight</mark>");
    expect(html).toContain("H<sub>2</sub>O");
    expect(html).toContain("x<sup>2</sup>");
  });

  it("normalizes Crepe-escaped highlight marker pairs before rendering", () => {
    const html = renderMarkdownFragment("Use \\==highlight\\==.");

    expect(html).toContain("<mark>highlight</mark>");
  });

  it("renders callout containers with markdown content safely", () => {
    const html = renderMarkdownFragment([
      ":::warning Check this",
      "**Important:** review <script>alert('xss')</script> before saving.",
      ":::",
    ].join("\n"));

    expect(html).toContain("<aside class=\"teamedit-callout teamedit-callout--warning\">");
    expect(html).toContain("<p class=\"teamedit-callout-title\">Check this</p>");
    expect(html).toContain("<strong>Important:</strong>");
    expect(html).toContain("&lt;script&gt;alert");
    expect(html).not.toContain("<script>");
  });

  it("normalizes Crepe-escaped callout markers before rendering", () => {
    const html = renderMarkdownFragment([
      "\\:\\:\\:tip Heads up",
      "Callout body.",
      "\\:\\:\\:",
    ].join("\n"));

    expect(html).toContain("<aside class=\"teamedit-callout teamedit-callout--tip\">");
    expect(html).toContain("<p class=\"teamedit-callout-title\">Heads up</p>");
    expect(html).toContain("Callout body.");
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
    expect(renderMarkdownFragment(markdown)).toContain("hljs-tag");
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
