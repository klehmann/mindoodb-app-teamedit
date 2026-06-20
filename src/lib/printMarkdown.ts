import { applyImageRatios } from "@/lib/imageRatio";
import highlightStyles from "highlight.js/styles/github.css?inline";
import katexStyles from "katex/dist/katex.min.css?inline";
import {
  escapeHtml,
  MARKDOWN_CONTENT_CLASS,
  renderMarkdownFragment,
  type MarkdownRenderOptions,
} from "@/lib/markdownRendering";
import { renderMermaidPlaceholders } from "@/lib/mermaid";

const PRINT_STYLES = `
${highlightStyles}

${katexStyles}

:root {
  color-scheme: light;
}

html,
body {
  margin: 0;
  padding: 0;
  background: #f8fafc;
  color: #0f172a;
  font-family: ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
}

body {
  padding: 2rem;
}

.print-shell {
  max-width: 52rem;
  margin: 0 auto;
  padding: 2.25rem;
  border-radius: 1rem;
  background: #ffffff;
  box-shadow: 0 24px 70px rgb(15 23 42 / 0.14);
}

.print-title {
  margin: 0 0 1.5rem;
  padding-bottom: 1rem;
  border-bottom: 1px solid #e2e8f0;
  font-size: 2rem;
  line-height: 1.15;
}

.${MARKDOWN_CONTENT_CLASS} {
  line-height: 1.65;
  font-size: 1rem;
}

.${MARKDOWN_CONTENT_CLASS} > :first-child {
  margin-top: 0;
}

.${MARKDOWN_CONTENT_CLASS} > :last-child {
  margin-bottom: 0;
}

.${MARKDOWN_CONTENT_CLASS} h1,
.${MARKDOWN_CONTENT_CLASS} h2,
.${MARKDOWN_CONTENT_CLASS} h3,
.${MARKDOWN_CONTENT_CLASS} h4,
.${MARKDOWN_CONTENT_CLASS} h5,
.${MARKDOWN_CONTENT_CLASS} h6 {
  line-height: 1.25;
  margin: 1.4rem 0 0.75rem;
  break-after: avoid;
}

.${MARKDOWN_CONTENT_CLASS} p,
.${MARKDOWN_CONTENT_CLASS} ul,
.${MARKDOWN_CONTENT_CLASS} ol,
.${MARKDOWN_CONTENT_CLASS} blockquote,
.${MARKDOWN_CONTENT_CLASS} pre,
.${MARKDOWN_CONTENT_CLASS} table,
.${MARKDOWN_CONTENT_CLASS} .teamedit-mermaid-figure {
  margin: 0.85rem 0;
}

.${MARKDOWN_CONTENT_CLASS} ul,
.${MARKDOWN_CONTENT_CLASS} ol {
  padding-left: 1.35rem;
}

.${MARKDOWN_CONTENT_CLASS} blockquote {
  margin-left: 0;
  padding: 0.85rem 1rem;
  border-left: 4px solid #cbd5e1;
  background: #f8fafc;
  border-radius: 0.75rem;
}

.${MARKDOWN_CONTENT_CLASS} code {
  font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
  font-size: 0.92em;
}

.${MARKDOWN_CONTENT_CLASS} :not(pre) > code {
  padding: 0.15rem 0.35rem;
  border-radius: 0.35rem;
  background: #f1f5f9;
}

.${MARKDOWN_CONTENT_CLASS} pre {
  overflow-x: auto;
  padding: 0.9rem 1rem;
  border: 1px solid #e2e8f0;
  border-radius: 0.8rem;
  background: #f1f5f9;
}

.${MARKDOWN_CONTENT_CLASS} pre code.hljs {
  padding: 0;
  background: transparent;
}

.${MARKDOWN_CONTENT_CLASS} .teamedit-align--left {
  text-align: left;
}

.${MARKDOWN_CONTENT_CLASS} .teamedit-align--center {
  text-align: center;
}

.${MARKDOWN_CONTENT_CLASS} .teamedit-align--right {
  text-align: right;
}

.${MARKDOWN_CONTENT_CLASS} .teamedit-align ul,
.${MARKDOWN_CONTENT_CLASS} .teamedit-align ol {
  display: inline-block;
  text-align: left;
}

.${MARKDOWN_CONTENT_CLASS} .teamedit-task-list {
  padding-left: 1.35rem;
}

.${MARKDOWN_CONTENT_CLASS} .teamedit-task-list-item {
  list-style: none;
}

.${MARKDOWN_CONTENT_CLASS} .teamedit-task-list-checkbox {
  width: 1rem;
  height: 1rem;
  margin: 0 0.5rem 0 -1.35rem;
  vertical-align: -0.15em;
  accent-color: #2563eb;
}

.${MARKDOWN_CONTENT_CLASS} .teamedit-heading {
  scroll-margin-top: 1.5rem;
}

.${MARKDOWN_CONTENT_CLASS} mark {
  padding: 0.05rem 0.25rem;
  border-radius: 0.3rem;
  background: #fef3c7;
  color: #78350f;
}

.${MARKDOWN_CONTENT_CLASS} sub,
.${MARKDOWN_CONTENT_CLASS} sup {
  line-height: 0;
}

.${MARKDOWN_CONTENT_CLASS} .footnotes {
  margin-top: 2rem;
  padding-top: 1rem;
  border-top: 1px solid #e2e8f0;
  color: #475569;
  font-size: 0.92em;
}

.${MARKDOWN_CONTENT_CLASS} .footnotes ol {
  padding-left: 1.35rem;
}

.${MARKDOWN_CONTENT_CLASS} .footnote-ref,
.${MARKDOWN_CONTENT_CLASS} .footnote-backref {
  font-weight: 600;
  text-decoration: none;
}

.${MARKDOWN_CONTENT_CLASS} .teamedit-callout {
  margin: 1rem 0;
  padding: 0.9rem 1rem;
  border: 1px solid #cbd5e1;
  border-left: 4px solid #2563eb;
  border-radius: 0.8rem;
  background: #f8fafc;
}

.${MARKDOWN_CONTENT_CLASS} .teamedit-callout--tip {
  border-left-color: #059669;
}

.${MARKDOWN_CONTENT_CLASS} .teamedit-callout--warning,
.${MARKDOWN_CONTENT_CLASS} .teamedit-callout--caution {
  border-left-color: #d97706;
}

.${MARKDOWN_CONTENT_CLASS} .teamedit-callout--danger {
  border-left-color: #dc2626;
}

.${MARKDOWN_CONTENT_CLASS} .teamedit-callout-title {
  margin: 0 0 0.5rem;
  font-weight: 700;
}

.${MARKDOWN_CONTENT_CLASS} .teamedit-callout-body > :first-child {
  margin-top: 0;
}

.${MARKDOWN_CONTENT_CLASS} .teamedit-callout-body > :last-child {
  margin-bottom: 0;
}

.${MARKDOWN_CONTENT_CLASS} table {
  width: 100%;
  min-width: max(100%, 28rem);
  table-layout: fixed;
  border-collapse: collapse;
  border: 1px solid #cbd5e1;
}

.${MARKDOWN_CONTENT_CLASS} th,
.${MARKDOWN_CONTENT_CLASS} td {
  min-width: 7rem;
  height: 2rem;
  padding: 4px 16px;
  text-align: left;
  vertical-align: top;
  border: 1px solid #cbd5e1;
  overflow-wrap: anywhere;
}

.${MARKDOWN_CONTENT_CLASS} th {
  background: #f8fafc;
  font-weight: 600;
}

.${MARKDOWN_CONTENT_CLASS} img,
.${MARKDOWN_CONTENT_CLASS} svg {
  max-width: 100%;
  height: auto;
}

.teamedit-mermaid-figure {
  break-inside: avoid;
}

.teamedit-mermaid-placeholder {
  display: grid;
  justify-items: center;
  min-height: 6rem;
  padding: 1rem;
  border: 1px solid #e2e8f0;
  border-radius: 0.85rem;
  background: #ffffff;
}

.teamedit-mermaid-placeholder pre {
  display: none;
}

.teamedit-mermaid-message {
  margin: 0;
  color: #64748b;
}

.teamedit-mermaid-message--error {
  color: #b91c1c;
}

@page {
  margin: 0.65in;
}

@media print {
  body {
    padding: 0;
    background: #ffffff;
  }

  .print-shell {
    max-width: none;
    margin: 0;
    padding: 0;
    box-shadow: none;
  }
}
`;

export interface PrintMarkdownOptions extends MarkdownRenderOptions {
  title: string;
  markdown: string;
}

function createPrintDocument(title: string, bodyHtml: string) {
  const safeTitle = escapeHtml(title.trim() || "Untitled document");
  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>${safeTitle}</title>
    <style>${PRINT_STYLES}</style>
  </head>
  <body>
    <main class="print-shell">
      <h1 class="print-title">${safeTitle}</h1>
      <article class="${MARKDOWN_CONTENT_CLASS}">${bodyHtml}</article>
    </main>
  </body>
</html>`;
}

async function waitForImages(document: Document, timeoutMs = 1500) {
  const images = Array.from(document.images).filter((image) => !image.complete);
  await Promise.all(images.map((image) =>
    new Promise<void>((resolve) => {
      const timeoutId = window.setTimeout(resolve, timeoutMs);
      const finish = () => {
        window.clearTimeout(timeoutId);
        resolve();
      };
      image.addEventListener("load", finish, { once: true });
      image.addEventListener("error", finish, { once: true });
    })));
}

async function blobToDataUrl(blob: Blob) {
  return await new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error ?? new Error("Unable to read image for print."));
    reader.readAsDataURL(blob);
  });
}

async function inlineBlobImages(document: Document) {
  const images = Array.from(document.images).filter((image) => image.src.startsWith("blob:"));
  await Promise.all(images.map(async (image) => {
    try {
      image.src = await blobToDataUrl(await fetch(image.src).then((response) => response.blob()));
    } catch {
      // Leave the original URL in place; the existing load/error wait below will
      // still let printing continue if the browser can resolve it on its own.
    }
  }));
}

async function waitForPrintAssets(document: Document) {
  await Promise.all([
    "fonts" in document ? document.fonts.ready.catch(() => undefined) : undefined,
    waitForImages(document),
  ]);
}

export async function renderAndPrintMarkdownWindow(printWindow: Window, options: PrintMarkdownOptions) {
  const bodyHtml = renderMarkdownFragment(options.markdown, {
    resolveImageUrl: options.resolveImageUrl,
  });

  printWindow.document.open();
  printWindow.document.write(createPrintDocument(options.title, bodyHtml));
  printWindow.document.close();

  await renderMermaidPlaceholders(printWindow.document, {
    idPrefix: "teamedit-print-mermaid",
  });
  await inlineBlobImages(printWindow.document);
  // Set up ratio scaling listeners up front so each image is resized as soon
  // as it loads. waitForPrintAssets() below resolves only after every image
  // has fired its load/error event, by which point the heights are in place.
  applyImageRatios(printWindow.document);
  await waitForPrintAssets(printWindow.document);

  printWindow.focus();
  printWindow.print();
}
