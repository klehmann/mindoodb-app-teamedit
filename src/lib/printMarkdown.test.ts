import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { MARKDOWN_CONTENT_CLASS } from "./markdownRendering";
import { renderAndPrintMarkdownWindow } from "./printMarkdown";

const mocks = vi.hoisted(() => ({
  renderMermaidPlaceholders: vi.fn<() => Promise<void>>().mockResolvedValue(undefined),
}));

vi.mock("@/lib/mermaid", () => ({
  MERMAID_PLACEHOLDER_CLASS: "teamedit-mermaid-placeholder",
  isMermaidLanguage: (language: string) => language.trim().toLowerCase() === "mermaid",
  renderMermaidPlaceholders: mocks.renderMermaidPlaceholders,
}));

function createPrintWindow() {
  const printDocument = document.implementation.createHTMLDocument("");
  return {
    document: printDocument,
    focus: vi.fn(),
    print: vi.fn(),
    close: vi.fn(),
  } as unknown as Window & { focus: ReturnType<typeof vi.fn>; print: ReturnType<typeof vi.fn> };
}

describe("printMarkdown", () => {
  beforeEach(() => {
    mocks.renderMermaidPlaceholders.mockClear();
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("writes a formatted print document and calls print after hydration", async () => {
    const printWindow = createPrintWindow();

    await renderAndPrintMarkdownWindow(printWindow, {
      title: "Diagram Doc",
      markdown: "# Heading\n\n```mermaid\nflowchart TD\n  A --> B\n```\n\n![image](mindoodb-attachment:one.png)",
      resolveImageUrl: (url) => `blob:${url}`,
    });

    expect(printWindow.document.title).toBe("Diagram Doc");
    expect(printWindow.document.querySelector(`.${MARKDOWN_CONTENT_CLASS}`)?.innerHTML).toContain("teamedit-mermaid-placeholder");
    expect(printWindow.document.body.innerHTML).toContain("blob:mindoodb-attachment:one.png");
    expect(mocks.renderMermaidPlaceholders).toHaveBeenCalledWith(printWindow.document, {
      idPrefix: "teamedit-print-mermaid",
    });
    expect(printWindow.focus).toHaveBeenCalled();
    expect(printWindow.print).toHaveBeenCalled();
  });

  it("embeds resolved blob images before printing", async () => {
    const printWindow = createPrintWindow();
    const fetchImage = vi.fn<(_: string) => Promise<Response>>().mockResolvedValue(
      new Response(new Blob(["image-bytes"], { type: "image/png" })),
    );
    vi.stubGlobal("fetch", fetchImage);

    await renderAndPrintMarkdownWindow(printWindow, {
      title: "Pasted Image",
      markdown: "![image](mindoodb-attachment:pasted.png)",
      resolveImageUrl: () => "blob:http://127.0.0.1:4206/pasted-image",
    });

    const image = printWindow.document.querySelector("img");
    expect(fetchImage).toHaveBeenCalledWith("blob:http://127.0.0.1:4206/pasted-image");
    expect(image?.getAttribute("src")).toMatch(/^data:image\/png;base64,/);
    expect(printWindow.print).toHaveBeenCalled();
  });

  it("escapes the print title", async () => {
    const printWindow = createPrintWindow();

    await renderAndPrintMarkdownWindow(printWindow, {
      title: "<Unsafe>",
      markdown: "Body",
    });

    expect(printWindow.document.querySelector(".print-title")?.innerHTML).toBe("&lt;Unsafe&gt;");
  });
});
