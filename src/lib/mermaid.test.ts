import { describe, expect, it } from "vitest";

import { convertForeignObjectLabelsToSvgText, sanitizeMermaidSvg } from "./mermaid";

describe("mermaid", () => {
  it("keeps Mermaid foreignObject label content after sanitizing", () => {
    const svg = sanitizeMermaidSvg(`
      <svg xmlns="http://www.w3.org/2000/svg">
        <g class="label">
          <foreignObject width="100" height="24">
            <div xmlns="http://www.w3.org/1999/xhtml" style="display: table-cell;">
              <span class="nodeLabel"><p>PDFFileInfo</p></span>
            </div>
          </foreignObject>
        </g>
        <script>alert("nope")</script>
      </svg>
    `);

    expect(svg).toContain("<foreignObject");
    expect(svg).toContain("PDFFileInfo");
    expect(svg).not.toContain("<script>");
  });

  it("converts Mermaid foreignObject labels to SVG text for Crepe previews", () => {
    const svg = convertForeignObjectLabelsToSvgText(sanitizeMermaidSvg(`
      <svg xmlns="http://www.w3.org/2000/svg">
        <g class="label" transform="translate(-50, -12)">
          <foreignObject width="100" height="24">
            <div xmlns="http://www.w3.org/1999/xhtml">
              <span class="nodeLabel"><p>Next step</p></span>
            </div>
          </foreignObject>
        </g>
      </svg>
    `));

    expect(svg).not.toContain("<foreignObject");
    expect(svg).toContain("<text");
    expect(svg).toContain("Next step");
  });
});
