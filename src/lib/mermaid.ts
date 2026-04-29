import DOMPurify from "dompurify";

export const MERMAID_LANGUAGE = "mermaid";
export const MERMAID_PLACEHOLDER_CLASS = "teamedit-mermaid-placeholder";

type MermaidModule = typeof import("mermaid");
type MermaidApi = MermaidModule["default"];

let mermaidLoader: Promise<MermaidApi> | null = null;
let renderCounter = 0;

const SVG_NAMESPACE = "http://www.w3.org/2000/svg";

export function isMermaidLanguage(language: string) {
  return language.trim().toLowerCase() === MERMAID_LANGUAGE;
}

async function getMermaid() {
  if (!mermaidLoader) {
    mermaidLoader = import("mermaid").then((module) => {
      module.default.initialize({
        startOnLoad: false,
        securityLevel: "strict",
        suppressErrorRendering: true,
        flowchart: {
          htmlLabels: true,
        },
      });
      return module.default;
    });
  }
  return await mermaidLoader;
}

function createRenderId(prefix = "teamedit-mermaid") {
  renderCounter += 1;
  return `${prefix}-${renderCounter}`;
}

export function sanitizeMermaidSvg(svg: string) {
  const sanitizedSvg = String(DOMPurify.sanitize(svg, {
    USE_PROFILES: {
      svg: true,
      svgFilters: true,
    },
    ADD_TAGS: ["foreignObject"],
    ADD_ATTR: ["xmlns"],
  }));

  if (typeof DOMParser === "undefined" || typeof XMLSerializer === "undefined") {
    return sanitizedSvg;
  }

  const parser = new DOMParser();
  const originalDocument = parser.parseFromString(svg, "image/svg+xml");
  const sanitizedDocument = parser.parseFromString(sanitizedSvg, "image/svg+xml");
  const originalLabels = Array.from(originalDocument.querySelectorAll("foreignObject"));
  const sanitizedLabels = Array.from(sanitizedDocument.querySelectorAll("foreignObject"));

  sanitizedLabels.forEach((label, index) => {
    const originalLabel = originalLabels[index];
    if (!originalLabel) {
      return;
    }

    const safeLabelHtml = DOMPurify.sanitize(originalLabel.innerHTML, {
      USE_PROFILES: {
        html: true,
      },
      ADD_ATTR: ["xmlns", "style", "class"],
    });
    label.innerHTML = safeLabelHtml;
  });

  return new XMLSerializer().serializeToString(sanitizedDocument.documentElement);
}

export function convertForeignObjectLabelsToSvgText(svg: string) {
  if (typeof DOMParser === "undefined" || typeof XMLSerializer === "undefined") {
    return svg;
  }

  const parser = new DOMParser();
  const document = parser.parseFromString(svg, "image/svg+xml");
  for (const label of Array.from(document.querySelectorAll("foreignObject"))) {
    const labelText = label.textContent?.replace(/\s+/g, " ").trim();
    if (!labelText) {
      label.remove();
      continue;
    }

    const width = Number.parseFloat(label.getAttribute("width") ?? "0");
    const height = Number.parseFloat(label.getAttribute("height") ?? "0");
    const text = document.createElementNS(SVG_NAMESPACE, "text");
    text.setAttribute("x", Number.isFinite(width) ? String(width / 2) : "0");
    text.setAttribute("y", Number.isFinite(height) ? String(height / 2) : "12");
    text.setAttribute("text-anchor", "middle");
    text.setAttribute("dominant-baseline", "middle");

    const tspan = document.createElementNS(SVG_NAMESPACE, "tspan");
    tspan.textContent = labelText;
    text.append(tspan);
    label.replaceWith(text);
  }

  return new XMLSerializer().serializeToString(document.documentElement);
}

export async function renderMermaidSvg(
  source: string,
  options: { idPrefix?: string; convertLabelsToSvgText?: boolean } = {},
) {
  const trimmed = source.trim();
  if (!trimmed) {
    throw new Error("Add Mermaid source to preview the diagram.");
  }

  const mermaid = await getMermaid();
  const { svg } = await mermaid.render(createRenderId(options.idPrefix), trimmed);
  const sanitizedSvg = sanitizeMermaidSvg(svg);
  return options.convertLabelsToSvgText ? convertForeignObjectLabelsToSvgText(sanitizedSvg) : sanitizedSvg;
}

function setMessage(target: HTMLElement, message: string, tone: "muted" | "error" = "muted") {
  target.replaceChildren();
  const paragraph = target.ownerDocument.createElement("p");
  paragraph.className = `teamedit-mermaid-message teamedit-mermaid-message--${tone}`;
  paragraph.textContent = message;
  target.append(paragraph);
}

export async function renderMermaidIntoElement(
  target: HTMLElement,
  source: string,
  options: { idPrefix?: string } = {},
) {
  setMessage(target, "Loading Mermaid preview...");
  try {
    target.innerHTML = await renderMermaidSvg(source, options);
  } catch (error) {
    setMessage(
      target,
      error instanceof Error ? error.message : "Unable to render Mermaid diagram.",
      "error",
    );
  }
}

export function createMermaidPreviewElement(source: string, options: { idPrefix?: string } = {}) {
  const element = document.createElement("div");
  element.className = MERMAID_PLACEHOLDER_CLASS;
  void renderMermaidIntoElement(element, source, options);
  return element;
}

export async function renderMermaidPlaceholders(
  root: ParentNode,
  options: { idPrefix?: string } = {},
) {
  const placeholders = Array.from(root.querySelectorAll<HTMLElement>(`.${MERMAID_PLACEHOLDER_CLASS}`));
  await Promise.all(placeholders.map(async (placeholder) => {
    const source = placeholder.querySelector("code")?.textContent ?? "";
    await renderMermaidIntoElement(placeholder, source, options);
  }));
}
