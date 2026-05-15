import type { MindooDBAppAttachmentInfo } from "mindoodb-app-sdk";
import {
  AlignmentType,
  BorderStyle,
  Document,
  ExternalHyperlink,
  HeadingLevel,
  HighlightColor,
  ImageRun,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  UnderlineType,
  WidthType,
  type IParagraphOptions,
  type ParagraphChild,
} from "docx";

import { parseAttachmentMarkdownUrl } from "@/lib/attachmentImages";
import { DOCX_MIME_TYPE, createExportFileName, saveBlobToDisk } from "@/lib/exportMarkdown";
import { renderMarkdownFragment, type MarkdownRenderOptions } from "@/lib/markdownRendering";
import { MERMAID_PLACEHOLDER_CLASS, renderMermaidSvg } from "@/lib/mermaid";

const MAX_IMAGE_WIDTH = 520;
const DEFAULT_IMAGE_WIDTH = 420;
const DEFAULT_IMAGE_HEIGHT = 260;
const TABLE_WIDTH_PERCENT = 100;

type DocxBlock = Paragraph | Table;
type DocxImageType = "bmp" | "gif" | "jpg" | "png";
type TextStyle = {
  bold?: boolean;
  italics?: boolean;
  strike?: boolean;
  subScript?: boolean;
  superScript?: boolean;
  highlight?: (typeof HighlightColor)[keyof typeof HighlightColor];
  color?: string;
  font?: string;
};

interface ImageData {
  bytes: Uint8Array;
  type: DocxImageType;
  width: number;
  height: number;
}

export interface ExportDocxOptions {
  markdown: string;
  title: string;
  attachments?: readonly MindooDBAppAttachmentInfo[];
  resolveImageUrl?: MarkdownRenderOptions["resolveImageUrl"];
  renderMermaidSvg?: typeof renderMermaidSvg;
  svgToPngBytes?: (svg: string) => Promise<ImageData>;
  fetchImage?: (url: string) => Promise<ImageData | null>;
}

function normalizeText(value: string) {
  return value.replace(/\s+/g, " ");
}

function createTextRun(text: string, style: TextStyle = {}) {
  return new TextRun({
    text,
    bold: style.bold,
    italics: style.italics,
    strike: style.strike,
    subScript: style.subScript,
    superScript: style.superScript,
    highlight: style.highlight,
    color: style.color,
    font: style.font,
  });
}

function createCodeRun(text: string) {
  return new TextRun({
    text,
    font: "Consolas",
    size: 20,
    shading: {
      type: ShadingType.CLEAR,
      fill: "F1F5F9",
    },
  });
}

function emptyParagraph() {
  return new Paragraph({ children: [createTextRun("")] });
}

function paragraphFromRuns(children: ParagraphChild[], options: Omit<IParagraphOptions, "children"> = {}) {
  return new Paragraph({
    ...options,
    children: children.length > 0 ? children : [createTextRun("")],
  });
}

function isElement(node: Node): node is Element {
  return node.nodeType === Node.ELEMENT_NODE;
}

function isText(node: Node): node is Text {
  return node.nodeType === Node.TEXT_NODE;
}

function getElementText(element: Element) {
  return normalizeText(element.textContent ?? "").trim();
}

function headingLevelForTag(tagName: string) {
  switch (tagName.toLowerCase()) {
    case "h1":
      return HeadingLevel.HEADING_1;
    case "h2":
      return HeadingLevel.HEADING_2;
    case "h3":
      return HeadingLevel.HEADING_3;
    case "h4":
      return HeadingLevel.HEADING_4;
    case "h5":
      return HeadingLevel.HEADING_5;
    case "h6":
      return HeadingLevel.HEADING_6;
    default:
      return undefined;
  }
}

function alignmentForElement(element: Element) {
  if (element.classList.contains("teamedit-align--center")) {
    return AlignmentType.CENTER;
  }
  if (element.classList.contains("teamedit-align--right")) {
    return AlignmentType.RIGHT;
  }
  return undefined;
}

function imageTypeFromMime(mimeType: string): DocxImageType | null {
  const normalized = mimeType.toLowerCase();
  if (normalized.includes("png")) {
    return "png";
  }
  if (normalized.includes("jpeg") || normalized.includes("jpg")) {
    return "jpg";
  }
  if (normalized.includes("gif")) {
    return "gif";
  }
  if (normalized.includes("bmp")) {
    return "bmp";
  }
  return null;
}

function imageTypeFromUrl(url: string): DocxImageType | null {
  const path = url.split(/[?#]/, 1)[0].toLowerCase();
  if (path.endsWith(".png")) {
    return "png";
  }
  if (path.endsWith(".jpg") || path.endsWith(".jpeg")) {
    return "jpg";
  }
  if (path.endsWith(".gif")) {
    return "gif";
  }
  if (path.endsWith(".bmp")) {
    return "bmp";
  }
  return null;
}

function scaleImageSize(width: number, height: number) {
  if (!Number.isFinite(width) || width <= 0 || !Number.isFinite(height) || height <= 0) {
    return { width: DEFAULT_IMAGE_WIDTH, height: DEFAULT_IMAGE_HEIGHT };
  }
  if (width <= MAX_IMAGE_WIDTH) {
    return { width: Math.round(width), height: Math.round(height) };
  }
  const scale = MAX_IMAGE_WIDTH / width;
  return {
    width: MAX_IMAGE_WIDTH,
    height: Math.max(1, Math.round(height * scale)),
  };
}

function parseSvgSize(svg: string) {
  const document = new DOMParser().parseFromString(svg, "image/svg+xml");
  const root = document.documentElement;
  const width = Number.parseFloat(root.getAttribute("width") ?? "");
  const height = Number.parseFloat(root.getAttribute("height") ?? "");
  if (Number.isFinite(width) && Number.isFinite(height) && width > 0 && height > 0) {
    return scaleImageSize(width, height);
  }

  const viewBox = root.getAttribute("viewBox")?.trim().split(/\s+/).map(Number) ?? [];
  if (viewBox.length === 4 && Number.isFinite(viewBox[2]) && Number.isFinite(viewBox[3])) {
    return scaleImageSize(viewBox[2], viewBox[3]);
  }

  return scaleImageSize(DEFAULT_IMAGE_WIDTH, DEFAULT_IMAGE_HEIGHT);
}

async function blobToUint8Array(blob: Blob) {
  return new Uint8Array(await blob.arrayBuffer());
}

async function loadImageElement(url: string) {
  return await new Promise<HTMLImageElement>((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("Unable to load image for DOCX export."));
    image.src = url;
  });
}

async function canvasToPngBytes(canvas: HTMLCanvasElement) {
  const blob = await new Promise<Blob>((resolve, reject) => {
    canvas.toBlob((nextBlob) => {
      if (nextBlob) {
        resolve(nextBlob);
      } else {
        reject(new Error("Unable to encode image for DOCX export."));
      }
    }, "image/png");
  });
  return await blobToUint8Array(blob);
}

export async function svgToPngImageData(svg: string): Promise<ImageData> {
  const size = parseSvgSize(svg);
  const blob = new Blob([svg], { type: "image/svg+xml" });
  const url = URL.createObjectURL(blob);
  try {
    const image = await loadImageElement(url);
    const canvas = document.createElement("canvas");
    canvas.width = size.width;
    canvas.height = size.height;
    const context = canvas.getContext("2d");
    if (!context) {
      throw new Error("Canvas is unavailable for DOCX diagram export.");
    }
    context.drawImage(image, 0, 0, size.width, size.height);
    return {
      bytes: await canvasToPngBytes(canvas),
      type: "png",
      ...size,
    };
  } finally {
    URL.revokeObjectURL(url);
  }
}

async function fetchImageData(url: string): Promise<ImageData | null> {
  if (!url || url.startsWith("mindoodb-attachment:")) {
    return null;
  }

  const response = await fetch(url);
  if (!response.ok) {
    return null;
  }

  const blob = await response.blob();
  if (blob.type === "image/svg+xml") {
    return await svgToPngImageData(await blob.text());
  }

  const type = imageTypeFromMime(blob.type) ?? imageTypeFromUrl(url);
  if (!type) {
    return null;
  }

  const objectUrl = URL.createObjectURL(blob);
  try {
    const image = await loadImageElement(objectUrl);
    return {
      bytes: await blobToUint8Array(blob),
      type,
      ...scaleImageSize(image.naturalWidth, image.naturalHeight),
    };
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
}

function createImageRun(image: ImageData, altText: string) {
  return new ImageRun({
    type: image.type,
    data: image.bytes,
    transformation: {
      width: image.width,
      height: image.height,
    },
    altText: {
      title: altText || "Embedded image",
      description: altText || "Embedded image",
      name: altText || "Embedded image",
    },
  });
}

async function imageRunFromElement(
  element: HTMLImageElement | SVGSVGElement,
  options: Required<Pick<ExportDocxOptions, "fetchImage" | "svgToPngBytes">>,
) {
  const altText = element instanceof HTMLImageElement
    ? element.alt || element.title || "Image"
    : element.getAttribute("aria-label") || "Diagram";

  const imageData = element instanceof SVGSVGElement
    ? await options.svgToPngBytes(new XMLSerializer().serializeToString(element))
    : await options.fetchImage(element.currentSrc || element.src);

  return imageData ? createImageRun(imageData, altText) : createTextRun(`[${altText}]`);
}

async function convertInlineNodes(nodes: Iterable<Node>, options: Required<Pick<ExportDocxOptions, "fetchImage" | "svgToPngBytes">>, style: TextStyle = {}) {
  const runs: ParagraphChild[] = [];
  for (const node of nodes) {
    if (isText(node)) {
      const text = normalizeText(node.nodeValue ?? "");
      if (text) {
        runs.push(createTextRun(text, style));
      }
      continue;
    }
    if (!isElement(node)) {
      continue;
    }

    const tagName = node.tagName.toLowerCase();
    if (tagName === "br") {
      runs.push(new TextRun({ break: 1 }));
      continue;
    }
    if (tagName === "img") {
      runs.push(await imageRunFromElement(node as HTMLImageElement, options));
      continue;
    }
    if (tagName === "svg") {
      runs.push(await imageRunFromElement(node as SVGSVGElement, options));
      continue;
    }
    if (tagName === "a") {
      const href = node.getAttribute("href") ?? "";
      const children = await convertInlineNodes(node.childNodes, options, {
        ...style,
        color: "2563EB",
      });
      if (/^https?:\/\//i.test(href) || /^mailto:/i.test(href)) {
        runs.push(new ExternalHyperlink({
          link: href,
          children: children.length > 0 ? children : [createTextRun(href, { color: "2563EB" })],
        }));
      } else {
        runs.push(...children, createTextRun(href ? ` (${href})` : "", style));
      }
      continue;
    }

    const nextStyle: TextStyle = { ...style };
    if (tagName === "strong" || tagName === "b") {
      nextStyle.bold = true;
    } else if (tagName === "em" || tagName === "i") {
      nextStyle.italics = true;
    } else if (tagName === "del" || tagName === "s") {
      nextStyle.strike = true;
    } else if (tagName === "sub") {
      nextStyle.subScript = true;
    } else if (tagName === "sup") {
      nextStyle.superScript = true;
    } else if (tagName === "mark") {
      nextStyle.highlight = HighlightColor.YELLOW;
    } else if (tagName === "code") {
      runs.push(createCodeRun(node.textContent ?? ""));
      continue;
    }
    runs.push(...await convertInlineNodes(node.childNodes, options, nextStyle));
  }
  return runs;
}

async function paragraphFromElement(
  element: Element,
  options: Required<Pick<ExportDocxOptions, "fetchImage" | "svgToPngBytes">>,
  paragraphOptions: Omit<IParagraphOptions, "children"> = {},
) {
  return paragraphFromRuns(await convertInlineNodes(element.childNodes, options), {
    alignment: alignmentForElement(element),
    ...paragraphOptions,
  });
}

async function convertList(list: Element, options: Required<Pick<ExportDocxOptions, "fetchImage" | "svgToPngBytes">>) {
  const ordered = list.tagName.toLowerCase() === "ol";
  const blocks: DocxBlock[] = [];
  let itemIndex = 1;
  for (const item of Array.from(list.children).filter((child) => child.tagName.toLowerCase() === "li")) {
    const checked = item.getAttribute("data-teamedit-task-checked");
    const prefix = checked
      ? checked === "true" ? "[x] " : "[ ] "
      : ordered ? `${itemIndex}. ` : "";
    const runs = await convertInlineNodes(item.childNodes, options);
    if (prefix) {
      runs.unshift(createTextRun(prefix));
    }
    blocks.push(paragraphFromRuns(runs, ordered ? {} : { bullet: { level: 0 } }));
    itemIndex += 1;
  }
  return blocks;
}

async function convertTable(table: Element, options: Required<Pick<ExportDocxOptions, "fetchImage" | "svgToPngBytes">>) {
  const rows = Array.from(table.querySelectorAll("tr")).map(async (row) =>
    new TableRow({
      children: await Promise.all(Array.from(row.children).map(async (cell) =>
        new TableCell({
          children: [paragraphFromRuns(await convertInlineNodes(
            cell.childNodes,
            options,
            cell.tagName.toLowerCase() === "th" ? { bold: true } : {},
          ))],
        }))),
    }));

  return new Table({
    width: {
      size: TABLE_WIDTH_PERCENT,
      type: WidthType.PERCENTAGE,
    },
    rows: await Promise.all(rows),
  });
}

async function convertBlockElement(element: Element, options: Required<Pick<ExportDocxOptions, "fetchImage" | "svgToPngBytes">>): Promise<DocxBlock[]> {
  const tagName = element.tagName.toLowerCase();
  const headingLevel = headingLevelForTag(tagName);
  if (headingLevel) {
    return [await paragraphFromElement(element, options, { heading: headingLevel })];
  }
  if (tagName === "p") {
    return [await paragraphFromElement(element, options)];
  }
  if (tagName === "ul" || tagName === "ol") {
    return await convertList(element, options);
  }
  if (tagName === "blockquote") {
    return [await paragraphFromElement(element, options, {
      border: {
        left: { style: BorderStyle.SINGLE, size: 8, color: "CBD5E1" },
      },
      indent: { left: 360 },
    })];
  }
  if (tagName === "aside") {
    const title = element.querySelector(".teamedit-callout-title")?.textContent?.trim();
    const body = element.querySelector(".teamedit-callout-body") ?? element;
    return [
      paragraphFromRuns([createTextRun(title || "Note", { bold: true })], {
        shading: { type: ShadingType.CLEAR, fill: "F8FAFC" },
      }),
      ...await convertBlockNodes(body.childNodes, options),
    ];
  }
  if (tagName === "pre") {
    return [new Paragraph({
      shading: { type: ShadingType.CLEAR, fill: "F1F5F9" },
      children: [new TextRun({
        text: element.textContent ?? "",
        font: "Consolas",
        size: 20,
      })],
    })];
  }
  if (tagName === "table") {
    return [await convertTable(element, options)];
  }
  if (tagName === "hr") {
    return [new Paragraph({ thematicBreak: true })];
  }
  if (tagName === "img" || tagName === "svg") {
    return [paragraphFromRuns([await imageRunFromElement(element as HTMLImageElement | SVGSVGElement, options)])];
  }
  if (element.classList.contains(MERMAID_PLACEHOLDER_CLASS)) {
    const svg = element.querySelector("svg");
    if (svg) {
      return [paragraphFromRuns([await imageRunFromElement(svg, options)], { alignment: AlignmentType.CENTER })];
    }
  }

  const childBlocks = await convertBlockNodes(element.childNodes, options);
  if (childBlocks.length > 0) {
    return childBlocks;
  }
  const text = getElementText(element);
  return text ? [new Paragraph(text)] : [];
}

async function convertBlockNodes(nodes: Iterable<Node>, options: Required<Pick<ExportDocxOptions, "fetchImage" | "svgToPngBytes">>) {
  const blocks: DocxBlock[] = [];
  for (const node of nodes) {
    if (isText(node)) {
      const text = normalizeText(node.nodeValue ?? "").trim();
      if (text) {
        blocks.push(new Paragraph(text));
      }
      continue;
    }
    if (isElement(node)) {
      blocks.push(...await convertBlockElement(node, options));
    }
  }
  return blocks;
}

function isImageAttachment(attachment: MindooDBAppAttachmentInfo | undefined) {
  return attachment?.mimeType.toLowerCase().startsWith("image/") ?? false;
}

export function collectNonImageAttachmentReferences(markdown: string, attachments: readonly MindooDBAppAttachmentInfo[] = []) {
  const attachmentsByName = new Map<string, MindooDBAppAttachmentInfo>();
  for (const attachment of attachments) {
    attachmentsByName.set(attachment.fileName, attachment);
    attachmentsByName.set(attachment.attachmentId, attachment);
  }

  const references = new Map<string, { label: string; attachment: MindooDBAppAttachmentInfo }>();
  const linkPattern = /(?<!!)\[([^\]]*)]\((mindoodb-attachment:[^\s)"']+)(?:\s+"[^"]*")?\)/g;
  let match: RegExpExecArray | null;
  while ((match = linkPattern.exec(markdown)) !== null) {
    const label = match[1]?.trim() || "Attachment";
    const attachmentName = parseAttachmentMarkdownUrl(match[2]);
    const attachment = attachmentName ? attachmentsByName.get(attachmentName) : undefined;
    if (attachment && !isImageAttachment(attachment)) {
      references.set(attachment.attachmentId, { label, attachment });
    }
  }

  return [...references.values()];
}

function createAttachmentAppendix(references: ReturnType<typeof collectNonImageAttachmentReferences>) {
  if (references.length === 0) {
    return [];
  }

  return [
    new Paragraph({ text: "Document attachments", heading: HeadingLevel.HEADING_2 }),
    ...references.map(({ label, attachment }) =>
      new Paragraph({
        bullet: { level: 0 },
        children: [
          createTextRun(`${label}: `, { bold: true }),
          createTextRun(`${attachment.fileName || attachment.attachmentId} (${attachment.mimeType}, ${attachment.size} bytes)`),
        ],
      })),
  ];
}

async function hydrateMermaidPlaceholders(container: HTMLElement, renderSvg: typeof renderMermaidSvg) {
  const placeholders = Array.from(container.querySelectorAll<HTMLElement>(`.${MERMAID_PLACEHOLDER_CLASS}`));
  await Promise.all(placeholders.map(async (placeholder, index) => {
    const source = placeholder.querySelector("code")?.textContent ?? "";
    try {
      placeholder.innerHTML = await renderSvg(source, {
        idPrefix: `teamedit-docx-mermaid-${index}`,
        convertLabelsToSvgText: true,
      });
    } catch (error) {
      placeholder.textContent = error instanceof Error ? error.message : "Unable to render Mermaid diagram.";
    }
  }));
}

async function createExportContainer(options: ExportDocxOptions) {
  const container = document.createElement("div");
  container.innerHTML = renderMarkdownFragment(options.markdown, {
    resolveImageUrl: options.resolveImageUrl,
  });
  container.style.position = "fixed";
  container.style.left = "-10000px";
  container.style.top = "0";
  container.style.width = `${MAX_IMAGE_WIDTH}px`;
  document.body.append(container);
  await hydrateMermaidPlaceholders(container, options.renderMermaidSvg ?? renderMermaidSvg);
  return container;
}

export async function createDocxExportBlob(options: ExportDocxOptions) {
  const container = await createExportContainer(options);
  try {
    const converterOptions = {
      fetchImage: options.fetchImage ?? fetchImageData,
      svgToPngBytes: options.svgToPngBytes ?? svgToPngImageData,
    };
    const content = await convertBlockNodes(container.childNodes, converterOptions);
    const attachmentReferences = collectNonImageAttachmentReferences(options.markdown, options.attachments);
    const document = new Document({
      title: options.title,
      creator: "MindooDB TeamEdit",
      sections: [{
        children: [
          new Paragraph({ text: options.title || "Untitled document", heading: HeadingLevel.TITLE }),
          ...(content.length > 0 ? content : [emptyParagraph()]),
          ...createAttachmentAppendix(attachmentReferences),
        ],
      }],
    });
    const blob = await Packer.toBlob(document);
    return blob.type === DOCX_MIME_TYPE ? blob : new Blob([blob], { type: DOCX_MIME_TYPE });
  } finally {
    container.remove();
  }
}

export async function exportDocxFile(options: ExportDocxOptions) {
  return await saveBlobToDisk(
    await createDocxExportBlob(options),
    createExportFileName(options.title, "docx"),
  );
}
