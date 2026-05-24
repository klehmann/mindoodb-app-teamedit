/**
 * Schema adapter that bridges the `@eigenpal/docx-editor-core` ProseMirror
 * schema to the Automerge rich-text model.
 */

import { SchemaAdapter, type MappedMarkSpec, type MappedNodeSpec } from "@automerge/prosemirror";
import type { MaterializeValue } from "@automerge/automerge";
import { schema as docxSchema } from "@eigenpal/docx-editor-core/prosemirror/schema";
import type { Mark, MarkSpec, Node, NodeSpec, NodeType, MarkType } from "prosemirror-model";

type SpecMap<T> = {
  forEach: (callback: (name: string, spec: T) => void) => void;
};

const STRUCTURAL_BLOCK_NODES = ["textBox", "table", "tableRow", "tableCell", "tableHeader"];
const EMBED_NODES = ["field", "hardBreak", "image", "math", "shape", "tab"];

let cachedAdapter: SchemaAdapter | null = null;

export function createDocxAutomergeSchemaAdapter(): SchemaAdapter {
  if (cachedAdapter) {
    return cachedAdapter;
  }

  const adapter = new SchemaAdapter({
    nodes: cloneNodes(),
    marks: cloneMarks(),
  });

  // docx-editor-core and @automerge/prosemirror may resolve different
  // prosemirror-model copies; retargeting uses the shared docx schema at runtime.
  const retargeted = adapter as SchemaAdapter & {
    schema: typeof docxSchema;
    nodeMappings: SchemaAdapter["nodeMappings"];
    markMappings: SchemaAdapter["markMappings"];
    unknownBlock: NodeType;
  };

  retargeted.schema = docxSchema;
  retargeted.nodeMappings = adapter.nodeMappings.map((mapping) => ({
    ...mapping,
    outer: retargetNodeType(mapping.outer),
    content: retargetNodeType(mapping.content) ?? mapping.content,
  })) as SchemaAdapter["nodeMappings"];
  retargeted.markMappings = adapter.markMappings.map((mapping) => ({
    ...mapping,
    prosemirrorMark: retargetMarkType(mapping.prosemirrorMark),
  })) as SchemaAdapter["markMappings"];
  retargeted.unknownBlock = docxSchema.nodes.paragraph ?? adapter.unknownBlock;

  cachedAdapter = retargeted;
  return retargeted;
}

function cloneNodes(): Record<string, MappedNodeSpec> {
  const nodes: Record<string, MappedNodeSpec> = {};

  (docxSchema.spec.nodes as SpecMap<NodeSpec>).forEach((name, spec) => {
    nodes[name] = { ...spec };
  });

  STRUCTURAL_BLOCK_NODES.forEach((name) => mapBlockNode(nodes, name));
  mapBlockNode(nodes, "paragraph", { unknownBlock: true });
  EMBED_NODES.forEach((name) => mapEmbedNode(nodes, name));

  return nodes;
}

function mapBlockNode(
  nodes: Record<string, MappedNodeSpec>,
  name: string,
  options: { unknownBlock?: boolean } = {},
): void {
  if (!nodes[name]) return;

  nodes[name] = {
    ...nodes[name],
    automerge: {
      block: name,
      ...(options.unknownBlock ? { unknownBlock: true } : {}),
      attrParsers: {
        fromProsemirror: serializableAttrs,
        fromAutomerge: (block) => block.attrs,
      },
    },
  };
}

function mapEmbedNode(nodes: Record<string, MappedNodeSpec>, name: string): void {
  if (!nodes[name]) return;

  nodes[name] = {
    ...nodes[name],
    automerge: {
      block: name,
      isEmbed: true,
      attrParsers: {
        fromProsemirror: serializableAttrs,
        fromAutomerge: (block) => block.attrs,
      },
    },
  };
}

function serializableAttrs(node: Node): Record<string, MaterializeValue> {
  return Object.fromEntries(
    Object.entries(node.attrs).filter(([, value]) => value !== null && value !== undefined),
  ) as Record<string, MaterializeValue>;
}

function cloneMarks(): Record<string, MappedMarkSpec> {
  const marks: Record<string, MappedMarkSpec> = {};

  (docxSchema.spec.marks as SpecMap<MarkSpec>).forEach((name, spec) => {
    marks[name] = {
      ...spec,
      automerge: {
        markName: name,
        parsers: {
          fromAutomerge: attrsFromAutomerge,
          fromProsemirror: valueFromProsemirror,
        },
      },
    };
  });

  return marks;
}

function valueFromProsemirror(mark: Mark): boolean | string {
  const attrs = Object.fromEntries(
    Object.entries(mark.attrs).filter(([, value]) => value !== null && value !== undefined),
  );

  return Object.keys(attrs).length === 0 ? true : JSON.stringify(attrs);
}

function attrsFromAutomerge(value: unknown): Record<string, unknown> {
  if (typeof value !== "string") return {};

  try {
    const parsed = JSON.parse(value);
    return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed : {};
  } catch {
    return {};
  }
}

function retargetNodeType(nodeType: NodeType | null): NodeType | null {
  if (!nodeType) return null;
  return docxSchema.nodes[nodeType.name] ?? nodeType;
}

function retargetMarkType(markType: MarkType): MarkType {
  return docxSchema.marks[markType.name] ?? markType;
}
