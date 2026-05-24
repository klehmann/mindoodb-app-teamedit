import * as A from "@automerge/automerge";
import type {
  MindooDBAppAutomergeSnapshot,
  MindooDBAppDatabase,
  MindooDBAppDocument,
  MindooDBAppRichTextSpan,
} from "mindoodb-app-sdk";
import { createDocxAutomergeSchemaAdapter } from "@/editors/word/lib/automergeSchemaAdapter";
import {
  dehydrateRichTextSpans,
  reviveRichTextSpans,
  updateSpansWithDiagnostics,
} from "@/editors/word/lib/automergeSpanUtils";

export interface CreateWordAutomergeHandleOptions {
  database: MindooDBAppDatabase;
  document: MindooDBAppDocument;
  path: Array<string | number>;
  snapshot: MindooDBAppAutomergeSnapshot;
  actor?: string;
}

export interface WordAutomergeHandleFlushResult {
  document: MindooDBAppDocument;
  spans: MindooDBAppRichTextSpan[];
  reconciled: boolean;
}

export type WordAutomergeHandleChangeListener = (snapshot: {
  spans: MindooDBAppRichTextSpan[];
}) => void;

const adapter = createDocxAutomergeSchemaAdapter();

export class WordAutomergeHandle {
  private doc: A.Doc<Record<string, unknown>>;
  private baseHeads: string[];
  private dirtyInternal = false;
  private listeners = new Set<WordAutomergeHandleChangeListener>();

  constructor(private readonly options: CreateWordAutomergeHandleOptions) {
    this.doc = A.load<Record<string, unknown>>(
      options.snapshot.binary,
      options.actor ? { actor: options.actor } : undefined,
    );
    this.baseHeads = [...options.snapshot.heads];
  }

  get spans() {
    return readRichTextSpans(this.doc, this.options.path);
  }

  get dirty() {
    return this.dirtyInternal;
  }

  get baseHeadsSnapshot() {
    return [...this.baseHeads];
  }

  on(event: "change", listener: WordAutomergeHandleChangeListener) {
    if (event === "change") {
      this.listeners.add(listener);
    }
  }

  off(event: "change", listener: WordAutomergeHandleChangeListener) {
    if (event === "change") {
      this.listeners.delete(listener);
    }
  }

  replaceSpans(spans: MindooDBAppRichTextSpan[]) {
    this.applySpansToLocalDoc(spans);
    this.emitChange();
  }

  private applySpansToLocalDoc(spans: MindooDBAppRichTextSpan[]) {
    const nativeSpans = reviveRichTextSpans(spans);
    this.doc = A.change(this.doc, (draft) => {
      updateSpansWithDiagnostics(
        draft,
        this.options.path as A.Prop[],
        nativeSpans,
        adapter.updateSpansConfig(),
        "word editor",
      );
    });
    this.dirtyInternal = true;
  }

  syncDocument(document: MindooDBAppDocument) {
    this.options.document = document;
  }

  async flush(options?: {
    spans?: MindooDBAppRichTextSpan[];
  }): Promise<WordAutomergeHandleFlushResult> {
    if (options?.spans) {
      this.applySpansToLocalDoc(options.spans);
    }

    if (!this.dirtyInternal) {
      return {
        document: this.options.document,
        spans: this.spans,
        reconciled: false,
      };
    }

    const changes = A.getChangesSince(this.doc, this.baseHeads as A.Heads);
    if (changes.length === 0) {
      this.dirtyInternal = false;
      return {
        document: this.options.document,
        spans: this.spans,
        reconciled: false,
      };
    }

    const localSpansAtFlush = this.spans;

    const replicaHeads = A.getHeads(this.doc) as A.Heads;

    const result = await this.options.database.documents.applyAutomergeChanges(
      this.options.document.id,
      {
        baseHeads: this.baseHeads,
        replicaHeads: [...replicaHeads],
        changes,
      },
    );

    this.reconcileLocalReplicaAfterApply(result, replicaHeads);
    this.baseHeads = [...result.heads];
    this.dirtyInternal = false;
    this.options.document = result.document;
    const mergedSpans = this.spans;
    const reconciled = JSON.stringify(localSpansAtFlush) !== JSON.stringify(mergedSpans);

    return {
      document: result.document,
      spans: mergedSpans,
      reconciled,
    };
  }

  async refresh() {
    return this.options.database.documents.getAutomergeSnapshot(
      this.options.document.id,
    );
  }

  reconcile(document: MindooDBAppDocument, snapshot: MindooDBAppAutomergeSnapshot) {
    const previous = JSON.stringify(this.spans);
    this.doc = A.load(snapshot.binary);
    this.baseHeads = [...snapshot.heads];
    this.dirtyInternal = false;
    this.options.document = document;
    const changed = JSON.stringify(this.spans) !== previous;
    if (changed) {
      this.emitChange();
    }
    return changed;
  }

  private async reconcileLocalReplicaAfterApply(
    result: Awaited<
      ReturnType<MindooDBAppDatabase["documents"]["applyAutomergeChanges"]>
    >,
    replicaHeadsBeforeApply: A.Heads,
  ) {
    if (result.changesSince) {
      if (result.changesSince.changes.length > 0) {
        [this.doc] = A.applyChanges(
          this.doc,
          result.changesSince.changes.map((change) => new Uint8Array(change)),
        );
      }
      return;
    }

    if (!automergeHeadsEqual(replicaHeadsBeforeApply, result.heads)) {
      const snapshot = await this.refresh();
      this.doc = A.load(snapshot.binary);
    }
  }

  private emitChange() {
    const snapshot = { spans: this.spans };
    for (const listener of this.listeners) {
      listener(snapshot);
    }
  }
}

export function createWordAutomergeHandle(
  options: CreateWordAutomergeHandleOptions,
): WordAutomergeHandle {
  return new WordAutomergeHandle(options);
}

export function readRichTextSpansFromAutomergeSnapshot(
  snapshot: MindooDBAppAutomergeSnapshot,
  path: Array<string | number>,
): MindooDBAppRichTextSpan[] {
  const doc = A.load<Record<string, unknown>>(snapshot.binary);
  return readRichTextSpans(doc, path);
}

function readRichTextSpans(
  doc: A.Doc<Record<string, unknown>>,
  path: Array<string | number>,
): MindooDBAppRichTextSpan[] {
  return dehydrateRichTextSpans(
    A.spans(doc, path as A.Prop[]) as A.Span[],
  );
}

function automergeHeadsEqual(left: readonly string[], right: readonly string[]) {
  if (left.length !== right.length) {
    return false;
  }
  return [...left].sort().join("|") === [...right].sort().join("|");
}
