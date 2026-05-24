# MindooDB App TeamEdit

TeamEdit is a Vue 3 sample app for [MindooDB Haven](https://mindoodb.com) that demonstrates **real-time-style collaborative markdown editing** on top of the [`mindoodb-app-sdk`](https://github.com/klehmann/mindoodb-app-sdk).

It opens MindooDB documents from one or more shared databases, edits the document `body` field with a Milkdown WYSIWYG markdown editor, and lets multiple users (or multiple windows of the same user) edit the same document at once. When two windows save concurrent changes, MindooDB merges them through the SDK's text patch API and the editor reconciles to the canonical merged result.

TeamEdit is intentionally small and well-commented so other developers can use it as a learning resource for building MindooDB apps that need granular collaborative text editing.

## New to MindooDB apps? Start here

If this is your first MindooDB app, the following 60-second crash course should make the rest of this README a lot easier to follow:

- **MindooDB Haven** is the host application. It runs in the browser, owns the user's MindooDB databases, and embeds third-party "apps" (like TeamEdit) in a **sandboxed iframe**.
- A **MindooDB app** is just a regular static web app (HTML + JS + CSS). It does not talk to MindooDB or the database directly. Instead it talks to Haven through a JSON message bridge provided by the [`mindoodb-app-sdk`](https://github.com/klehmann/mindoodb-app-sdk). Haven is the only side that holds the real Automerge documents; the app only sees JSON snapshots.
- A **document** is the unit of data inside a MindooDB database. It is a JSON-shaped object stored as an Automerge document, so it can be edited collaboratively. TeamEdit reads and writes the `body` field as a markdown string.
- An **attachment** is a binary blob (image, PDF, anything) that "belongs" to a document. Haven streams attachment bytes over the bridge so apps never have to base64-encode files into JSON.
- **Capabilities** are the permissions Haven grants the app for each mapped database (`read`, `create`, `update`, `delete`, `attachments`, `history`). The UI in TeamEdit is gated on those capabilities, so the same app code Just Works whether the host gave it read-only or full access.
- The **SDK bridge** is asynchronous and JSON-only. Apps don't ship Automerge or any database runtime. They just call SDK methods like `database.documents.get(id)` or `database.documents.update(id, { set: {…} })`.

If you want a 5-minute tour of the SDK surface, browse [`mindoodb-app-sdk/README.md`](https://github.com/klehmann/mindoodb-app-sdk) — it covers everything below at a higher level.

## What TeamEdit demonstrates

- **Granular collaborative text editing over a JSON bridge** using `MindooDBTextBuffer` from the SDK.
- **Save-time conflict resolution** based on the document's Automerge heads, performed by Haven on behalf of the app.
- **A customized Milkdown / Crepe editor** with attachment-backed images, Mermaid diagram support, LaTeX authoring, and app-owned typography/theme styling.
- **Markdown preview and print rendering** using the same renderer pipeline as the editor, including attachment image URL resolution, hydrated Mermaid diagrams, LaTeX math, syntax-highlighted code blocks, task lists, footnotes, and safe alignment wrappers.
- **MindooDoc attachments** for non-text data: a bottom dropzone panel for upload, preview, download, and delete.
- **Attachment-backed markdown images** via Milkdown's image upload/proxy hooks. Images are stored as document attachments, the markdown stores a stable `mindoodb-attachment:` URL, and the app resolves it to a temporary `blob:` URL for the editor and preview pane.
- **Capability-aware UI**: File / New, File / Save, File / Delete, and the attachment panel are gated by the SDK capabilities granted to the app.
- **Read-only revision browsing** through the SDK history capability. TeamEdit shows a single status badge (`Current · Saved`, `Current · Unsaved`, or `Historical · <date>`) and opens selected revisions in the same Crepe editor with editing disabled.
- **Host-aware UI preferences**: the toolbar reserves space on the left for iPadOS multitasking controls when Haven asks for it, and adapts to the small-screen PrimeVue mobile menu layout.
- **Persistent UI preferences**: the preview pane visibility and right/bottom placement are stored in `localStorage`.

## TeamEdit features

- **Document lifecycle actions**: create, open, refresh, save, inspect document metadata, delete, and print markdown documents from the File menu.
- **Subject editing**: edit the document subject independently from the markdown body while saving both through MindooDB document updates.
- **Live preview pane**: toggle the rendered preview on or off and place it to the right of the editor or below it. Editor-side image resizing (via the Crepe drag handle) is mirrored in the preview and the print view.
- **Print view**: open a dedicated print window with sanitized rendered markdown, resolved attachment images, rendered Mermaid diagrams, KaTeX math, syntax-highlighted code blocks, task lists, footnotes, print-focused styles, and asset loading before `window.print()`.
- **Markdown export**: save the current document as a plain `.md` file or, when attachments exist, as a portable ZIP package containing rewritten markdown plus attachment files.
- **Document revision history**: click the top-right status badge to browse timestamp-sorted revisions backed by stable MindooDB DAG revision IDs. Historical revisions are read-only, but print and export still work.
- **Attachment panel**: upload files by dropzone, preview them in Haven, download them, delete them, and reuse uploaded files in the markdown body.
- **Insert existing attachments from the editor**: a slash-menu **Attachment** action opens a picker dialog so you can drop a previously uploaded image or file into the markdown without re-uploading it.
- **Theme integration**: follows Haven's light/dark theme mode and uses bundled editor fonts (`Inter`, `Rubik`, and `JetBrains Mono`) so deployed builds do not depend on third-party font CDNs.

## Architecture overview

```
Haven (browser tab)                       Database / Sync
  ┌────────────┐                            ┌─────────────┐
  │ Bridge host│ ──── SDK message bridge ───┤  MindooDB   │
  └─────┬──────┘                            └─────────────┘
        │ documents.update({ text: [...], baseHeads }) }
  ┌─────▼──────────────────────────────────────────────┐
  │ TeamEdit (sandboxed iframe / window)               │
  │                                                    │
  │  Milkdown / Crepe editor                           │
  │       │  full-markdown updates                     │
  │       ▼                                            │
  │  MindooDBTextBuffer  ── pending text edits ──┐     │
  │       ▲                                       │   │
  │       │ canonical text after flush            │   │
  │       │                                       ▼   │
  │  documents.update({ text: [...], baseHeads }) ────┘
  └────────────────────────────────────────────────────┘
```

The app never owns an Automerge replica. It only sees JSON snapshots of documents, the document's current heads, and a JSON-safe text patch shape. Haven applies the patch against the real Automerge document, merges any concurrent changes, and returns the new canonical document.

## How collaborative text editing works

The MindooDB App SDK ships a small helper, `createMindooDBTextBuffer`, designed for editors like Milkdown that emit full-markdown strings rather than raw text splices. It is the backbone of TeamEdit's collaborative editing.

A buffer is created per opened document and pinned to a JSON path inside the document, in TeamEdit's case `["body"]`:

```ts
// src/App.vue (excerpt)
import { createMindooDBTextBuffer } from "mindoodb-app-sdk";

textBuffer.value = createMindooDBTextBuffer({
  database,
  document,        // includes the document's current Automerge heads
  path: ["body"],  // the markdown text field to edit
});
markdown.value = textBuffer.value.value;
```

The buffer keeps a local string and tracks pending text operations. Whenever Milkdown reports a new full markdown value, TeamEdit translates that into a single splice into the buffer:

```ts
function onEditorUpdate(value: string) {
  markdown.value = value;
  textBuffer.value?.replaceText(value);   // diffs to one minimal splice
  isDirty.value = textBuffer.value?.dirty ?? false;
}
```

Saving sends the buffered splices to Haven together with the heads the buffer was created at. Haven applies the operations causally against the original version and merges them with any concurrent changes from other apps or syncing peers:

```ts
const result = await textBuffer.value.flush();
// result.document is the new canonical document
// result.value    is the canonical text after merge
// result.reconciled is true when Haven returned text that differs from
//                   what we just sent, i.e. concurrent edits were merged.
if (result.reconciled) {
  markdown.value = result.value; // replace editor content with merged result
}
```

The SDK call under the hood is a regular `documents.update` with a `text` patch:

```ts
await db.documents.update(docId, {
  text: [{
    path: ["body"],
    baseHeads: previousHeads,           // version this client is editing against
    edits: [{ index, deleteCount, insert }, /* ... */],
  }],
});
```

This keeps the bridge purely JSON, so apps never have to ship an Automerge runtime themselves, while still benefiting from Automerge's CRDT-based text merging on the Haven side.

### What this means in practice

- Open the same document in two TeamEdit windows.
- Edit different parts of the markdown in each window.
- Save in window 1, then save in window 2.
- Window 2's save is applied at its original heads and merged with the changes window 1 already saved. The editor in window 2 then displays the merged document.

If you want to inspect the merge result without saving, use the **Refresh** action in the toolbar to re-read the canonical document from Haven.

## Word document editing

TeamEdit also supports **Word documents** (`type: "word"`) through the `@eigenpal/docx-editor-vue` editor. Word docs share the same top-level field name as markdown — `["body"]` — but use Automerge **rich text** rather than plain text. The app never calls `set: { body: ... }` on Word saves; all body edits go through the SDK's `richText` patch API.

### Load / edit / save flow

1. **Load:** `documents.getRichText(docId, ["body"])` returns JSON-safe rich-text spans. TeamEdit converts them to a docx-editor `Document` via `richTextSpansToDocument`.
2. **Edit:** `createMindooDBRichTextHandle` tracks local span snapshots. Each editor change calls `handle.replaceSpans(documentToRichTextSpans(document))`.
3. **Save:** `handle.flush()` sends incremental rich-text splices when text changed (merge-friendly, like the markdown text buffer), or full span snapshots for formatting-only edits. After every save, the editor reloads from the canonical spans Haven returns.
4. **Manual refresh:** Use **Refresh** in the toolbar to re-read the document from Haven when another client may have saved changes.

Comments remain a separate top-level JSON field (`comments`) and are saved with ordinary `set` patches, independent of the rich-text body.

### Span conversion

Conversion between docx-editor ProseMirror documents and MindooDB spans lives under `src/editors/word/lib/`:

- `automergeSchemaAdapter.ts` — maps the docx-editor schema to `@automerge/prosemirror`'s `SchemaAdapter`
- `automergeSpanUtils.ts` — table block markers, span normalization, and JSON dehydration/revival
- `richTextSpans.ts` — public `documentToRichTextSpans` / `richTextSpansToDocument` helpers

`@automerge/automerge` is used only for span conversion in the browser bundle; the app still does not own an Automerge replica. Haven applies rich-text patches on the server side.

### Creating and importing Word documents

New Word files, DOCX imports, and documents created from Word templates seed rich text explicitly after document creation:

```ts
await db.documents.create({ set: { type: "word", form: "teamedit", comments: [], /* ... */ } });
await db.documents.update(docId, {
  richText: [{ path: ["body"], spans: documentToRichTextSpans(docxDocument) }],
});
```

Historical Word revisions load read-only content via `documents.getRichText(docId, ["body"], { revisionId })`.

## Milkdown editor extensions

TeamEdit uses Milkdown's Crepe editor as the editing surface, but extends it in a few app-specific places:

- **Image upload / proxy hooks**: Crepe's image block feature is connected to MindooDB attachments. Uploads stream into the current document and the editor receives a stable `mindoodb-attachment:` markdown URL; display uses Crepe's `proxyDomURL` hook to resolve that stable URL to a temporary object URL.
- **Attachment slash-menu entry**: a TeamEdit-specific **Attachment** action in the block edit menu opens a picker dialog with the document's existing attachments. Picking an image inserts it as a markdown image; picking any other file inserts it as a markdown link. The picker shares its image URL resolver with the editor and preview, so thumbnails are shown without a second download.
- **Diagram slash-menu entry**: the block edit menu gets a TeamEdit-specific **Diagram** action that inserts a starter Mermaid flowchart fence.
- **Mermaid code fence previews**: fenced code blocks with the `mermaid` language render as diagrams. Inside the editor, TeamEdit customizes Crepe's CodeMirror preview renderer; in the preview pane and print window, the markdown renderer emits hydratable placeholders that are rendered with Mermaid after the HTML is mounted.
- **Safe diagram rendering**: Mermaid output is sanitized before insertion. For editor previews, foreign-object labels are converted to SVG text so diagrams render reliably inside Crepe's preview surface.
- **Image ratio handling in preview / print**: Crepe's image block stores the user's drag-handle resize ratio inside the markdown image's `alt` field (e.g. `![0.50](src "caption")`). TeamEdit's renderer translates that into a `data-teamedit-image-ratio` attribute, promotes the markdown title to a real `alt` for accessibility, and a small post-load helper applies the same height calculation Crepe uses, so resized images look identical in the editor, the preview pane, and the print view.
- **App-owned editor theme**: TeamEdit imports Crepe's common structural CSS, then supplies its own Crepe color variables and typography. Headings use `Rubik`, body text uses `Inter`, and code uses `JetBrains Mono`.

## Markdown preview, print, and portability

TeamEdit stores the document body as plain Markdown, but the preview pane and print window use a shared renderer so common technical-documentation syntax looks like it does in the editor:

- **Syntax highlighting**: fenced code blocks are highlighted with `highlight.js`, including common languages such as `markdown`, `ts`, `js`, `json`, `html`, and `css`.
- **LaTeX math**: inline `$E = mc^2$` and display `$$ ... $$` math render through KaTeX.
- **Mermaid diagrams**: ```` ```mermaid ```` fences render as sanitized SVG while preserving the original portable Markdown source.
- **Task lists**: `[ ]`, `[x]`, and `[X]` list items render as disabled checkboxes in preview and print.
- **Footnotes**: `[^1]` references and `[^1]: ...` definitions render as linked footnotes with backlinks.
- **Safe alignment wrappers**: legacy snippets using `<div align="left">`, `<div align="center">`, or `<div align="right">` are converted to internal alignment classes. Arbitrary raw HTML is still disabled and escaped.

For example:

````markdown
```ts
const message = "Highlighted in preview and print";
```

Inline math: $E = mc^2$

$$
\int_{-\infty}^{\infty} e^{-x^2} dx = \sqrt{\pi}
$$

* [ ] To-do item
* [x] Completed item

This has a footnote.[^1]

[^1]: Footnotes render at the bottom of the document.
````

## Attachments and images

TeamEdit treats binary content as MindooDoc attachments rather than embedding base64 in the markdown. Three different surfaces collaborate to make this feel seamless:

- The bottom **Attachments panel** is a dropzone with upload, preview, download, and delete actions. Preview uses Haven's built-in viewer through the SDK (`attachments.openPreview` in iframe mode, `preparePreviewSession` + `window.open` in window mode).
- The Milkdown **image upload** widget is wired to the SDK's `attachments.openWriteStream`. After the upload finishes, TeamEdit inserts a stable `mindoodb-attachment:<encoded-name>` URL into the markdown so the document body never depends on per-session blob URLs.
- The Milkdown slash-menu **Attachment** action lets users pick from the document's *existing* attachments without re-uploading. The picker (a Vue dialog living in the host) reads the document's attachment list, shows image thumbnails through the same resolver the editor uses, and inserts either `![caption](mindoodb-attachment:NAME "caption")` for image attachments or `[label](mindoodb-attachment:NAME)` for any other file.

For display, a small image resolver loads each attachment as a `Blob`, creates a temporary `blob:` URL, and proxies it into the editor via Crepe's `proxyDomURL` hook. The same cached object URLs are reused by the markdown preview pane and the print window, so an attachment is downloaded at most once per session. Object URLs are revoked when the document is closed, refreshed, or deleted, keeping memory and IndexedDB blob references clean.

When viewing a historical revision, the attachment panel switches to the revision's historical attachment list. Preview, download, print, and ZIP export pass the selected `revisionId` through the SDK so Haven resolves the attachment from that revision's `_attachments` array. This matters for attachments that grew over time: MindooDB appends chunks and stores the new `lastChunkId` in newer document revisions, so historical reads must start from the old `lastChunkId` to avoid reading bytes that did not exist yet.

### Why a custom URL scheme?

Browser `blob:` URLs are session-local, so embedding them in the document body would break as soon as anyone reloads the page. Instead TeamEdit stores `mindoodb-attachment:<encoded-name>` in the markdown — a stable identifier that any peer can resolve back to a fresh `blob:` URL when they open the document. This keeps the markdown portable: when you export the document with attachments, the export pipeline rewrites those URLs into relative paths inside the ZIP package automatically.

## Markdown export

TeamEdit has two export paths in the File menu:

- **Export Markdown** writes the current editor body to a `.md` file. This preserves TeamEdit's stable `mindoodb-attachment:` URLs exactly as stored in the document, which is useful for backups or debugging the raw markdown.
- **Export Markdown with attachments** is shown only when the current document has attachments. It writes a ZIP package with `document.md` and an `attachments/` folder. The exported `document.md` rewrites TeamEdit attachment URLs to relative package paths such as `attachments/teamedit-images/photo.png`, so the markdown can be opened outside Haven together with its files.

The export helper first tries the browser File System Access API (`showSaveFilePicker`) to display a native save dialog. Browsers without that API fall back to a regular single-file download. Attachment packages are generated client-side with `fflate`; no attachment bytes leave the browser except through the user's explicit export download.

When you export from a historical revision, TeamEdit packages the historical markdown body together with the attachment versions referenced by that same revision.

## Project structure

```text
src/
  App.vue                                 Top-level app shell, menus, dialogs, save/refresh
  main.ts                                 Vue + PrimeVue + Crepe styles bootstrap
  components/
    AttachmentPickerDialog.vue            Slash-menu picker for inserting an existing attachment
    DocumentRevisionDialog.vue            Revision picker opened from the status badge
    DocumentAttachmentsPanel.vue          Bottom dropzone panel with upload/preview/download
    MilkdownMarkdownEditor.vue            Crepe wrapper with image hooks, Mermaid previews, slash-menu items
  composables/
    useAttachmentImageResolver.ts         Caches attachment blob URLs, preloads markdown images
    useDocumentAttachments.ts             Wraps SDK attachment API for upload, preview, download, delete
  lib/
    attachmentImages.ts                   Pure helpers: URL scheme, file naming, stream IO, formatting
    attachmentImages.test.ts              Focused tests for the helpers
    exportMarkdown.ts                     Browser save helpers, ZIP packaging, attachment URL rewriting
    exportMarkdown.test.ts                Focused tests for markdown and package export helpers
    imageRatio.ts                         Mirrors Crepe's drag-handle ratio in preview / print views
    markdownRendering.ts                  Markdown-it renderer, math, code highlighting, task lists, footnotes, Mermaid placeholders
    markdownRendering.test.ts             Renderer tests including math, highlighting, task lists, footnotes, alignment, ratio / caption translation
    mermaid.ts                            Lazy Mermaid renderer and sanitized SVG helpers
    printMarkdown.ts                      Print-window rendering with images, ratios and Mermaid hydration
    theme.ts                              Haven theme bootstrap
  assets/styles/                          App theme styling
  env.d.ts                                Side-effect CSS module declarations
```

The collaborative text editing logic lives in `App.vue` together with the document lifecycle (open, refresh, save, print, export, delete). The attachment work is split into composables, a focused panel component, and a separate slash-menu picker dialog so each integration point is easy to study independently. Rendering and export helpers live under `src/lib/` so the editor preview pane, print window, and downloaded packages share the same markdown, image, and Mermaid behavior.

### Suggested reading order

If you are using TeamEdit as a learning resource, a useful order to read the source in is:

1. `src/main.ts` — bootstraps Vue, PrimeVue, the Crepe styles, and the Haven theme. Tiny, but the entry point.
2. `src/App.vue` — the heart of the integration. Read the top-of-file imports, the `onMounted` Haven bridge bootstrap, and the `loadDocumentIntoEditor` / `saveFile` / `loadHistoricalRevision` / `refreshCurrentDocument` functions; that's the SDK lifecycle in one screen.
3. `src/components/MilkdownMarkdownEditor.vue` — the editor wrapper. The `featureConfigs` block shows how to plug Crepe's image upload and slash menu into your own host code.
4. `src/composables/useAttachmentImageResolver.ts` and `src/lib/attachmentImages.ts` — the attachment URL scheme and the small cache that powers images in the editor, preview, and print view.
5. `src/lib/markdownRendering.ts` and `src/lib/printMarkdown.ts` — how the same renderer pipeline serves both the live preview pane and the print window.
6. `src/lib/exportMarkdown.ts` — how the markdown body and its attachments are bundled into a portable ZIP package.

## Development

### First-time setup

```bash
pnpm install
pnpm dev
```

`pnpm dev` starts a Vite dev server on **http://127.0.0.1:4206** using the published `mindoodb-app-sdk` from npm. Hot reload is enabled, so saving a `.vue` or `.ts` file refreshes the running editor inside Haven.

`pnpm dev:local` keeps an optional local-monorepo workflow available by aliasing `mindoodb-app-sdk` to a sibling checkout. Use this when you are working on the SDK itself; otherwise stick to `pnpm dev`.

### Connect TeamEdit to your local Haven instance

TeamEdit cannot run on its own — it needs Haven to host it. Once `pnpm dev` is running:

1. In Haven, open **Application settings** and register a new application pointing to `http://127.0.0.1:4206` (or the deployed URL of your TeamEdit instance).
2. Map one or more databases to the app and grant capabilities. For the full TeamEdit experience you want `read`, `create`, `update`, `delete`, `attachments`, and `history`. Removing any of those capabilities is a useful way to see how the UI gracefully degrades.
3. Launch TeamEdit from Haven. Use **File / New** to create an empty document or **File / Open** to pick an existing one.

The Haven launcher embeds TeamEdit in a sandboxed iframe and feeds it a launch context (theme, capabilities, mapped databases). You can also launch TeamEdit on a separate browser window — TeamEdit's code adapts to either case automatically (`session.runtime` is `"iframe"` or `"window"`).

### Tips for hacking on the code

- The `dev` server uses Vite's HMR; most edits don't require a full reload. The Crepe editor instance is rebuilt on prop changes, so changes to `MilkdownMarkdownEditor.vue` always take effect immediately.
- If something looks wrong after a hot reload, just refresh TeamEdit inside Haven — the SDK bridge re-establishes a fresh session.
- The `dist/` folder is what gets uploaded by `pnpm deploy`. Don't edit it; it's regenerated on each build.

## Commands

| Command | Description |
|---|---|
| `pnpm dev` | Start Vite dev server using published packages |
| `pnpm dev:local` | Start Vite dev server using sibling MindooDB packages |
| `pnpm build` | Type-check and build a production bundle into `dist/` |
| `pnpm build:local` | Build with local Vite aliasing for sibling MindooDB packages |
| `pnpm test` | Run the Vitest suite |
| `pnpm test:watch` | Run tests in watch mode |
| `pnpm preview` | Build and preview the Cloudflare deployment locally with Wrangler |
| `pnpm deploy` | Build and deploy the app to Cloudflare |

## Deployment to Cloudflare

TeamEdit is configured for Cloudflare Workers static asset hosting, the same pattern used by `mindoodb-app-example`, `mindoodb-app-vega`, and `mindoodb-app-todomanager`.

The deployment config lives in `wrangler.jsonc`:

```jsonc
{
  "name": "mindoodb-app-teamedit",
  "compatibility_date": "2026-04-28",
  "assets": {
    "directory": "./dist",
    "not_found_handling": "single-page-application"
  }
}
```

Deploy with:

```bash
pnpm deploy
```

This runs `vue-tsc --noEmit && vite build` and then `wrangler deploy`, uploading the `dist/` folder as static assets with SPA fallback enabled. Any other static host works too, but the Cloudflare Workers + Wrangler path matches the rest of the MindooDB sample apps.

## Testing

The Vitest suite covers the non-trivial pure helpers used by the attachment, rendering, Mermaid, and print flows:

- `src/lib/attachmentImages.test.ts` -- attachment URL scheme round-trip, file-name sanitization, markdown image URL extraction, formatted size, chunked upload over the SDK write stream, and read stream into a blob.
- `src/lib/exportMarkdown.test.ts` -- export filename cleanup, ZIP attachment paths, markdown URL rewriting, and packaged attachment bytes.
- `src/components/DocumentRevisionDialog.test.ts` -- revision selection, double-click open, and empty-state behavior.
- `src/lib/markdownRendering.test.ts` -- Mermaid placeholder generation, syntax-highlighted code fences, KaTeX math, task lists, footnotes, safe alignment wrappers, image URL rewriting, Milkdown `<br>` normalization, and the Crepe ratio/caption translation used by the preview and print views.
- `src/lib/mermaid.test.ts` -- Mermaid SVG sanitization and label conversion for Crepe previews.
- `src/lib/printMarkdown.test.ts` -- print-window rendering, asset waiting, attachment image resolution, and Mermaid hydration.
- `src/editors/word/lib/richTextSpans.test.ts` -- Word rich-text span round-trips, bold mark preservation, and table structure.

Run them with:

```bash
pnpm test
```

The collaborative text editing flow is covered end-to-end by the SDK and Haven test suites; TeamEdit relies on those by composing the `MindooDBTextBuffer` API documented above.

## Further reading

- [`mindoodb-app-sdk`](https://github.com/klehmann/mindoodb-app-sdk) -- the SDK reference, including the launch context, capability model, document and attachment APIs, and host events.
- [`mindoodb-app-example`](https://github.com/klehmann/mindoodb-app-example) -- a broader reference app exercising every SDK area.
- [`mindoodb-app-vega`](https://github.com/klehmann/mindoodb-app-vega) -- a richer Vue 3 sample using attachments and Haven views.

## License

TeamEdit is released under the [Apache License 2.0](./LICENSE). See the [`LICENSE`](./LICENSE) file for the full text.

Copyright (c) 2026 Mindoo GmbH.
