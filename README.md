# MindooDB App TeamEdit

TeamEdit is a Vue 3 sample app for [MindooDB Haven](https://mindoodb.com) that demonstrates **real-time-style collaborative markdown editing** on top of the [`mindoodb-app-sdk`](https://github.com/klehmann/mindoodb-app-sdk).

It opens MindooDB documents from one or more shared databases, edits the document `body` field with a Milkdown WYSIWYG markdown editor, and lets multiple users (or multiple windows of the same user) edit the same document at once. When two windows save concurrent changes, MindooDB merges them through the SDK's text patch API and the editor reconciles to the canonical merged result.

TeamEdit is intentionally small and well-commented so other developers can use it as a learning resource for building MindooDB apps that need granular collaborative text editing.

## What TeamEdit demonstrates

- **Granular collaborative text editing over a JSON bridge** using `MindooDBTextBuffer` from the SDK.
- **Save-time conflict resolution** based on the document's Automerge heads, performed by Haven on behalf of the app.
- **A customized Milkdown / Crepe editor** with attachment-backed images, Mermaid diagram support, and app-owned typography/theme styling.
- **Markdown preview and print rendering** using the same renderer pipeline as the editor, including attachment image URL resolution and hydrated Mermaid diagrams.
- **MindooDoc attachments** for non-text data: a bottom dropzone panel for upload, preview, download, and delete.
- **Attachment-backed markdown images** via Milkdown's image upload/proxy hooks. Images are stored as document attachments, the markdown stores a stable `mindoodb-attachment:` URL, and the app resolves it to a temporary `blob:` URL for the editor and preview pane.
- **Capability-aware UI**: File / New, File / Save, File / Delete, and the attachment panel are gated by the SDK capabilities granted to the app.
- **Host-aware UI preferences**: the toolbar reserves space on the left for iPadOS multitasking controls when Haven asks for it, and adapts to the small-screen PrimeVue mobile menu layout.
- **Persistent UI preferences**: the preview pane visibility and right/bottom placement are stored in `localStorage`.

## TeamEdit features

- **Document lifecycle actions**: create, open, refresh, save, inspect document metadata, delete, and print markdown documents from the File menu.
- **Subject editing**: edit the document subject independently from the markdown body while saving both through MindooDB document updates.
- **Live preview pane**: toggle the rendered preview on or off and place it to the right of the editor or below it.
- **Print view**: open a dedicated print window with sanitized rendered markdown, resolved attachment images, rendered Mermaid diagrams, print-focused styles, and asset loading before `window.print()`.
- **Markdown export**: save the current document as a plain `.md` file or, when attachments exist, as a portable ZIP package containing rewritten markdown plus attachment files.
- **Attachment panel**: upload files by dropzone, preview them in Haven, download them, delete them, and reuse uploaded files as markdown images.
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

## Milkdown editor extensions

TeamEdit uses Milkdown's Crepe editor as the editing surface, but extends it in a few app-specific places:

- **Image upload / proxy hooks**: Crepe's image block feature is connected to MindooDB attachments. Uploads stream into the current document and the editor receives a stable `mindoodb-attachment:` markdown URL; display uses Crepe's `proxyDomURL` hook to resolve that stable URL to a temporary object URL.
- **Mermaid code fence previews**: fenced code blocks with the `mermaid` language render as diagrams. Inside the editor, TeamEdit customizes Crepe's CodeMirror preview renderer; in the preview pane and print window, the markdown renderer emits hydratable placeholders that are rendered with Mermaid after the HTML is mounted.
- **Diagram slash-menu entry**: the block edit menu gets a TeamEdit-specific **Diagram** action that inserts a starter Mermaid flowchart fence.
- **Safe diagram rendering**: Mermaid output is sanitized before insertion. For editor previews, foreign-object labels are converted to SVG text so diagrams render reliably inside Crepe's preview surface.
- **App-owned editor theme**: TeamEdit imports Crepe's common structural CSS, then supplies its own Crepe color variables and typography. Headings use `Rubik`, body text uses `Inter`, and code uses `JetBrains Mono`.

## Attachments and images

TeamEdit treats binary content as MindooDoc attachments rather than embedding base64 in the markdown.

- The bottom **Attachments panel** is a dropzone with upload, preview, download, and delete actions. Preview uses Haven's built-in viewer through the SDK (`attachments.openPreview` in iframe mode, `preparePreviewSession` + `window.open` in window mode).
- The Milkdown image picker is wired to the SDK's `attachments.openWriteStream`. After the upload finishes, TeamEdit inserts a stable `mindoodb-attachment:<encoded-name>` URL into the markdown.
- For display, a small image resolver loads the attachment as a `Blob`, creates a temporary `blob:` URL, and proxies it into the editor via Crepe's `proxyDomURL` hook. The same cached object URLs are also used by the markdown preview pane.
- Object URLs are revoked when the document is closed, refreshed, or deleted, so memory and IndexedDB blob references stay clean.

## Markdown export

TeamEdit has two export paths in the File menu:

- **Export Markdown** writes the current editor body to a `.md` file. This preserves TeamEdit's stable `mindoodb-attachment:` URLs exactly as stored in the document, which is useful for backups or debugging the raw markdown.
- **Export Markdown with attachments** is shown only when the current document has attachments. It writes a ZIP package with `document.md` and an `attachments/` folder. The exported `document.md` rewrites TeamEdit attachment URLs to relative package paths such as `attachments/teamedit-images/photo.png`, so the markdown can be opened outside Haven together with its files.

The export helper first tries the browser File System Access API (`showSaveFilePicker`) to display a native save dialog. Browsers without that API fall back to a regular single-file download. Attachment packages are generated client-side with `fflate`; no attachment bytes leave the browser except through the user's explicit export download.

## Project structure

```text
src/
  App.vue                                 Top-level app shell, menus, dialogs, save/refresh
  main.ts                                 Vue + PrimeVue + Crepe styles bootstrap
  components/
    DocumentAttachmentsPanel.vue          Bottom dropzone panel with upload/preview/download
    MilkdownMarkdownEditor.vue            Crepe wrapper with image hooks and Mermaid editor previews
  composables/
    useAttachmentImageResolver.ts         Caches attachment blob URLs, preloads markdown images
    useDocumentAttachments.ts             Wraps SDK attachment API for upload, preview, download, delete
  lib/
    attachmentImages.ts                   Pure helpers: URL scheme, file naming, stream IO, formatting
    attachmentImages.test.ts              Focused tests for the helpers
    exportMarkdown.ts                     Browser save helpers, ZIP packaging, attachment URL rewriting
    exportMarkdown.test.ts                Focused tests for markdown and package export helpers
    markdownRendering.ts                  Markdown-it renderer, attachment image rewriting, Mermaid placeholders
    mermaid.ts                            Lazy Mermaid renderer and sanitized SVG helpers
    printMarkdown.ts                      Print-window rendering with images and Mermaid hydration
    theme.ts                              Haven theme bootstrap
  assets/styles/                          App theme styling
  env.d.ts                                Side-effect CSS module declarations
```

The collaborative text editing logic lives in `App.vue` together with the document lifecycle (open, refresh, save, print, export, delete). The attachment work is split into composables and a focused panel component so each integration point is easy to study independently. Rendering and export helpers live under `src/lib/` so the editor preview pane, print window, and downloaded packages share the same markdown, image, and Mermaid behavior.

## Development

```bash
npm install
npm run dev
```

`npm run dev` starts a Vite dev server on **http://127.0.0.1:4206** using the published `mindoodb-app-sdk` from npm.

`npm run dev:local` keeps an optional local-monorepo workflow available by aliasing `mindoodb-app-sdk` to a sibling checkout, useful when working on the SDK itself.

Then in Haven:

1. Open **Application settings** and register a new application pointing to `http://127.0.0.1:4206` (or the deployed URL of your TeamEdit instance).
2. Map one or more databases to the app and grant capabilities. For TeamEdit you typically want `read`, `create`, `update`, `delete`, and `attachments`.
3. Launch TeamEdit. Use **File / New** to create an empty document or **File / Open** to pick an existing one.

## Commands

| Command | Description |
|---|---|
| `npm run dev` | Start Vite dev server using published packages |
| `npm run dev:local` | Start Vite dev server using sibling MindooDB packages |
| `npm run build` | Type-check and build a production bundle into `dist/` |
| `npm run build:local` | Build with local Vite aliasing for sibling MindooDB packages |
| `npm test` | Run the Vitest suite |
| `npm run test:watch` | Run tests in watch mode |
| `npm run preview` | Build and preview the Cloudflare deployment locally with Wrangler |
| `npm run deploy` | Build and deploy the app to Cloudflare |

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
npm run deploy
```

This runs `vue-tsc --noEmit && vite build` and then `wrangler deploy`, uploading the `dist/` folder as static assets with SPA fallback enabled. Any other static host works too, but the Cloudflare Workers + Wrangler path matches the rest of the MindooDB sample apps.

## Testing

The Vitest suite covers the non-trivial pure helpers used by the attachment, rendering, Mermaid, and print flows:

- `src/lib/attachmentImages.test.ts` -- attachment URL scheme round-trip, file-name sanitization, markdown image URL extraction, formatted size, chunked upload over the SDK write stream, and read stream into a blob.
- `src/lib/exportMarkdown.test.ts` -- export filename cleanup, ZIP attachment paths, markdown URL rewriting, and packaged attachment bytes.
- `src/lib/markdownRendering.test.ts` -- Mermaid placeholder generation, non-Mermaid code fence preservation, image URL rewriting, and Milkdown `<br>` normalization.
- `src/lib/mermaid.test.ts` -- Mermaid SVG sanitization and label conversion for Crepe previews.
- `src/lib/printMarkdown.test.ts` -- print-window rendering, asset waiting, attachment image resolution, and Mermaid hydration.

Run them with:

```bash
npm test
```

The collaborative text editing flow is covered end-to-end by the SDK and Haven test suites; TeamEdit relies on those by composing the `MindooDBTextBuffer` API documented above.

## Further reading

- [`mindoodb-app-sdk`](https://github.com/klehmann/mindoodb-app-sdk) -- the SDK reference, including the launch context, capability model, document and attachment APIs, and host events.
- [`mindoodb-app-example`](https://github.com/klehmann/mindoodb-app-example) -- a broader reference app exercising every SDK area.
- [`mindoodb-app-vega`](https://github.com/klehmann/mindoodb-app-vega) -- a richer Vue 3 sample using attachments and Haven views.
