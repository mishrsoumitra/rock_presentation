# rock_presentation

A PowerPoint (.pptx) editor that runs in the browser, built to the *PPTX Editor Build Plan* (25 Oct 2026). All six
phases of the plan are implemented:
1. A lossless package and XML layer.
2. A typed model with PowerPoint's inheritance resolved.
3. A renderer for shapes, text, pictures, tables, charts, SmartArt and EMF/WMF.
4. A React editor with undo.
5. Animations, transitions and a slide show.
6. A server that renders slides to PNG and PDF and opens password-protected decks.

Everything is tested against about 900 third-party decks. What could not be tested is fidelity against PowerPoint
itself: the plan's oracles need PowerPoint on Windows and the .NET SDK (see [Not run](#not-run-the-powerpoint-oracles)).

## Running it

```sh
corepack pnpm install               # pnpm 12 (pinned in package.json); no dependency runs a build script
corepack pnpm corpus:fetch          # ~1,050 third-party test decks into corpus/external/ (gitignored)
corepack pnpm check                 # four type-checks, then every test: 10,233 tests in about 90 s
corepack pnpm dev                   # the editor, at http://localhost:5173
```

The server draws slides with the same packages in headless Chromium (Playwright's, installed by
`corepack pnpm exec playwright install chromium`):

```sh
corepack pnpm --filter @pptx/server build:render-page    # the page Chromium renders with
DATA_DIR=/tmp/pptx PORT=3000 corepack pnpm --filter @pptx/server start
corepack pnpm --filter @pptx/server corpus               # phase 6 exit run: upload, PNG and PDF of every corpus deck
```

Tools for looking at decks:

```sh
corepack pnpm inspect deck.pptx --slide 3          # resolved shapes, fills, fonts and text (milestone M2)
corepack pnpm render --width 960 --out out/r deck.pptx   # PNG of every slide, drawn in Chromium
node tools/render/contact.ts out/r out/sheets/r          # contact sheets of those PNGs
node tools/render/filmstrip.ts --slides 2 deck.pptx      # an animated slide, frame by frame
corepack pnpm resave in.pptx out.pptx              # milestone M1: re-save a deck, verified before anything is written
corepack pnpm corpus:roundtrip                     # the phase 1 numbers, in out/corpus-roundtrip.json
UPDATE_GOLDEN=1 node_modules/.bin/vitest run tools/render/golden.test.ts   # approve new golden renders
```

Node 24 runs the TypeScript directly (type stripping), so only the web app and the render page have a build step.
The code uses only erasable TypeScript syntax (`erasableSyntaxOnly`), with explicit `.ts` import paths.

## Layout

| Path | What it is |
| --- | --- |
| `packages/xml` | Lossless XML tree, strict tokenizer, serializer that writes the way PowerPoint does, canonical form for tests |
| `packages/opc` | Open Packaging Conventions: ZIP reading with limits checked before inflation, part names, content types, relationships, saving |
| `packages/model` | Typed views over the XML (shapes, text, fills, timing) that write back in schema order; `mc:AlternateContent`; units |
| `packages/resolve` | Inheritance: placeholders through layout and master, theme colours and fonts, styles, backgrounds, tables, text properties |
| `packages/geometry` | Every preset shape and text warp (from the standard's definitions), the guide formula evaluator, custom geometry, arcs |
| `packages/text` | Fonts (HarfBuzz shaping), line breaking (UAX #14), bidi, paragraph layout, bullets, tabs, columns, vertical text, autofit |
| `packages/render` | The render tree and its Canvas 2D backend: fills, lines, effects (shadows, glow, soft edges, reflection), pictures, tables, SmartArt |
| `packages/charts` | DrawingML charts as render primitives |
| `packages/metafile` | EMF, EMF+ and WMF playback on Canvas 2D, and the TIFF decoder |
| `packages/anim` | Timing trees, playback (frame sampling), slide transitions, the effect presets, and editing animations |
| `packages/editor-core` | Editing commands, undo/redo with coalescing, copy and paste, templates, saving |
| `packages/crypto` | MS-OFFCRYPTO: agile and standard decryption, agile encryption |
| `apps/web` | The editor: Vite, React 19.3, react-konva, Zustand; workers for thumbnails, saving, decryption and pictures |
| `apps/server` | Fastify 5: uploads checked in worker threads with heap caps and deadlines, storage, slide PNGs and PDFs, oracle endpoints |
| `tools/corpus` | Corpus download, file lists, the definition of a lossless save, the round-trip report, the PowerPoint COM oracle script |
| `tools/render` | The browser harnesses, contact sheets, filmstrips, the golden ratchet (`golden/`) and the end-to-end editor tests |
| `tools/inspect`, `tools/resave` | The M2 deck inspector and the M1 re-save tool |
| `tools/geometry`, `tools/schema`, `tools/anim` | Generators: preset geometry, schema child order, effect presets |
| `tools/policy` | The dependency allowlist, enforced as a test |
| `tools/validator`, `tools/crypto` | The Open XML SDK validator (oracle O1), and a differential test against msoffcrypto-tool |

`packages/*` import nothing from Node, React or the DOM. They are type-checked against the `es2024` and `webworker`
libraries only (`tsconfig.src.json`), so they run unchanged in browsers, workers and Node.

## Status against the exit criteria

Measured on the downloaded corpus: 903 OPC packages from Apache POI, python-pptx, the Open XML SDK and LibreOffice
(888 open; 886 are presentations), plus 150 legacy `.ppt` files.

**Phase 1: lossless package and XML layer.**

| Criterion | Status |
| --- | --- |
| Every corpus file round-trips losslessly | **Done.** All 888 packages that open come back with every entry present, XML parts equal in canonical form and every other part byte for byte. The same holds with all 32,920 XML parts re-serialized from their trees. Saving is deterministic and idempotent. |
| Parse then serialize is byte-identical for PowerPoint-written parts | **Done** for about 28,900 machine-written parts. Parts that test authors pretty-printed or edited are excluded. |
| Unopenable files are refused cleanly; fuzz and limit tests | **Done.** The 15 refusals match `corpus/expected-refusals.json`. All 150 `.ppt` files are refused with a reason. ZIP bombs, deep nesting and huge attributes fail within seconds. |
| Saved files open in PowerPoint (O2); validator count unchanged (O1) | **Not run** (see below). |

**Phase 2: typed model and inheritance.**

| Criterion | Status |
| --- | --- |
| Resolved properties match PowerPoint's (`Shape.Left`, `Fill.ForeColor.RGB`...) for every corpus shape | **Not run:** needs PowerPoint's COM interface. Instead, every slide, layout and master of the corpus resolves without error, and unit tests pin the inheritance rules. `corepack pnpm inspect` shows what was resolved. |
| Editing transforms on every corpus file and saving still passes the phase 1 oracle | **Done**, with this repository's oracle (save, reopen, compare canonically, check references). Every shape of each deck's first slide is moved, and the editor's commands move, resize, rotate and flip shapes on every deck. |

**Phase 3: rendering.**

| Criterion | Status |
| --- | --- |
| Line breaks and pixels match PowerPoint (1% shapes, 2% text, 5% charts) | **Not run:** needs PowerPoint's own export (O3, O4). Instead, every corpus slide builds and draws, every corpus chart and metafile draws, and a ratchet compares 14 representative slides with approved renders under the same 1/2/5% limits (`tools/render/golden.test.ts`). This catches regressions, not fidelity. |

**Phase 4: the editor.**

| Criterion | Status |
| --- | --- |
| Every editing command, on every corpus deck, saves and passes the phase 1 oracle | **Done.** Up to 45 commands on each of the 886 decks, each where it applies (text commands need a text box; cropping needs a picture). They cover: inserting shapes, text boxes and pictures (PNG and SVG); moving, resizing, rotating, flipping, aligning, distributing, grouping and z-order; deleting, duplicating, copying and pasting; fill, outline, name and alt text; cropping and changing pictures; typing, deleting and formatting text; slides (new, duplicate, move, delete, hide, layout, background, size); importing masters; animations and transitions. |
| Undo restores a tree equal to the one before | **Done.** Checked after every command (undo, redo, undo). |
| A 300-slide, 150 MB deck opens in under 3 s; dragging runs at 60 fps | **Done** in the end-to-end tests. A 300-slide, 140 MB deck shows its first slide in about 0.2 s, and a frame of a drag redraws in about 1 ms. |

**Phase 5: animations and transitions.**

| Criterion | Status |
| --- | --- |
| The generated effect library round-trips through PowerPoint; video comparison within 10% | **Not run:** both need PowerPoint. |
| Playback runs the corpus without errors | **Done.** Every timing tree in the corpus parses, schedules and samples. The slide show plays them with triggers, "after previous" and timed advance. |

**Phase 6: the server.**

| Criterion | Status |
| --- | --- |
| Upload, render and PDF pass for the whole corpus, parser heap capped at 1 GB | **Done** (`corepack pnpm --filter @pptx/server corpus`). 903 packages uploaded in 55 s: 886 decks gave 1,957 slide PNGs and PDFs with a page for each; 16 were refused with their reviewed codes; the encrypted one was stored as uploaded. |
| A 10 GB-declared ZIP bomb, 200,000 entries and a malformed CFB fail in time without affecting other requests | **Done** (`apps/server/test/hostile.test.ts`), along with out-of-heap, deadline and crashed-worker cases. |
| Encrypted decks from PowerPoint 2007, 2013 and Microsoft 365 open with the right password and not the wrong one | **Partly.** Both schemes decrypt files that Office, Apache POI and msoffcrypto-tool encrypted, and our own output decrypts with msoffcrypto. The corpus's only PowerPoint-encrypted deck has an undocumented password, so it tests detection and wrong-password rejection only. |

## Not run: the PowerPoint oracles

The plan checks fidelity against Microsoft's own software. None of that ran here:
- **O1**, the Open XML SDK validator, needs the .NET SDK (`tools/validator`).
- **O2**, "PowerPoint opens it without a repair prompt", and **O3/O4**, pixels and text metrics against
  PowerPoint's export, need the Windows runner with Microsoft 365 (`tools/corpus/oracle.ps1`).
- The generated corpus (every enumeration value, written by PowerPoint) and the video comparison of animations need
  that runner too.

The server has the oracle endpoints (`/api/oracle/*`, enabled with `ENABLE_ORACLES=1`). Pixel comparison and the
round-trip check work as they are. The validator needs `PPTX_VALIDATOR`, and the PowerPoint ones need a bridge to the
Windows runner (`POWERPOINT_BRIDGE_URL`); until then they answer 501.
PowerPoint for Mac on the development machine could stand in for O2–O4 locally, but only by automating the
desktop app, which has not been done without the owner's go-ahead.

## Where the code departs from the plan, and why

Phase 1 changes, made after testing the plan's code against real files:

1. **Processing instructions are accepted.** Office starts every SharePoint `customXml` part with
   `<?mso-contentType?>`. PIs are inert, so they are kept; a DOCTYPE is still rejected, which rules out entity
   expansion and XXE.
2. **Line ends and attribute whitespace are normalized** as XML 1.0 requires, and CR is escaped on output.
3. **Entities are looked up in a `Map`.** With an object literal, `&constructor;` decoded to
   `Object.prototype.constructor`.
4. **Escaping matches PowerPoint:** `&#xA;` in attribute values, `"` left raw in text.
5. **Namespaces are resolved,** so views match Strict files and unusual prefixes.
6. **ZIP limits are checked before anything is inflated,** and duplicate names are caught in the filter.
7. **`[Content_Types].xml` is preserved** as written unless parts are added or removed.
8. **ZIP output is deterministic,** and already-compressed media is stored, not deflated again.

Later phases:

1. **Undo restores snapshots instead of running inverse commands.** Each command declares the parts and elements it
   touches; the history keeps those as they were and puts them back. One mechanism serves every command, so undo
   cannot drift from what a command did. Coalescing (typing, nudging) merges consecutive snapshots.
2. **EMF and WMF are played in TypeScript, in a worker,** instead of being rasterized by LibreOffice on the server.
   Previews then work offline and in the editor. TIFF, which only Safari decodes, is decoded by the same package. The
   plan's rule that pure-JS decoders are only for formats the browser cannot decode still holds.
3. **Effect presets come from the corpus** (`tools/anim/gen-presets.ts`): the timing trees PowerPoint wrote for each
   preset, not trees generated through PowerPoint's COM interface.
4. **Storage is the file system behind interfaces** (`apps/server/src/storage.ts`), not S3 and Postgres. Jobs run in
   in-process queues and worker pools rather than BullMQ or pg-boss.
5. **PDF export is raster:** each page is the slide's PNG, 1,920 pixels wide. Vector PDF needs a second backend.
6. **Fonts:** Office's fonts cannot be shipped, so text is shaped with metric-compatible substitutes (Carlito for
   Calibri, Caladea for Cambria, Liberation for Arial, Times New Roman and Courier New), as LibreOffice does.
7. **SmartArt is drawn from the drawing PowerPoint caches in the file,** not laid out from its data. A diagram
   without a cached drawing, or with an empty one, shows its text as an outline.
8. **A SmartArt text box's `rot` adds to its shape's,** as in LibreOffice; when `rot` is missing, the text turns
   with the shape or stays upright, whichever fits it. The files in the corpus disagree, so this is a fit test.

## Known gaps

- **Charts:** 3-D charts are drawn flat, and surface charts as a placeholder.
- **Pictures:** TIFF with JPEG or fax compression, and the HD Photo (`.wdp`) layers behind artistic effects, are not
  decoded. The flattened result that PowerPoint stores beside them is what is drawn.
- **Animations:** a shape's background animated apart from its text fades rather than moves.
- **ZIP:** CRCs are not checked on read (fflate does not expose them), and a truncated archive with no central
  directory is refused, not repaired.
- **Corpus:** it is third-party test files. Most were written by PowerPoint, but some come from LibreOffice, POI and
  test authors' hands.

## Dependencies

- **Parser packages:** they depend only on `fflate` 0.8.3.
- **Other packages:**
  - `text` uses `harfbuzzjs`, `linebreak` and `bidi-js`;
  - `crypto` uses `cfb`;
  - the web app uses React 19.3, Konva, react-konva and Zustand;
  - the server uses Fastify 5 and its plugins, Playwright, pixelmatch, pngjs and TypeBox.
- **Policy:** every version is pinned exactly, and `tools/policy` fails the tests if the parser packages gain a
  dependency.
- **Install safety:** `pnpm-workspace.yaml`
  - holds releases younger than 14 days back;
  - refuses provenance downgrades and git or tarball subdependencies;
  - runs no install scripts.

  Its exceptions, each a security fix, are listed there with the advisory.
