# GenOffice Mobile (genoffice-android)

**GenOffice for Android** - a BYOK (bring-your-own-key) AI office suite built as a native Android
app via Capacitor + React + TypeScript. It is the mobile port strategy for
[kuttappu507/genoffice](https://github.com/kuttappu507/genoffice): keep the portable core
(BYOK AI provider layer, document workflows), reimplement the desktop shell for touch.

The UI is modeled directly on the **Microsoft Office mobile apps** (Word / Excel / PowerPoint /
Office hub): brand-colored app bars, tabbed icon ribbons, a white page canvas in Docs, a Name Box +
fx formula bar with sticky A-B-C headers in Sheets, a 16:9 slide canvas with filmstrip in Slides,
and an Office-hub style launcher with real W/X/P file icons.

No account. No login. No telemetry. Your API key and documents never leave your device
(except calls to the AI provider you choose).

## What works today

| App | Features |
| --- | --- |
| **Home** | Office-hub launcher: create-new tiles with Word/Excel/PowerPoint icons, templates, recent + pinned files (rename / duplicate / delete), open-file banner, first-run onboarding, light / dark / system theme |
| **AI Chat** | Streaming responses, any OpenAI-compatible model, collapsible thinking section, auto-trimmed history, stop button, multiple saved conversations |
| **Docs** (Word-blue) | Home / Insert / Layout / Review / View ribbons with live format states; fonts, sizes, styles, colour + highlight palettes, super/subscript, lists, indents, alignment, line spacing; tables, images, links, page breaks, footnotes, table of contents, symbols; page setup (A4 / Letter / Legal, orientation, margins), headers / footers, page numbers; find & replace, comments, word count, spell-check toggle, reading + dark canvas view; open **.docx** / .txt / .md / .html, save as **.docx**, **.pdf**, **.md**, .html; AI continue / summarize / rewrite; autosave + undo |
| **Sheets** (Excel-green) | Name Box + fx formula bar with function picker, sticky headers, frozen panes, multi-sheet tab strip; 75 functions (`SUM`, `IF`, `VLOOKUP`, `INDEX/MATCH`, `SUMIFS`, `TEXT`, dates, …) with cross-sheet references (`=Sheet2!A1`, `='My Sheet'!B2:B9`); number / currency / percent / date formats, bold / italic / colours / fills / borders / wrap / merge, column widths + row heights, sort, filter, find & replace, fill series, charts (column / bar / line / pie / area), cell notes, go-to; open **.xlsx / .xls / .csv / .ods** with formulas, styles, merges, widths and freeze panes; save real **.xlsx** (all sheets, live formulas, fonts / fills / borders / alignment / freeze panes) or **.csv**; AI table generation |
| **Slides** (PPT-orange) | 16:9 canvas with shape-level editing (text, images, geometric shapes; move / resize, fill, line, font, alignment), 6 themes + layouts, background colours, speaker notes, transitions (fade / push / zoom / flip), slide sorter, filmstrip add / duplicate / delete / reorder; presenter mode with swipe, tap zones, timer, notes, laser pointer; open **.pptx** with positions / styles / images / notes; save as **.pptx**, **.pdf** or PNG; AI outline generation |
| **Settings** | Provider switch (OpenRouter / NVIDIA NIM / custom), live model catalogue, "Test connection" ping with latency, theme, haptics, currency, default font, backup export / import, clear data |

Everything is stored locally on the device (localStorage in the app sandbox). AI features need
your own key:

- **OpenRouter** - free models available (any id ending in `:free`), get a key at <https://openrouter.ai/keys>
- **NVIDIA NIM** - free tier at <https://build.nvidia.com>

## Open and edit real Office files

Tap **Open file** on Home (or **Open** inside Docs / Sheets / Slides) and pick a file:

| Format | Open | Save |
| --- | --- | --- |
| .docx | yes - headings, inline formatting, lists, tables, links, images, footnotes, page breaks, alignment | yes - real .docx (docx library) incl. headers / footers / page numbers |
| .xlsx / .xls / .csv / .ods | yes - every sheet, formulas, number formats, fonts / fills / borders / alignment, merges, column widths, frozen panes | yes - real .xlsx with live formulas, cached values and cell styles; .csv |
| .pptx | yes - shapes with positions, text runs, images, backgrounds, notes | yes - real .pptx (pptxgenjs); .pdf; .png |
| .txt / .md / .html | yes | .docx / .pdf / .md / .html |

On Android, saving opens the **system share sheet** so you can send the file to Drive, mail it,
or store it anywhere. On desktop browsers it downloads directly. Parsing and writing happen
entirely on-device; nothing is uploaded.

Sample files to try: run `node scripts/make-samples.mjs` to generate `samples/sample.docx / .xlsx / .pptx`.

## Built-in token savings (free-tier friendly)

- **Think-tag stripping** - models like Nemotron leak `<think>...</think>` blocks; they are
  filtered out of the visible answer and **never sent back** to the API
- **`reasoning.exclude`** - enabled by default on OpenRouter so reasoning tokens are not returned
- **History trimming** - chat history sent per request is capped (about 60K chars on free tiers),
  oldest messages dropped first, system prompt always kept
- **429 backoff with Retry-After** - up to 4 attempts, honors the server's `Retry-After`
  (capped at 20 s), then exponential 1/2/4 s
- **Friendly error mapping** - 401/402/404/429 become human-readable hints instead of raw dumps

## Porting notes (what carried over, what did not)

The original GenOffice is an Electron desktop app (main process + preload IPC + React renderers).
Electron does not run on Android, so this port:

- **Reused in spirit**: BYOK provider layer design, document models (doc/sheet/deck/chat),
  token-optimization ideas, the "local-first, no-login" philosophy
- **Reimplemented for touch**: the shell (tab navigation), editors, file storage
  (app sandbox instead of desktop file system)
- **Dropped**: Electron main process, window/menu system, desktop updater, Genspark CLI
  integrations (the desktop fork is removing those anyway)

## Build an APK

### Option A - GitHub Actions (no local setup)

Push to `main` and the included workflow (`.github/workflows/android.yml`) builds a debug APK
and uploads it as an artifact. Or trigger it manually from the **Actions** tab
(**Run workflow**).

### Option B - locally

```bash
npm install
npm test               # formula engine, sheet model, docx/xlsx/pptx round-trip checks (Node only)
npm run build          # type-check + vite build into dist/
npx cap add android    # only needed once; the repo may already contain android/
npx cap sync android
npx cap open android   # opens Android Studio, then Run
```

Requirements: Node 20+, Java 17 (Temurin), Android SDK (Android Studio handles this).

The debug APK ends up at `android/app/build/outputs/apk/debug/app-debug.apk`.

## Project structure

```
src/
  main.tsx            entry point
  App.tsx             shell: icon bottom nav, per-screen Office accent, full-screen editors, back-button handling
  types.ts            shared types
  components/
    Icon.tsx          stroke SVG icon set + Office file-type icons (W/X/P)
    Ribbon.tsx        Office-mobile ribbon primitives (tabs, groups, buttons, steppers, palettes)
    Sheet.tsx         bottom sheets, menus, prompts, toasts
    ChartView.tsx     SVG charts for Sheets
    FunctionPicker.tsx  searchable function catalogue for the formula bar
    SlideView.tsx     slide renderer (canvas, thumbnails, presenter, PNG/PDF export)
  lib/
    ai-client.ts      BYOK streaming client: SSE, think-strip, 429 backoff, test ping
    models.ts         provider presets, free-tier model list, context budgets
    storage.ts        local document store, prefs, theme, backup export/import
    native.ts         Capacitor glue: haptics, status bar, back button, share, keep-awake
    formulas.ts       spreadsheet formula engine (75 functions, cross-sheet refs)
    sheet-model.ts    workbook model: styles, merges, freeze, filter, charts, structural edits, templates
    deck-model.ts     slide model: themes, layouts, shapes, outline -> deck
    fileio.ts         .docx / .xlsx / .pptx / .pdf / .md import + export
    xlsx-styles.ts    cell-style + freeze-pane round trip for .xlsx (patches OOXML parts)
    markdown.tsx      tiny markdown renderer for chat answers
  screens/            one file per app screen (Office-mobile style)
scripts/
  test-formulas.mjs   formula engine checks
  test-sheet-model.mjs  cross-sheet evaluation, structural edits, CSV
  test-io.mjs         docx / xlsx / pptx export -> import round trips (jsdom)
  test-pptx-shapes.mjs  shape-level pptx import/export
android/              native Android shell (Capacitor)
.github/workflows/    tests + APK build CI
```

## Roadmap

- PDF viewer (pdf.js) for opening existing PDFs
- Open password-protected Office files
- Conditional formatting and pivot-style summaries in Sheets
- Release signing + Play Store pipeline
- iOS target from the same code base (`npx cap add ios`)
- Optional sync of the desktop fork's document format

## License

MIT - see [LICENSE](LICENSE).
