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
| **Home** | Office-hub launcher: create-new tiles with Word/Excel/PowerPoint icons, recent files, open-file banner |
| **AI Chat** | Streaming responses, any OpenAI-compatible model, collapsible thinking section, auto-trimmed history, stop button, multiple saved conversations |
| **Docs** (Word-blue) | Blue app bar + Home/Insert/AI icon ribbon with live format states, font-color & highlight palette, white page canvas, open/edit **.docx** (mammoth), .txt, .md, .html; save as real **.docx**; AI continue/summarize/rewrite; autosave |
| **Sheets** (Excel-green) | Green app bar, Name Box + fx formula bar, sticky A-Z / 1-60 headers with green selection + fill handle, Sheet1 tab strip; open/edit **.xlsx / .xls / .csv** (SheetJS) with formulas preserved; save as real **.xlsx** with live formulas + cached values; formula engine (`=A1+B2*2`, `SUM/AVERAGE/COUNT/MIN/MAX/ABS/ROUND`, ranges); AI table generation |
| **Slides** (PPT-orange) | Orange app bar with present + save actions, 16:9 slide canvas with inline title/bullet editing, filmstrip with add/duplicate/delete/reorder; open **.pptx** (JSZip); save as real **.pptx** (pptxgenjs); fullscreen present mode; AI outline generation |
| **Settings** | Provider switch (OpenRouter / NVIDIA NIM / custom), free-tier model presets, "Test connection" ping with latency, backup export/import |

Everything is stored locally on the device (localStorage in the app sandbox). AI features need
your own key:

- **OpenRouter** - free models available (any id ending in `:free`), get a key at <https://openrouter.ai/keys>
- **NVIDIA NIM** - free tier at <https://build.nvidia.com>

## Open and edit real Office files

Tap **Open file** on Home (or **Open** inside Docs / Sheets / Slides) and pick a file:

| Format | Open | Save |
| --- | --- | --- |
| .docx | yes - converted to editable rich text | yes - real .docx (docx library) |
| .xlsx / .xls / .csv | yes - first sheet, formulas kept | yes - real .xlsx with live formulas |
| .pptx | yes - titles + bullets extracted | yes - real .pptx (pptxgenjs) |
| .txt / .md / .html | yes | .docx / .html |

On Android, saving opens the **system share sheet** so you can send the file to Drive, mail it,
or store it anywhere. On desktop browsers it downloads directly. Parsing and writing happen
entirely on-device; nothing is uploaded.

Sample files to try: `download/genoffice-ui/samples/sample.docx|xlsx|pptx` in this project
(regenerable via `node scripts/make-samples.mjs`).

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
  App.tsx             shell: icon bottom nav, per-screen Office accent, full-screen editors
  types.ts            shared types
  components/
    Icon.tsx          stroke SVG icon set + Office file-type icons (W/X/P)
  lib/
    ai-client.ts      BYOK streaming client: SSE, think-strip, 429 backoff, test ping
    models.ts         provider presets, free-tier model list, context budgets
    storage.ts        local document store, settings, backup export/import
    formulas.ts       spreadsheet formula engine
    markdown.tsx      tiny markdown renderer for chat answers
  screens/            one file per app screen (Office-mobile style)
android/              native Android shell (Capacitor)
.github/workflows/    APK build CI
```

## Roadmap

- PDF viewer (pdf.js)
- More ribbon tabs (Layout, Review) and full font controls (size, family)
- Multi-sheet workbooks
- Open password-protected Office files
- Release signing + Play Store pipeline
- iOS target from the same code base (`npx cap add ios`)
- Optional sync of the desktop fork's document format

## License

MIT - see [LICENSE](LICENSE).
