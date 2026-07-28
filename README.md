# excel-anakin

Microsoft Excel Office Add-in for [Anakin](https://anakin.io) — web scraping,
AI-powered search and research, site mapping/crawling, Wire automation
discovery and read actions, AI-visibility comparison, and monitor/session
inspection as native worksheet formulas (`=ANAKIN.SCRAPE`, `=ANAKIN.SEARCH`,
`=ANAKIN.MAP`, …), plus a task pane for everything state-changing (Wire write
actions, sign-in, monitor create/control, session delete, browser
automation).

Covers 20 of Anakin's 21 API capabilities — every one except
`wire_identities` (credential listing), left out deliberately rather than by
oversight; see "What's not covered" below.

## What's covered

### Read-only / idempotent → Custom Functions

Safe to live in a cell formula: Excel can recalculate these automatically
(on open, on a referenced cell changing, on manual recalc) with no
side effect beyond re-fetching data.

| Function | Endpoint(s) | Returns |
| --- | --- | --- |
| `=ANAKIN.SCRAPE(url)` | `POST /url-scraper` → poll `GET /url-scraper/:jobId` | Markdown (scalar) |
| `=ANAKIN.SEARCH(query)` | `POST /search` (sync) | `[url, title, snippet]` rows |
| `=ANAKIN.AGENTICSEARCH(prompt)` | `POST /agentic-search` → poll `GET /agentic-search/:jobId` | AI summary (scalar) |
| `=ANAKIN.MAP(url)` | `POST /map` → poll `GET /map/:jobId` | One discovered URL per row |
| `=ANAKIN.CRAWL(url)` | `POST /crawl` → poll `GET /crawl/:jobId` | `[url, status, markdown]` rows |
| `=ANAKIN.WIREDISCOVER(query)` | `GET /wire/resolve` (sync) | `[action_id, type, details]` rows |
| `=ANAKIN.WIRECATALOG([slug])` | `GET /wire/catalog[/:slug]` (sync) | Raw JSON (scalar) |
| `=ANAKIN.WIREREADACTION(actionId, [paramsJson])` | `POST /wire/task` → poll `GET /wire/jobs/:jobId` | `[field, value]` rows from the action's extracted data |
| `=ANAKIN.AIVISIBILITYSEARCH(query)` | `POST /ai-visibility/search` → poll `GET /ai-visibility/search/:id` | Per-engine result rows + a synthesis row |
| `=ANAKIN.AIVISIBILITYSOURCES()` | `GET /ai-visibility/sources` (sync) | `[slug, label]` rows |
| `=ANAKIN.SESSIONLIST([domain])` | `GET /sessions` (sync) | One JSON-encoded session per row |
| `=ANAKIN.MONITORLIST([id])` | `GET /monitors[/:id]` (sync) | One JSON-encoded monitor per row |
| `=ANAKIN.MONITORCHANGES(id)` | `GET /monitors/:id/changes` (sync) | One JSON-encoded change record per row |

Several of these (`wire_catalog`, `wire_read_action`'s `data`, `session_list`,
`monitor_list`, `monitor_changes`) return API responses whose *item* shape
isn't documented in `anakin-mcp`'s TypeScript types (`Promise<unknown>`) —
rather than guess field names, those functions surface the real JSON
verbatim (as a scalar or one JSON string per row) instead of a fabricated
column layout. `=ANAKIN.WIREDISCOVER` and `=ANAKIN.AIVISIBILITYSOURCES` use
named columns because their fields (`action_id`/`type`, `slug`/`label`) are
explicitly documented in `anakin-mcp/src/tools/wire.ts` and `client.ts`.

### State-changing → task pane only, never a formula

| Action | Endpoint(s) | Task pane control |
| --- | --- | --- |
| Run a Wire write action | `POST /wire/task` → poll `GET /wire/jobs/:jobId` | "Run write action" |
| Sign in to a Wire site | `POST /wire/login` (sync) | "Sign in" |
| Request a new Wire action | `POST /wire/build-request` (sync) | "Request build" |
| Create a monitor | `POST /monitors` (sync) | "Create monitor" |
| Control a monitor (pause/resume/run_now/delete) | `POST/DELETE /monitors/:id/...` | "Apply" |
| Delete a saved session | `DELETE /sessions/:id` | "Delete session" |
| Run an AI browser task | `POST /ai/evaluate` (async) → poll `GET /ai/jobs/:id` | "Run task into selected cell" |

Every one of these is exposed *only* as a task-pane button (an explicit,
one-shot user click, confirmed with `window.confirm()` before it runs) —
never as a Custom Function. See "Read vs. write: why the split" below.

Each surface's code lives at:

| Surface | How | Where |
| --- | --- | --- |
| Task pane | Buttons that call `Excel.run(async (context) => ...)` and write results into the selected cell(s) | `src/taskpane/taskpane.js` |
| Custom functions | `=ANAKIN.SCRAPE(url)`, `=ANAKIN.MAP(url)`, … typed directly into any cell | `src/functions/functions.js` |
| Shared HTTP client | Every `/v1/*` request/response shape used by both surfaces above | `src/utilities/api.js` |

## What's not covered

- **`wire_identities`** (`GET /wire/identities`) — lists saved credentials.
  Secondary to `wire_login` (which returns a fresh `credential_id` directly)
  and to the dashboard's own credential view; adding a formula or button
  whose only output is an ID to paste into another field wasn't judged
  worth the added surface for this pass. Straightforward to add later
  following the `SESSIONLIST` pattern (same "unknown item shape → one JSON
  row per credential" design) if needed.

That's 3 (baseline: scrape/search/agentic search) + 10 new Custom Functions
+ 7 new task-pane actions = 20 of 21 covered.

## Read vs. write: why the split

Office.js Custom Functions have real execution constraints, confirmed live
against Microsoft's docs (not assumed):

- Custom functions that call an external API "return a JavaScript `Promise`"
  and Excel "automatically waits for the promise to resolve before
  displaying the result in the cell" — this is the documented, intended
  pattern for exactly what `SCRAPE`/`MAP`/`CRAWL`/etc. do here (submit +
  poll inside the returned promise). Streaming functions (`@streaming`,
  repeated `setResult`) are a separate mechanism for continuously-updating
  values (e.g. a live stock price) and weren't needed for anything Anakin
  exposes — every new function here has a single terminal result, matching
  the existing `SCRAPE`/`SEARCH`/`AGENTICSEARCH` shape.
  (learn.microsoft.com/office/dev/add-ins/excel/custom-functions-web-reqs,
  read live.)
- Cancellation is real but coarse: Excel calls a function's `onCanceled`
  when "the user edits or deletes a cell that references the function" or
  "one of the arguments... changes" — and explicitly does **not** guarantee
  ordering between the old call's cancellation and a new invocation
  ("Your add-in code should not depend on `onCanceled` firing before the
  next invocation"). Nothing in the Custom Functions runtime provides a
  confirmation step, an "are you sure" gate, or any way to distinguish a
  user-intended recalculation from Excel's own automatic one.
- Anakin's own MCP tool definitions (`anakin-mcp/src/tools/*.ts`) mark
  `wire_write_action`, `wire_login`, `wire_build`, `monitor_create`,
  `monitor_control`, `session_delete`, and `browser_task` **all** with
  `destructiveHint: true` — Anakin's own engineers already decided these
  seven need a human in the loop before they run. A live formula is the
  opposite of that: it can fire on workbook open, on a cell it references
  changing, or on `Ctrl+Alt+F9`, with nobody watching. Each of the seven
  runs from Wire's own endpoint or resource: signing in, submitting a form
  or adding to a cart, requesting/publishing a new scraper, creating or
  deleting a recurring billed job, deleting stored login state, or driving
  a real browser that can change state on a third-party site — none of
  those should re-fire silently.

So the split is structural, not stylistic: **read-only and idempotent →
Custom Function; anything `destructiveHint: true` → task-pane button behind
an explicit click and a `window.confirm()`.** This mirrors exactly how
Anakin's own MCP server (`anakin-mcp`) splits `wire_read_action` from
`wire_write_action` into two separate tools with different safety
annotations for the same underlying reason (see the comment atop
`anakin-mcp/src/tools/wire.ts`).

## Why both surfaces (for the original three)

Excel add-ins support several extension points (task panes, custom functions,
ribbon commands, dialogs, content add-ins — see
[Build Excel add-ins with Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-overview),
"What you can build" section, read live). For a scraping/search API, the two
that fit are a task pane (for one-off, interactive requests where the result
should land in a specific cell the user is looking at) and custom functions
(for the idiomatic spreadsheet case: `=ANAKIN.SEARCH("competitor pricing")`
recalculating like any other formula, composable with other cells —
`=ANAKIN.SCRAPE(A1)` where `A1` holds a URL from another formula). Custom
functions support async JS natively (a function can `return` a `Promise`, or
be `@streaming`/`@cancelable` — see
[Create custom functions in Excel](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-overview),
read live), so the same submit-then-poll pattern the task pane uses for
`SCRAPE`/`AGENTICSEARCH` works unchanged inside a custom function.

The two surfaces share one API key and one HTTP client (`src/utilities/api.js`,
`src/utilities/state.js`) via Excel's **shared runtime** — see "Manifest
format" below for why that's structural, not a style choice.

## Manifest format: XML (add-in-only), not the unified JSON manifest

Confirmed live against Microsoft's current docs, not assumed:

- Microsoft's Yeoman generator overview states, of the JSON unified-manifest
  project option: **"The unified manifest for Microsoft 365 is in preview for
  Excel, PowerPoint, and Word add-ins. It should not be used for production
  add-ins."** (learn.microsoft.com/office/dev/add-ins/develop/yeoman-generator-overview,
  "Word, PowerPoint, or Excel with unified manifest option" section.)
- The client/platform support table on
  learn.microsoft.com/office/dev/add-ins/develop/unified-manifest-overview
  confirms unified-manifest add-ins are *not* supported on Office on Windows
  (perpetual versions) or Office on mobile, and require recent-enough
  Microsoft 365 builds even where supported.
- The unified manifest **is** production-ready for Outlook specifically — the
  Yeoman generator's own Outlook-only follow-up question recommends it
  ("unless your add-in will include an extensibility feature that isn't yet
  supported"). Excel is the exception, not the rule.

So `manifest.xml` here is the classic add-in-only XML manifest — the format
every current Excel add-in, including Microsoft's own reference samples,
still ships in production. Its `Hosts`/`Requirements`/`VersionOverrides`/
`ExtensionPoint` shape, the `SharedRuntime` requirement set, and the
`CustomFunctions` extension point fields were copied structurally from the
real, current `OfficeDev/Office-Add-in-samples` reference add-in at
`Samples/excel-shared-runtime-scenario/manifest.xml`, pulled live via
`raw.githubusercontent.com` — not guessed. See SUBMIT.md for exactly what was
checked.

## Why a shared runtime

By default, a task pane and its custom functions each get an *isolated* JS
runtime — no shared state. Anakin needs the API key entered once in the task
pane to be usable by `ANAKIN.*` formulas in cells, so this add-in declares
the `SharedRuntime` requirement set and points the `CustomFunctions`
extension point's `Script`/`Page` at the *same* URL as the task pane
(`taskpane.html`) instead of a separate `functions.html`. Under a shared
runtime, that one page's JS context hosts both the task pane UI and the
custom functions — confirmed against the real reference sample, where
`ContosoAddin.Url` (the custom-functions `Script`/`Page` resource) resolves
to the same `taskpane.html` as the task pane's own `SourceLocation`, and
`taskpane.html` loads `functions.js` via a plain `<script>` tag alongside
`taskpane.js`. `src/utilities/state.js` is the shared global (`AnakinState`)
that makes the API key visible to both.

## Install / build

```sh
npm install
```

Installs the real, current `office-addin-manifest` (manifest validator),
`office-addin-debugging` (sideload/debug CLI), and `office-addin-lint` — see
SUBMIT.md for which of these actually ran in this sandbox.

This project is hand-authored rather than scaffolded by `yo office`, so there
is no webpack build step and `functions.json` was written by hand instead of
generated by the `custom-functions-metadata-plugin` Webpack plugin Microsoft
recommends for production projects (confirmed at
learn.microsoft.com/office/dev/add-ins/excel/custom-functions-json-autogeneration).
It's also a real, documented alternative — the "Manually create JSON metadata
for custom functions" doc page exists specifically for hand-written
`functions.json`, and its JSON metadata example was used as the schema
reference for the one in `src/functions/functions.json`.

## Auth

Anakin uses a static API key sent as the `X-API-Key` header, entered once in
the task pane's **Anakin API key** field and persisted with
`Office.context.document.settings` (roams with the workbook). Get a free key
(300 credits, no card required) at
[anakin.io/dashboard](https://anakin.io/dashboard).

## CORS

Because the task pane and custom functions call `api.anakin.io` directly from
the browser/webview (`fetch()` in `src/utilities/api.js`), this depends on
Anakin's API allowing cross-origin requests — confirmed with a live CORS
preflight against `https://api.anakin.io/v1/search`, which returned
`access-control-allow-origin: *`. See SUBMIT.md for the full request/response.

## Steps (needs the account owner)

1. Host `manifest.xml` and everything under `src/` and `assets/` over HTTPS —
   this submission uses `https://excel-addin.anakin.io/...` as the hosting
   domain placeholder throughout (matching the `anakin.io`-subdomain pattern
   other Anakin integrations in this batch use); an actual domain needs to
   be provisioned and the URLs in `manifest.xml` updated to match.
2. Supply real icon assets (`assets/icon-16.png`, `icon-32.png`, `icon-80.png`)
   — `manifest.xml` references them but no Anakin-branded PNGs exist in this
   sandbox.
3. `npm install && npx office-addin-manifest validate manifest.xml` against
   the real hosted URLs (see SUBMIT.md for what validating the placeholder
   URLs in this sandbox already caught and didn't catch).
4. Sideload for testing: `npx office-addin-debugging start manifest.xml` (or
   Excel's **Insert > Add-ins > Upload My Add-in** with the hosted
   `manifest.xml`), then confirm the API key round-trips, `ANAKIN.SEARCH`
   resolves in a cell, and `Excel.run` writes land where expected — none of
   this was exercised against a running Excel instance in this sandbox (see
   SUBMIT.md, "Not done").
5. Submit to Microsoft Marketplace via **Make your solutions available in
   Microsoft Marketplace** (Partner Center) once ready for public listing —
   a separate manual process from building/hosting the add-in itself.
