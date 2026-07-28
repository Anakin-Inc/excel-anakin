# excel-anakin

Microsoft Excel Office Add-in for [Anakin](https://anakin.io) — web scraping,
AI-powered search, and multi-stage agentic research, available both as a task
pane and as native worksheet formulas (`=ANAKIN.SCRAPE`, `=ANAKIN.SEARCH`,
`=ANAKIN.AGENTICSEARCH`).

## What's covered

The same baseline every other single-purpose Anakin integration built this
session covers:

- **Scrape** — submits `POST /v1/url-scraper`, polls `GET /v1/url-scraper/:jobId`
  until the job completes or fails. Returns markdown.
- **Search** — calls `POST /v1/search`. Synchronous, no polling. Costs 3
  credits per call.
- **Agentic search** — submits `POST /v1/agentic-search`, polls
  `GET /v1/agentic-search/:jobId`. Returns an AI-generated summary (plus
  structured data, when the job's schema produced any). Costs 10 credits.

Each is exposed twice, on two different Excel surfaces:

| Surface | How | Where |
| --- | --- | --- |
| Task pane | Buttons that call `Excel.run(async (context) => ...)` and write results into the selected cell(s) | `src/taskpane/taskpane.js` |
| Custom functions | `=ANAKIN.SCRAPE(url)`, `=ANAKIN.SEARCH(query)`, `=ANAKIN.AGENTICSEARCH(prompt)`, typed directly into any cell | `src/functions/functions.js` |

## Why both surfaces

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
