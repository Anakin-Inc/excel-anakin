# Excel — submission instructions

Real, but not self-serve the way Coda was: Excel (Office) add-ins are
submitted to **Microsoft Marketplace** through **Partner Center**
(learn.microsoft.com/office/dev/add-ins/excel/excel-add-ins-overview, "What
an Excel add-in includes" section — "publish its manifest to Microsoft
Marketplace or to an add-ins catalog... see Make your solutions available in
Microsoft Marketplace and within Office"). That's a manual, account-gated
submission through a web portal, not a CLI push like Coda's `packs upload` or
a GitHub PR like Dagster's community-integrations repo. No CLI in the
`office-addin-*` toolchain performs that submission step — the CLIs (below)
only build, sideload, and validate.

## What's here

```
manifest.xml              Add-in-only (XML) manifest — see "Manifest format" below
src/taskpane/              Task pane: taskpane.html, taskpane.js, taskpane.css
src/functions/              Custom functions: functions.js, functions.json
src/utilities/              Shared code: api.js (Anakin HTTP client), state.js (shared API-key state)
assets/                     Empty — real icon PNGs still needed, see "Not done"
package.json, package-lock.json   Real office-addin-* devDependencies, from a real npm install
.gitignore                  node_modules/, dist/, *.log
```

## Full API coverage (feat/full-tool-coverage)

The initial submission covered 3 of Anakin's 21 REST API capabilities
(scrape, search, agentic_search). This branch adds 10 more Custom Functions
(MAP, CRAWL, WIREDISCOVER, WIRECATALOG, WIREREADACTION, AIVISIBILITYSEARCH,
AIVISIBILITYSOURCES, SESSIONLIST, MONITORLIST, MONITORCHANGES) and 7 new
task-pane-only actions (wire_write_action, wire_login, wire_build,
monitor_create, monitor_control, session_delete, browser_task) — 20 of 21
covered; `wire_identities` deliberately left out. See README.md's "What's
covered" / "What's not covered" / "Read vs. write: why the split" sections
for the full breakdown and the reasoning, and the comment atop
`src/functions/functions.js` for why the 7 write actions are task-pane-only.

Ground truth for every new endpoint (request bodies, response shapes, job
polling behavior, credit-cost notes) was read directly from Anakin's own MCP
server source — `anakin-mcp/src/client.ts` and `anakin-mcp/src/tools/*.ts`
— not guessed. Where the source itself types a response as `unknown` (Wire
catalog browsing, a Wire action's extracted `data`, session/monitor list
items), the corresponding function surfaces the real JSON rather than
inventing field names — see the README's "What's covered" table notes.

The read-vs-write execution-constraint claims (Custom Functions "return a
Promise" and Excel awaits it; `onCanceled` ordering is not guaranteed;
streaming functions are a separate mechanism for continuously-updating
values) were confirmed live against
`learn.microsoft.com/office/dev/add-ins/excel/custom-functions-web-reqs` and
`.../custom-functions-parameter-options`, not assumed. Microsoft's docs
don't state a "don't put side effects in a custom function" rule in so many
words; the design instead follows Anakin's own MCP tool annotations
(`destructiveHint: true` on all 7 of the write-side tools, in
`anakin-mcp/src/tools/*.ts`), which is the actual, load-bearing signal that
those 7 operations need a human-confirmed, one-shot trigger rather than a
formula Excel can silently re-fire.

Hand-authored, not `yo office`-scaffolded — there's no webpack build step, and
`functions.json` was written by hand against the real "Manually create JSON
metadata for custom functions" reference doc's schema (its own text
recommends the Webpack `custom-functions-metadata-plugin` for production
projects; hand-writing `functions.json` is the documented fallback for
projects built from scratch).

## Two real design questions, answered from source

**1. XML manifest or the newer unified JSON manifest?** The task's premise
was that the unified manifest "now supersedes the older XML one" — checked
live and found **false for Excel specifically**. Microsoft's own Yeoman
generator docs say, of the JSON-unified-manifest project option:
> "The unified manifest for Microsoft 365 is in preview for Excel,
> PowerPoint, and Word add-ins. It should not be used for production
> add-ins."
(learn.microsoft.com/office/dev/add-ins/develop/yeoman-generator-overview,
"Word, PowerPoint, or Excel with unified manifest option" section, fetched
live.) The unified manifest *is* production-ready for Outlook (the
generator's Outlook flow recommends it by default) — Excel/PowerPoint/Word
are the exception. The client/platform support table on
`develop/unified-manifest-overview` (also fetched live) confirms unified
add-ins aren't supported on Office on Windows perpetual or Office on mobile
even where the format is otherwise usable. So `manifest.xml` uses the
add-in-only XML format — real, current, production, and what Microsoft's own
reference samples still ship for Excel.

**2. Task pane, custom functions, or both?** The task's premise (Custom
Functions might be "a better fit... since users might want
`=ANAKIN.SEARCH(...)` directly in a cell") checked out, but the interesting
finding was *how* to wire the two together, which isn't spelled out in prose
docs — only visible by reading real source. The custom-functions-overview doc
says a shared runtime "allows for the sharing of data between functions and
the task pane" but doesn't show the mechanism. Reading the real reference
sample (`OfficeDev/Office-Add-in-samples`,
`Samples/excel-shared-runtime-scenario/`) live via `raw.githubusercontent.com`
showed it directly: the manifest's `CustomFunctions` `Script` and `Page`
elements both resolve to the *same URL* as the task pane's own
`SourceLocation` (`ContosoAddin.Url` in that sample), and `taskpane.html`
loads `functions.js` via a plain `<script>` tag right alongside `taskpane.js`
— under a shared runtime, they're one page, one JS context, sharing globals
directly, not two isolated webviews passing messages. This add-in's
`manifest.xml` and `src/taskpane/taskpane.html` mirror that structure exactly
— `src/utilities/state.js` is the shared global (`AnakinState`) holding the
API key that both the task pane buttons and the `ANAKIN.*` formulas read.

## Verified, not assumed

- **`npm install` succeeded for real** against the live npm registry — 552
  packages, including the real, current `office-addin-manifest@2.1.6`,
  `office-addin-debugging@6.1.2`, and `office-addin-lint@3.0.10` (versions
  confirmed with `npm view` before writing `package.json`, not guessed).
- **`npx office-addin-manifest validate manifest.xml` passes clean** — this
  is Microsoft's own official manifest validator (the same tool Office Store
  submission tooling uses), and it made a real network call to Microsoft's
  hosted "Acceptance test service" as part of validation. Full relevant
  output:
  ```
  Valid Manifest Schema: Your manifest does adhere to the current set of XML
  schema definitions for Add-in manifests.
  ...
  Acceptance Test Completed: Acceptance test service has finished checking
  provided add-in.
  ...
  Based on the requirements specified in your manifest, your add-in can run
  on the following platforms:
    - Excel 2019 or later on Mac
    - Excel on Windows (Microsoft 365)
    - Excel on the web
    - Excel on Mac (Microsoft 365)
  The manifest is valid.
  ```
  Note what it does *not* claim: Excel on Windows (perpetual) and mobile are
  absent from that list — expected, since the `SharedRuntime` requirement set
  this add-in declares isn't supported on those platforms
  (learn.microsoft.com/javascript/api/requirement-sets, cross-checked).
- **The validator does not catch a wrong-but-well-typed built-in tab ID.**
  `manifest.xml` initially had `<OfficeTab id="TabHome">`; the validator
  passed regardless (schema treats `id` as a plain string, not an enum
  against the real per-app ID list). Cross-checked separately against
  `learn.microsoft.com/office/dev/add-ins/develop/built-in-ui-ids`, which
  confirms `TabHome` actually *is* the correct built-in Home-tab ID for
  Excel (Word/PowerPoint/OneNote share it too — only Page Layout needed an
  Excel-specific suffix, `TabPageLayoutExcel`, to disambiguate from Word's).
  So the value was right, but it's worth flagging that the validator
  wouldn't have caught it if it hadn't been — anyone editing this manifest's
  ribbon surface should check that reference page by hand.
- **`node --check` passes on all four hand-written `.js` files**
  (`src/taskpane/taskpane.js`, `src/utilities/api.js`,
  `src/utilities/state.js`, `src/functions/functions.js`) — re-run after
  `feat/full-tool-coverage`'s additions to `taskpane.js`, `api.js`, and
  `functions.js`; still passes clean.
- **`manifest.xml` parses as well-formed XML** (`python3 -m xml.dom.minidom`)
  and **`functions.json`/`package.json` parse as valid JSON**
  (`python3 -m json`) — re-checked after `feat/full-tool-coverage` added 10
  new function entries to `functions.json` and rewrote `manifest.xml`'s
  `Description`/`GetStarted.Description`/`BtnShowTaskpane.Tooltip` strings
  to mention the new capabilities (the first draft of the new `Description`
  exceeded the schema's `LongString` max length — caught by
  `office-addin-manifest validate`, then shortened to fit; see git history).
- **Every `getElementById` call in the new `taskpane.js` handlers resolves
  to a real element in the new `taskpane.html` sections** — cross-checked
  by diffing the two ID sets, not just eyeballed.
- **`npx office-addin-lint check` ran for real** and flagged 65 real
  findings on the original 3-function submission — almost all either (a)
  browser globals (`window`, `document`, `fetch`, `setTimeout`) reported as
  `no-undef` because this hand-authored project has no ESLint config
  declaring a browser environment (a `yo office`-scaffolded project ships
  one by default; this one doesn't), or (b) Prettier formatting deltas.
  None were logic bugs. Not fixed — left as a known gap rather than
  papering over it; see "Not done." **Re-run after `feat/full-tool-coverage`
  added ~4x the code**: 246 findings now (174 Prettier, 68 `no-undef` on the
  same browser globals — `document`×42, `AbortController`×13, `window`×10,
  `setTimeout`/`fetch`/`clearTimeout`×1 each — plus 4 `no-unused-vars` false
  positives on `Anakin`/`AnakinState`, which the plain-`<script>`-tag shared
  runtime assigns but never `export`s, so a modules-aware linter can't see
  the cross-file usage). Same three categories as before, scaled with the
  code; manually reviewed and confirmed none are logic bugs. Still not
  fixed, for the same reason as before.
- **Live CORS preflight against `https://api.anakin.io/v1/search`** (a real
  `curl -X OPTIONS` with `Origin: https://localhost:3000`) returned
  `HTTP/2 204` with `access-control-allow-origin: *` and
  `access-control-allow-headers` including `X-API-Key`. A follow-up real
  `POST` with an `Origin` header (no API key) returned a real `401` from the
  live API, still with `access-control-allow-origin: *` on the error
  response. Confirms `src/utilities/api.js`'s direct browser-context
  `fetch()` calls to `api.anakin.io` will work with no CORS proxy — this
  matters specifically for a task-pane/custom-functions add-in, where all
  network calls run in a browser/webview, unlike a server-side SDK.
- **`functions.json`'s hand-written schema** (`result`/`parameters`/
  `options.cancelable` shapes) was checked field-by-field against the real
  "Manually create JSON metadata for custom functions" reference doc's own
  JSON example, fetched live — not inferred from the `@customfunction` JSDoc
  alone.
- Response field names used in `api.js`/`functions.js`/`taskpane.js`
  (`jobId`, `generatedJson.summary`, `generatedJson.structured_data`,
  `durationMs`, job `status` values `pending`/`processing`/`completed`/
  `failed`) come directly from `anakin-py`'s `src/anakin/client.py` and
  `src/anakin/models.py` — the reference SDK implementation for this API —
  matching the same source `coda-anakin`'s `pack.ts` used, not guessed
  independently.

## Not done

- **Never sideloaded or ran this add-in in an actual Excel client** — no
  Excel instance (desktop, web, or Mac) available in this sandbox. The task
  pane's `Excel.run(async (context) => ...)` cell-writing code
  (`writeScalarToSelectedCell`/`writeMatrixToSelectedCell` in
  `src/taskpane/taskpane.js`) and the `ANAKIN.SCRAPE`/`ANAKIN.SEARCH`/
  `ANAKIN.AGENTICSEARCH` custom functions were checked by close reading
  against the real Excel JavaScript API and Custom Functions Runtime docs,
  never exercised against a live Excel calc engine.
- **Never tested against a real, valid Anakin API key.** The CORS check
  above confirms reachability and header shape; it does not confirm that a
  real `200` response from `/url-scraper`, `/search`, or `/agentic-search`
  actually deserializes the way `api.js` assumes (e.g. that `job.markdown`
  and `job.generatedJson.structured_data` are really at those paths on a
  genuine completed job, not just in the Pydantic model definitions). Same
  caveat applies to every endpoint `feat/full-tool-coverage` added — `map`,
  `crawl`, `wire/*`, `ai-visibility/*`, `sessions`, `monitors`, and
  `ai/evaluate` request/response shapes were read from `anakin-mcp`'s own
  TypeScript client/tool source (not guessed), but never exercised against
  live API responses.
- **Nothing is hosted.** Every URL in `manifest.xml`
  (`https://excel-addin.anakin.io/...`) is a placeholder matching the
  `anakin.io`-subdomain pattern other integrations in this batch use. Excel
  requires HTTPS-hosted manifest/source files to sideload or install an
  add-in — none of `src/`, `assets/`, or `manifest.xml` are actually served
  from anywhere yet.
- **No real icon assets.** `manifest.xml` references
  `assets/icon-16.png`/`icon-32.png`/`icon-80.png`; `assets/` is empty in
  this submission.
- **`office-addin-lint` findings not fixed** (see above) — flagged and left
  for whoever hosts and finishes this add-in, rather than hand-patching
  Prettier formatting/adding an ESLint browser-env config without being able
  to re-verify the result against a real Excel runtime.
- **Marketplace/Partner Center submission itself never attempted** — needs a
  Microsoft Partner Center developer account, which doesn't exist in this
  sandbox; see "Steps" below.

## Steps (needs the account owner)

1. Provision HTTPS hosting for this add-in (e.g. an `excel-addin.anakin.io`
   subdomain) and update every URL in `manifest.xml` to match; supply real
   `icon-16.png`/`icon-32.png`/`icon-80.png` brand assets into `assets/`.
2. `npm install && npx office-addin-manifest validate manifest.xml` again
   against the real hosted URLs — this sandbox only validated the
   placeholder URLs (which the validator accepts, since it checks URL
   *shape*, not that something live is actually there yet).
3. Sideload for real testing: `npx office-addin-debugging start manifest.xml`
   (or Excel's **Insert > Add-ins > Upload My Add-in**) and confirm, against
   a real Excel session: the API-key field round-trips through
   `Office.context.document.settings`, `ANAKIN.SEARCH`/`ANAKIN.SCRAPE`/
   `ANAKIN.AGENTICSEARCH` resolve in cells (including cancellation via
   Excel's own recalculation/cell-edit triggers), and the task pane's
   `Excel.run` calls write to the correct cell range.
4. Once verified against a real key end-to-end, submit through **Partner
   Center > Microsoft Marketplace** (referenced from
   learn.microsoft.com/partner-center/marketplace-offers/submit-to-appsource-via-partner-center)
   — a manual listing process (icon, screenshots, category, privacy/terms
   URLs), not a CLI step.
