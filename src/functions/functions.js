// Copyright Anakin (https://anakin.io).
//
// Excel Custom Functions: =ANAKIN.SCRAPE(url), =ANAKIN.SEARCH(query),
// =ANAKIN.AGENTICSEARCH(prompt) — callable directly from any cell.
//
// Loaded by taskpane.html (the SharedRuntime page — see manifest.xml's
// CustomFunctions ExtensionPoint, whose Script/Page both point at
// Taskpane.Url) after state.js and api.js, so `AnakinState` and `Anakin`
// below are the same globals the task pane UI (taskpane.js) reads and
// writes — that's how the API key entered once in the task pane becomes
// visible to formulas in cells.
//
// Each function is @cancelable: Excel calls the assigned `oncanceled`
// callback if the user deletes the formula or edits the cell mid-flight
// while a scrape/search job is still polling — an AbortController tied to
// the fetch() calls in api.js makes that actually stop the in-flight
// request instead of leaving it running invisibly.

/* global CustomFunctions, Anakin */

/**
 * Scrape a URL with Anakin and return clean markdown. Submits the scrape
 * job and polls until it completes or fails.
 * @customfunction SCRAPE
 * @param {string} url The full URL of the page to scrape.
 * @param {boolean} [useBrowser] Use a headless browser — best for JS-heavy sites. Defaults to FALSE.
 * @param {CustomFunctions.CancelableInvocation} invocation Invocation object (for cancellation).
 * @returns {Promise<string>} The scraped page as markdown.
 */
function scrape(url, useBrowser, invocation) {
  var controller = new AbortController();
  invocation.onCanceled = function () {
    controller.abort();
  };
  return Anakin.scrapeUrl({ url: url, useBrowser: !!useBrowser }, controller.signal).then(
    function (doc) {
      return doc.markdown || doc.summary || "";
    }
  );
}

/**
 * AI-powered web search with citations. Synchronous on Anakin's side (no
 * polling). Costs 3 credits per call.
 * @customfunction SEARCH
 * @param {string} query The search query or question.
 * @param {number} [limit] Maximum number of results to return. Defaults to 5.
 * @param {CustomFunctions.CancelableInvocation} invocation Invocation object (for cancellation).
 * @returns {Promise<string[][]>} A dynamic array of [url, title, snippet] rows.
 */
function search(query, limit, invocation) {
  var controller = new AbortController();
  invocation.onCanceled = function () {
    controller.abort();
  };
  return Anakin.search({ prompt: query, limit: limit || 5 }, controller.signal).then(function (
    result
  ) {
    var rows = (result.results || []).map(function (item) {
      return [item.url || "", item.title || "", item.snippet || ""];
    });
    return rows.length ? rows : [["", "", "No results."]];
  });
}

/**
 * Run Anakin's multi-stage AI research pipeline (search, scrape citations,
 * extract structured data) and return the AI-generated summary. Submits the
 * job and polls until it completes or fails. Costs 10 credits.
 * @customfunction AGENTICSEARCH
 * @param {string} prompt The research question or topic.
 * @param {boolean} [useBrowser] Use a headless browser when scraping citation pages. Defaults to TRUE.
 * @param {CustomFunctions.CancelableInvocation} invocation Invocation object (for cancellation).
 * @returns {Promise<string>} The AI-generated research summary.
 */
function agenticSearch(prompt, useBrowser, invocation) {
  var controller = new AbortController();
  invocation.onCanceled = function () {
    controller.abort();
  };
  var options = { prompt: prompt };
  if (useBrowser !== undefined && useBrowser !== null) {
    options.useBrowser = !!useBrowser;
  }
  return Anakin.agenticSearch(options, controller.signal).then(function (job) {
    return (job.generatedJson && job.generatedJson.summary) || "";
  });
}

CustomFunctions.associate("SCRAPE", scrape);
CustomFunctions.associate("SEARCH", search);
CustomFunctions.associate("AGENTICSEARCH", agenticSearch);
