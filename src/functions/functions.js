// Copyright Anakin (https://anakin.io).
//
// Excel Custom Functions: =ANAKIN.SCRAPE(url), =ANAKIN.SEARCH(query),
// =ANAKIN.AGENTICSEARCH(prompt), and the read-only/idempotent functions
// below (MAP, CRAWL, WIREDISCOVER, WIRECATALOG, WIREREADACTION,
// AIVISIBILITYSEARCH, AIVISIBILITYSOURCES, SESSIONLIST, MONITORLIST,
// MONITORCHANGES) — callable directly from any cell.
//
// Why only these are custom functions, not the other 7 of Anakin's 21 REST
// endpoints (wire_write_action, wire_login, wire_build, monitor_create,
// monitor_control, session_delete, browser_task): every one of those is
// state-changing on Anakin's own account state or on a third-party site —
// Excel recalculates a cell formula automatically and silently (on open, on
// a referenced cell changing, on manual recalc), which would replay a write
// with no human in the loop each time. Anakin's own MCP tool definitions
// (anakin-mcp/src/tools/*.ts) mark every one of those seven
// `destructiveHint: true` for exactly this reason. Custom functions have no
// built-in confirmation step (see learn.microsoft.com/office/dev/add-ins/
// excel/custom-functions-overview and .../custom-functions-web-reqs — read
// live: functions "return a Promise"/stream, nothing about a user-approval
// gate), so those seven are task-pane buttons instead (see taskpane.js),
// where a click is an explicit, one-shot, confirmable action. See
// README.md's "Read vs. write" section for the full explanation.
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
// while a job is still polling — an AbortController tied to the fetch()
// calls in api.js makes that actually stop the in-flight request instead of
// leaving it running invisibly.

/* global CustomFunctions, Anakin */

/**
 * Best-effort unwrap of a list-shaped API response whose item fields aren't
 * documented (session_list, monitor_list, monitor_changes) — tries common
 * REST list-wrapper keys, then falls back to treating a single object as a
 * one-item list. Never guesses at item-level field names; each row is
 * returned as its own JSON string so nothing about the item shape is
 * fabricated.
 * @param {*} response
 * @returns {Array<Object>}
 */
function asItemList(response) {
  if (Array.isArray(response)) return response;
  if (response && typeof response === "object") {
    var candidateKeys = ["sessions", "monitors", "changes", "results", "data", "items"];
    for (var i = 0; i < candidateKeys.length; i++) {
      if (Array.isArray(response[candidateKeys[i]])) return response[candidateKeys[i]];
    }
    return [response];
  }
  return [];
}

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

/**
 * Discover all reachable URLs under a site. Submits the map job and polls
 * until it completes or fails.
 * @customfunction MAP
 * @param {string} url The starting URL for discovery (typically a homepage or section root).
 * @param {number} [limit] Maximum number of URLs to return overall. Defaults to 100.
 * @param {number} [depth] How many link-hops from the starting URL to follow. Defaults to 2.
 * @param {boolean} [useBrowser] Render with a headless browser — best for single-page apps. Defaults to FALSE.
 * @param {CustomFunctions.CancelableInvocation} invocation Invocation object (for cancellation).
 * @returns {Promise<string[][]>} A dynamic array with one discovered URL per row.
 */
function mapUrls(url, limit, depth, useBrowser, invocation) {
  var controller = new AbortController();
  invocation.onCanceled = function () {
    controller.abort();
  };
  var options = { url: url };
  if (limit !== undefined && limit !== null) options.limit = limit;
  if (depth !== undefined && depth !== null) options.depth = depth;
  if (useBrowser !== undefined && useBrowser !== null) options.useBrowser = !!useBrowser;
  return Anakin.map(options, controller.signal).then(function (result) {
    var links = result.links || [];
    if (!links.length) return [["No URLs found."]];
    return links.map(function (link) {
      return [link];
    });
  });
}

/**
 * Bulk-fetch markdown across a site. Submits the crawl job and polls until
 * it completes or fails.
 * @customfunction CRAWL
 * @param {string} url The starting URL.
 * @param {number} [maxPages] Hard cap on pages fetched. Defaults to 10.
 * @param {number} [depth] Link-hops from the starting URL to follow. Defaults to 1.
 * @param {boolean} [useBrowser] Render each page in a headless browser — best for single-page apps. Defaults to FALSE.
 * @param {CustomFunctions.CancelableInvocation} invocation Invocation object (for cancellation).
 * @returns {Promise<string[][]>} A dynamic array of [url, status, markdown] rows, one per crawled page.
 */
function crawlSite(url, maxPages, depth, useBrowser, invocation) {
  var controller = new AbortController();
  invocation.onCanceled = function () {
    controller.abort();
  };
  var options = { url: url };
  if (maxPages !== undefined && maxPages !== null) options.maxPages = maxPages;
  if (depth !== undefined && depth !== null) options.depth = depth;
  if (useBrowser !== undefined && useBrowser !== null) options.useBrowser = !!useBrowser;
  return Anakin.crawl(options, controller.signal).then(function (result) {
    var pages = result.pages || [];
    if (!pages.length) return [["", "", "No pages crawled."]];
    return pages.map(function (page) {
      return [page.url || "", page.status || "", page.markdown || page.error || ""];
    });
  });
}

/**
 * Find Wire actions for a task from a natural-language intent. Wire is
 * Anakin's catalog of pre-built automation actions across hundreds of
 * websites. Synchronous — no polling.
 * @customfunction WIREDISCOVER
 * @param {string} query The intent in natural language, e.g. "top phones on walmart".
 * @param {number} [limit] Maximum number of candidate actions to return. Defaults to 5.
 * @param {CustomFunctions.CancelableInvocation} invocation Invocation object (for cancellation).
 * @returns {Promise<string[][]>} A dynamic array of [action_id, type, details (JSON)] rows, one per candidate action. Run a "read" action with WIREREADACTION.
 */
function wireDiscover(query, limit, invocation) {
  var controller = new AbortController();
  invocation.onCanceled = function () {
    controller.abort();
  };
  var options = { q: query };
  if (limit !== undefined && limit !== null) options.limit = limit;
  return Anakin.wireDiscover(options, controller.signal).then(function (result) {
    var results = (result && result.results) || [];
    if (!results.length) return [["", "", "No matching Wire actions."]];
    return results.map(function (item) {
      return [item.action_id || "", item.type || "", JSON.stringify(item)];
    });
  });
}

/**
 * Browse the Wire catalog. With no slug, lists every supported website and
 * its action count. Pass a slug (e.g. "walmart") for that site's full action
 * list with parameter schemas. Synchronous — no polling. The response shape
 * differs by call (all-catalogs list vs. one catalog's action list), so this
 * returns the raw JSON rather than guessing a fixed column layout.
 * @customfunction WIRECATALOG
 * @param {string} [slug] Catalog slug to inspect (e.g. "walmart"). Omit to list all catalogs.
 * @param {CustomFunctions.CancelableInvocation} invocation Invocation object (for cancellation).
 * @returns {Promise<string>} The catalog response as a JSON string.
 */
function wireCatalog(slug, invocation) {
  var controller = new AbortController();
  invocation.onCanceled = function () {
    controller.abort();
  };
  var options = {};
  if (slug) options.slug = slug;
  return Anakin.wireCatalog(options, controller.signal).then(function (result) {
    return JSON.stringify(result);
  });
}

/**
 * Run a Wire READ action — one whose type is "read" (extracts data, no
 * state change on the target site). Find action_ids and their param schemas
 * with WIREDISCOVER / WIRECATALOG first. Transparently polls the job to
 * completion. For WRITE actions (type "write"), use the task pane's "Run a
 * Wire write action" button instead — see README.md.
 * @customfunction WIREREADACTION
 * @param {string} actionId The action to run (from WIREDISCOVER / WIRECATALOG).
 * @param {string} [paramsJson] The action's input parameters as a JSON object string, e.g. {"query":"iphone 15"}. Omit for actions that take none.
 * @param {string} [credentialId] Required when the action's auth_mode is "required"; honored when "optional". Get one from the Anakin dashboard.
 * @param {CustomFunctions.CancelableInvocation} invocation Invocation object (for cancellation).
 * @returns {Promise<string[][]>} A dynamic array of [field, value] rows from the action's extracted data.
 */
function wireReadAction(actionId, paramsJson, credentialId, invocation) {
  var controller = new AbortController();
  invocation.onCanceled = function () {
    controller.abort();
  };
  var params = {};
  if (paramsJson) {
    try {
      params = JSON.parse(paramsJson);
    } catch (e) {
      return Promise.reject(new Error("paramsJson is not valid JSON: " + e.message));
    }
  }
  var options = { actionId: actionId, params: params };
  if (credentialId) options.credentialId = credentialId;
  return Anakin.wireRun(options, controller.signal).then(function (job) {
    var data = job.data || {};
    var keys = Object.keys(data);
    if (!keys.length) return [["(no data)", ""]];
    return keys.map(function (key) {
      var value = data[key];
      return [key, value !== null && typeof value === "object" ? JSON.stringify(value) : value];
    });
  });
}

/**
 * Ask multiple AI answer engines (ChatGPT, Gemini, Google AI Overview) the
 * same question and compare their answers. Submits the search and polls
 * until it's no longer "running". Billed per source at that source's rate.
 * @customfunction AIVISIBILITYSEARCH
 * @param {string} query The question to ask every engine (max 2000 characters).
 * @param {string} [sources] Comma-separated engine slugs to query (see AIVISIBILITYSOURCES). Omit to query all enabled engines.
 * @param {string} [country] Two-letter ISO country for the search geography. Defaults to "us".
 * @param {CustomFunctions.CancelableInvocation} invocation Invocation object (for cancellation).
 * @returns {Promise<string[][]>} A dynamic array with one [result label, JSON] row per engine, plus a final "synthesis" row.
 */
function aiVisibilitySearch(query, sources, country, invocation) {
  var controller = new AbortController();
  invocation.onCanceled = function () {
    controller.abort();
  };
  var options = { query: query };
  if (sources) {
    options.sources = sources
      .split(",")
      .map(function (s) {
        return s.trim();
      })
      .filter(Boolean);
  }
  if (country) options.country = country;
  return Anakin.aiVisibilitySearch(options, controller.signal).then(function (search) {
    var results = search.results || [];
    var rows = results.map(function (r, i) {
      return ["result_" + (i + 1), JSON.stringify(r)];
    });
    rows.push(["synthesis", search.synthesis || ""]);
    return rows;
  });
}

/**
 * List the AI answer engines available to AIVISIBILITYSEARCH. Synchronous —
 * no polling.
 * @customfunction AIVISIBILITYSOURCES
 * @param {CustomFunctions.CancelableInvocation} invocation Invocation object (for cancellation).
 * @returns {Promise<string[][]>} A dynamic array of [slug, label] rows, one per available engine.
 */
function aiVisibilitySources(invocation) {
  var controller = new AbortController();
  invocation.onCanceled = function () {
    controller.abort();
  };
  return Anakin.aiVisibilitySources(controller.signal).then(function (result) {
    var list = asItemList(result);
    if (!list.length) return [["", "No engines available."]];
    return list.map(function (item) {
      return [item.slug || "", item.label || ""];
    });
  });
}

/**
 * List your saved browser sessions — encrypted login states used by SCRAPE/
 * CRAWL (sessionId), monitors, and the task pane's browser-task button.
 * Synchronous — no polling.
 * @customfunction SESSIONLIST
 * @param {string} [domain] Filter to sessions for one website domain, e.g. "amazon.com".
 * @param {CustomFunctions.CancelableInvocation} invocation Invocation object (for cancellation).
 * @returns {Promise<string[][]>} A dynamic array with one JSON-encoded session per row.
 */
function sessionList(domain, invocation) {
  var controller = new AbortController();
  invocation.onCanceled = function () {
    controller.abort();
  };
  return Anakin.sessionsList(domain, controller.signal).then(function (result) {
    var list = asItemList(result);
    if (!list.length) return [["No saved sessions."]];
    return list.map(function (item) {
      return [JSON.stringify(item)];
    });
  });
}

/**
 * List your website monitors, or pass an id to fetch one monitor's full
 * configuration. Synchronous — no polling.
 * @customfunction MONITORLIST
 * @param {string} [id] Monitor ID — fetch just this monitor instead of the full list.
 * @param {CustomFunctions.CancelableInvocation} invocation Invocation object (for cancellation).
 * @returns {Promise<string[][]>} A dynamic array with one JSON-encoded monitor per row.
 */
function monitorList(id, invocation) {
  var controller = new AbortController();
  invocation.onCanceled = function () {
    controller.abort();
  };
  return Anakin.monitorList(id, controller.signal).then(function (result) {
    var list = asItemList(result);
    if (!list.length) return [["No monitors found."]];
    return list.map(function (item) {
      return [JSON.stringify(item)];
    });
  });
}

/**
 * Get the detected changes for a monitor. Find the monitor id with
 * MONITORLIST first. Synchronous — no polling.
 * @customfunction MONITORCHANGES
 * @param {string} id The monitor ID (from MONITORLIST).
 * @param {CustomFunctions.CancelableInvocation} invocation Invocation object (for cancellation).
 * @returns {Promise<string[][]>} A dynamic array with one JSON-encoded change record per row.
 */
function monitorChanges(id, invocation) {
  var controller = new AbortController();
  invocation.onCanceled = function () {
    controller.abort();
  };
  return Anakin.monitorChanges(id, controller.signal).then(function (result) {
    var list = asItemList(result);
    if (!list.length) return [["No changes recorded."]];
    return list.map(function (item) {
      return [JSON.stringify(item)];
    });
  });
}

CustomFunctions.associate("SCRAPE", scrape);
CustomFunctions.associate("SEARCH", search);
CustomFunctions.associate("AGENTICSEARCH", agenticSearch);
CustomFunctions.associate("MAP", mapUrls);
CustomFunctions.associate("CRAWL", crawlSite);
CustomFunctions.associate("WIREDISCOVER", wireDiscover);
CustomFunctions.associate("WIRECATALOG", wireCatalog);
CustomFunctions.associate("WIREREADACTION", wireReadAction);
CustomFunctions.associate("AIVISIBILITYSEARCH", aiVisibilitySearch);
CustomFunctions.associate("AIVISIBILITYSOURCES", aiVisibilitySources);
CustomFunctions.associate("SESSIONLIST", sessionList);
CustomFunctions.associate("MONITORLIST", monitorList);
CustomFunctions.associate("MONITORCHANGES", monitorChanges);
