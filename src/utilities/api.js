// Copyright Anakin (https://anakin.io).
//
// Thin Anakin API client shared by the task pane (taskpane.js) and the
// custom functions (functions.js) — loaded before both, see taskpane.html.
//
// Ground truth for the endpoints, request bodies, and response shapes below
// was read directly from anakin-py's src/anakin/client.py and
// src/anakin/models.py (the reference SDK for this API) for scrape/search/
// agentic-search, and from anakin-mcp's src/client.ts + src/tools/*.ts (the
// same REST API, wrapped for Anakin's own MCP server) for every endpoint
// added after that — map, crawl, wire/*, ai-visibility/*, sessions,
// monitors, and the browser-automation endpoint:
//   - POST /url-scraper (submit) + GET /url-scraper/:jobId (poll)
//   - POST /search (synchronous, no polling)
//   - POST /agentic-search (submit) + GET /agentic-search/:jobId (poll)
//   - POST /map (submit) + GET /map/:jobId (poll)
//   - POST /crawl (submit) + GET /crawl/:jobId (poll)
//   - GET /wire/resolve?q=&limit= (synchronous)
//   - GET /wire/catalog or /wire/catalog/:slug (synchronous)
//   - POST /wire/task {action_id, params?, credential_id?, identity_id?} —
//     either returns data inline (sync actions) or a job_id polled at
//     GET /wire/jobs/:jobId (paced by the server's retry_after_ms hint)
//   - POST /wire/login, POST /wire/build-request (synchronous)
//   - GET /ai-visibility/sources (synchronous)
//   - POST /ai-visibility/search (submit) + GET /ai-visibility/search/:id
//     (poll until status !== "running" — a "failed" result is still
//     returned, not thrown, since it carries per-source data)
//   - GET /sessions?domain=, DELETE /sessions/:id
//   - GET /monitors, GET /monitors/:id, GET /monitors/:id/changes,
//     POST /monitors, POST /monitors/:id/{pause,resume,run}, DELETE /monitors/:id
//   - POST /ai/evaluate {async:true} (submit) + GET /ai/jobs/:id (poll) —
//     browser_task; the server hard-caps a run at ~330s, so the poll window
//     here is set well past that (see BROWSER_TASK_POLL_TIMEOUT_MS)
// Base URL and header confirmed live: a CORS preflight against
// https://api.anakin.io/v1/search returned "access-control-allow-origin: *",
// so these fetch() calls work directly from a browser/webview context with
// no proxy — see SUBMIT.md for the exact request/response.

/* global AnakinState */

var Anakin = (function () {
  var BASE_URL = "https://api.anakin.io/v1";
  var POLL_INTERVAL_MS = 1500;
  var POLL_TIMEOUT_MS = 120000; // 2 minutes — url-scraper/agentic-search/map/
  // crawl/wire/ai-visibility jobs are not instant; this is a generous but
  // bounded ceiling so a bad job (or a bad key) fails loudly in a cell
  // instead of spinning forever.
  var BROWSER_TASK_POLL_TIMEOUT_MS = 360000; // 6 minutes — browser_task is
  // hard-capped server-side at ~330s (see anakin-mcp's src/client.ts), so
  // this poll window has to outlast the longest legitimate run.

  function sleep(ms, signal) {
    return new Promise(function (resolve, reject) {
      var timer = setTimeout(resolve, ms);
      if (signal) {
        signal.addEventListener("abort", function () {
          clearTimeout(timer);
          reject(new Error("Cancelled"));
        });
      }
    });
  }

  function apiKeyOrThrow() {
    var key = AnakinState.getApiKey();
    if (!key) {
      throw new Error(
        "No Anakin API key set. Open the Anakin task pane and enter your API " +
          "key (get one free at anakin.io/dashboard) before using ANAKIN.* functions."
      );
    }
    return key;
  }

  function request(method, path, body, signal) {
    var key = apiKeyOrThrow();
    return fetch(BASE_URL + path, {
      method: method,
      headers: {
        "Content-Type": "application/json",
        "X-API-Key": key,
      },
      body: body ? JSON.stringify(body) : undefined,
      signal: signal,
    }).then(function (response) {
      return response.json().then(function (json) {
        if (!response.ok) {
          var message =
            (json && (json.message || json.error)) ||
            "Anakin API request failed with status " + response.status;
          throw new Error(message);
        }
        return json;
      });
    });
  }

  function isTerminal(job) {
    return job.status === "completed" || job.status === "failed";
  }

  function pollJob(resourcePath, jobId, signal) {
    var deadline = Date.now() + POLL_TIMEOUT_MS;
    function poll() {
      return request("GET", resourcePath + "/" + jobId, null, signal).then(function (job) {
        if (isTerminal(job)) {
          return job;
        }
        if (Date.now() >= deadline) {
          throw new Error(
            "Anakin job " + jobId + " is still \"" + (job.status || "processing") +
              "\" after " + Math.round(POLL_TIMEOUT_MS / 1000) + "s. The job keeps " +
              "running on Anakin's side — try again in a moment."
          );
        }
        return sleep(POLL_INTERVAL_MS, signal).then(poll);
      });
    }
    return poll();
  }

  function assertSucceeded(job, verb) {
    if (job.status === "failed") {
      throw new Error("Anakin " + verb + " job failed: " + (job.error || "unknown error"));
    }
    return job;
  }

  // ─── scrape (POST /url-scraper, poll GET /url-scraper/:jobId) ───────────
  function scrapeUrl(options, signal) {
    var body = {
      url: options.url,
      country: options.country || "us",
      useBrowser: !!options.useBrowser,
      generateJson: !!options.generateJson,
      forceFresh: !!options.forceFresh,
    };
    if (options.sessionId) body.sessionId = options.sessionId;
    if (options.sessionName) body.sessionName = options.sessionName;

    return request("POST", "/url-scraper", body, signal)
      .then(function (submitted) {
        var jobId = submitted.jobId || submitted.job_id || submitted.id;
        if (!jobId) {
          throw new Error("Anakin's scrape submission response did not include a jobId.");
        }
        return pollJob("/url-scraper", jobId, signal);
      })
      .then(function (job) {
        return assertSucceeded(job, "scrape");
      });
  }

  // ─── search (POST /search, synchronous) ──────────────────────────────────
  function search(options, signal) {
    var body = { prompt: options.prompt, limit: options.limit || 5 };
    return request("POST", "/search", body, signal);
  }

  // ─── agentic search (POST /agentic-search, poll GET /agentic-search/:jobId) ─
  function agenticSearch(options, signal) {
    var body = {
      prompt: options.prompt,
      useBrowser: options.useBrowser !== false, // defaults to true
    };
    if (options.schema) body.schema = options.schema;

    return request("POST", "/agentic-search", body, signal)
      .then(function (submitted) {
        var jobId = submitted.jobId || submitted.job_id || submitted.id;
        if (!jobId) {
          throw new Error("Anakin's agentic-search submission response did not include a jobId.");
        }
        return pollJob("/agentic-search", jobId, signal);
      })
      .then(function (job) {
        return assertSucceeded(job, "agentic-search");
      });
  }

  // ─── map (POST /map, poll GET /map/:jobId) ───────────────────────────────
  function map(options, signal) {
    var body = {
      url: options.url,
      limit: options.limit || 100,
      depth: options.depth || 2,
      limitPerLevel: options.limitPerLevel || 100,
      includeSubdomains: !!options.includeSubdomains,
      includeExternalLinks: !!options.includeExternalLinks,
      useBrowser: !!options.useBrowser,
    };
    if (options.search) body.search = options.search;

    return request("POST", "/map", body, signal)
      .then(function (submitted) {
        var jobId = submitted.jobId || submitted.job_id || submitted.id;
        if (!jobId) {
          throw new Error("Anakin's map submission response did not include a jobId.");
        }
        return pollJob("/map", jobId, signal);
      })
      .then(function (job) {
        return assertSucceeded(job, "map");
      });
  }

  // ─── crawl (POST /crawl, poll GET /crawl/:jobId) ─────────────────────────
  function crawl(options, signal) {
    var body = {
      url: options.url,
      maxPages: options.maxPages || 10,
      depth: options.depth || 1,
      country: options.country || "us",
      useBrowser: !!options.useBrowser,
    };
    if (options.includePatterns) body.includePatterns = options.includePatterns;
    if (options.excludePatterns) body.excludePatterns = options.excludePatterns;
    if (options.sessionId) body.sessionId = options.sessionId;
    if (options.sessionName) body.sessionName = options.sessionName;

    return request("POST", "/crawl", body, signal)
      .then(function (submitted) {
        var jobId = submitted.jobId || submitted.job_id || submitted.id;
        if (!jobId) {
          throw new Error("Anakin's crawl submission response did not include a jobId.");
        }
        return pollJob("/crawl", jobId, signal);
      })
      .then(function (job) {
        return assertSucceeded(job, "crawl");
      });
  }

  // ─── wire: discovery (GET, synchronous) ──────────────────────────────────
  function wireDiscover(options, signal) {
    var qs = "q=" + encodeURIComponent(options.q);
    if (options.limit) qs += "&limit=" + encodeURIComponent(options.limit);
    return request("GET", "/wire/resolve?" + qs, null, signal);
  }

  function wireCatalog(options, signal) {
    var slug = options && options.slug;
    var path = slug ? "/wire/catalog/" + encodeURIComponent(slug) : "/wire/catalog";
    return request("GET", path, null, signal);
  }

  // ─── wire: execution (POST /wire/task, poll GET /wire/jobs/:jobId) ──────
  // Mirrors anakin-mcp's client.ts pollWireJob: paced by the server's
  // retry_after_ms hint (clamped 500ms–10s) instead of a fixed interval.
  function pollWireJob(jobId, signal) {
    var deadline = Date.now() + POLL_TIMEOUT_MS;
    var path = "/wire/jobs/" + encodeURIComponent(jobId);
    function poll() {
      return request("GET", path, null, signal).then(function (job) {
        if (job.status === "completed") return job;
        if (job.status === "failed") {
          var reason = (job.error && (job.error.message || job.error.code)) || "unknown error";
          throw new Error("Wire job failed: " + reason);
        }
        if (Date.now() >= deadline) {
          throw new Error(
            "Wire job " + jobId + " is still \"" + (job.status || "processing") +
              "\" after " + Math.round(POLL_TIMEOUT_MS / 1000) + "s. The job keeps " +
              "running on Anakin's side — try again in a moment."
          );
        }
        var wait = typeof job.retry_after_ms === "number" ? job.retry_after_ms : POLL_INTERVAL_MS;
        wait = Math.min(Math.max(wait, 500), 10000);
        return sleep(wait, signal).then(poll);
      });
    }
    return poll();
  }

  // Runs both wire_read_action and wire_write_action — the read/write split
  // is which caller invokes this (functions.js for read, taskpane.js for
  // write), not a different request shape; the API distinguishes them by
  // the action's own type, not by a body field.
  function wireRun(options, signal) {
    var body = { action_id: options.actionId };
    if (options.params && Object.keys(options.params).length > 0) {
      body.params = options.params;
    }
    if (options.credentialId) body.credential_id = options.credentialId;
    if (options.identityId) body.identity_id = options.identityId;

    return request("POST", "/wire/task", body, signal).then(function (accepted) {
      if (!accepted.job_id) {
        return accepted; // Sync action — terminal data came back inline.
      }
      return pollWireJob(accepted.job_id, signal);
    });
  }

  function wireLogin(options, signal) {
    var body = { catalog_slug: options.catalogSlug };
    if (options.params) body.params = options.params;
    if (options.identityName) body.identity_name = options.identityName;
    if (options.sourceId) body.source_id = options.sourceId;
    if (options.sourceRef) body.source_ref = options.sourceRef;
    return request("POST", "/wire/login", body, signal);
  }

  function wireBuild(options, signal) {
    var body = { website_url: options.websiteUrl, goal: options.goal };
    if (options.catalogId) body.catalog_id = options.catalogId;
    if (options.visibility) body.visibility = options.visibility;
    if (options.force !== undefined) body.force = !!options.force;
    return request("POST", "/wire/build-request", body, signal);
  }

  // ─── AI Visibility ────────────────────────────────────────────────────────
  function aiVisibilitySources(signal) {
    return request("GET", "/ai-visibility/sources", null, signal);
  }

  function aiVisibilitySearch(options, signal) {
    var body = { query: options.query };
    if (options.sources) body.sources = options.sources;
    if (options.country) body.country = options.country;

    return request("POST", "/ai-visibility/search", body, signal).then(function (submitted) {
      var searchId = submitted.search_id;
      if (!searchId) return submitted;
      var deadline = Date.now() + POLL_TIMEOUT_MS;
      var path = "/ai-visibility/search/" + encodeURIComponent(searchId);
      function poll() {
        return request("GET", path, null, signal).then(function (search) {
          // Unlike pollJob, a "failed" status is returned (not thrown): the
          // payload still carries per-source results/errors the caller needs.
          if (search.status !== "running") return search;
          if (Date.now() >= deadline) {
            throw new Error(
              "AI visibility search " + searchId + " timed out after " +
                Math.round(POLL_TIMEOUT_MS / 1000) + "s; poll it later via the " +
                "dashboard or retry."
            );
          }
          return sleep(POLL_INTERVAL_MS, signal).then(poll);
        });
      }
      return poll();
    });
  }

  // ─── browser sessions ─────────────────────────────────────────────────────
  function sessionsList(domain, signal) {
    var path = "/sessions" + (domain ? "?domain=" + encodeURIComponent(domain) : "");
    return request("GET", path, null, signal);
  }

  function sessionDelete(id, signal) {
    return request("DELETE", "/sessions/" + encodeURIComponent(id), null, signal);
  }

  // ─── website monitoring ───────────────────────────────────────────────────
  function monitorList(id, signal) {
    var path = id ? "/monitors/" + encodeURIComponent(id) : "/monitors";
    return request("GET", path, null, signal);
  }

  function monitorChanges(id, signal) {
    return request("GET", "/monitors/" + encodeURIComponent(id) + "/changes", null, signal);
  }

  function monitorCreate(options, signal) {
    var body = { url: options.url, intervalMinutes: options.intervalMinutes };
    // Only send fields the caller set — the API applies its own defaults,
    // mirroring anakin-mcp's client.ts monitorCreate.
    var passthrough = [
      "scope", "watchMode", "watchFormat", "outputSchema", "aiMode", "aiGoal",
      "useBrowser", "country", "sessionId", "isActive", "expiresAt",
      "alertWebhookUrl", "alertEmails", "maxPages", "maxDepth",
      "includePatterns", "excludePatterns", "wireActionId", "wireCatalogSlug",
      "wireCredentialId", "wireParams", "wireWatchPaths",
    ];
    passthrough.forEach(function (key) {
      if (options[key] !== undefined) body[key] = options[key];
    });
    return request("POST", "/monitors", body, signal);
  }

  function monitorControl(id, action, signal) {
    var base = "/monitors/" + encodeURIComponent(id);
    if (action === "pause") return request("POST", base + "/pause", null, signal);
    if (action === "resume") return request("POST", base + "/resume", null, signal);
    if (action === "run_now") return request("POST", base + "/run", null, signal);
    if (action === "delete") return request("DELETE", base, null, signal);
    return Promise.reject(
      new Error("Unknown monitor action \"" + action + "\" — use pause, resume, run_now, or delete.")
    );
  }

  // ─── AI browser automation (POST /ai/evaluate async, poll GET /ai/jobs/:id) ─
  function browserTask(options, signal) {
    var body = { prompt: options.prompt, async: true };
    if (options.url) body.url = options.url;
    if (options.sessionId) body.session_id = options.sessionId;
    if (options.maxSteps) body.max_steps = options.maxSteps;
    if (options.timeoutMs) body.timeout_ms = options.timeoutMs;
    if (options.outputSchema) body.output_schema = options.outputSchema;

    return request("POST", "/ai/evaluate", body, signal).then(function (accepted) {
      if (!accepted.workflow_id) {
        return accepted; // Service answered synchronously (shouldn't happen with async:true).
      }
      var deadline = Date.now() + BROWSER_TASK_POLL_TIMEOUT_MS;
      var path = "/ai/jobs/" + encodeURIComponent(accepted.workflow_id);
      function poll() {
        return request("GET", path, null, signal).then(function (job) {
          if (job.status === "completed") return job.result || job;
          if (job.status === "failed" || job.status === "timed_out") {
            throw new Error("Browser task " + job.status + ": " + (job.error || "unknown error"));
          }
          if (Date.now() >= deadline) {
            throw new Error(
              "Browser task is still running after " +
                Math.round(BROWSER_TASK_POLL_TIMEOUT_MS / 60000) + " minutes of polling."
            );
          }
          return sleep(POLL_INTERVAL_MS, signal).then(poll);
        });
      }
      return poll();
    });
  }

  return {
    scrapeUrl: scrapeUrl,
    search: search,
    agenticSearch: agenticSearch,
    map: map,
    crawl: crawl,
    wireDiscover: wireDiscover,
    wireCatalog: wireCatalog,
    wireRun: wireRun,
    wireLogin: wireLogin,
    wireBuild: wireBuild,
    aiVisibilitySources: aiVisibilitySources,
    aiVisibilitySearch: aiVisibilitySearch,
    sessionsList: sessionsList,
    sessionDelete: sessionDelete,
    monitorList: monitorList,
    monitorChanges: monitorChanges,
    monitorCreate: monitorCreate,
    monitorControl: monitorControl,
    browserTask: browserTask,
  };
})();
