// Copyright Anakin (https://anakin.io).
//
// Thin Anakin API client shared by the task pane (taskpane.js) and the
// custom functions (functions.js) — loaded before both, see taskpane.html.
//
// Ground truth for the endpoints, request bodies, and response shapes below
// was read directly from anakin-py's src/anakin/client.py and
// src/anakin/models.py (the reference SDK for this API), not guessed:
//   - POST /url-scraper (submit) + GET /url-scraper/:jobId (poll)
//   - POST /search (synchronous, no polling)
//   - POST /agentic-search (submit) + GET /agentic-search/:jobId (poll)
// Base URL and header confirmed live: a CORS preflight against
// https://api.anakin.io/v1/search returned "access-control-allow-origin: *",
// so these fetch() calls work directly from a browser/webview context with
// no proxy — see SUBMIT.md for the exact request/response.

/* global AnakinState */

var Anakin = (function () {
  var BASE_URL = "https://api.anakin.io/v1";
  var POLL_INTERVAL_MS = 1500;
  var POLL_TIMEOUT_MS = 120000; // 2 minutes — url-scraper/agentic-search jobs
  // are not instant; this is a generous but bounded ceiling so a bad job
  // (or a bad key) fails loudly in a cell instead of spinning forever.

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

  return {
    scrapeUrl: scrapeUrl,
    search: search,
    agenticSearch: agenticSearch,
  };
})();
