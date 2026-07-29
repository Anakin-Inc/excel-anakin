// Copyright Anakin (https://anakin.io).
//
// Task pane UI: enter an API key once, then scrape/search/agentic-search
// interactively and write results directly into the selected cell(s) via
// Excel.run(async (context) => { ... }) — the standard Excel JavaScript API
// pattern (context.workbook / context.sync()).

/* global Office, Excel, AnakinState, Anakin */

var elements = {};

function setStatus(el, message, kind) {
  el.textContent = message || "";
  el.className = "status" + (kind ? " " + kind : "");
}

function setBusy(button, busy, busyLabel, idleLabel) {
  button.disabled = busy;
  button.textContent = busy ? busyLabel : idleLabel;
}

function writeScalarToSelectedCell(value) {
  return Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load("rowCount,columnCount");
    return context.sync().then(function () {
      var target = range.getResizedRange(1 - range.rowCount, 1 - range.columnCount);
      target.values = [[value]];
      return context.sync();
    });
  });
}

/** Parse a textarea's JSON content; empty input yields {}. Throws on bad JSON. */
function parseJsonField(text, fieldLabel) {
  var trimmed = (text || "").trim();
  if (!trimmed) return {};
  try {
    return JSON.parse(trimmed);
  } catch (e) {
    throw new Error(fieldLabel + " must be valid JSON: " + e.message);
  }
}

function writeMatrixToSelectedCell(rows) {
  if (!rows.length) {
    return writeScalarToSelectedCell("No results.");
  }
  var colCount = rows[0].length;
  return Excel.run(function (context) {
    var anchor = context.workbook.getSelectedRange();
    anchor.load("rowCount,columnCount");
    return context.sync().then(function () {
      var target = anchor.getResizedRange(rows.length - anchor.rowCount, colCount - anchor.columnCount);
      target.values = rows;
      return context.sync();
    });
  });
}

function initApiKeyField() {
  Office.context.document.settings.refreshAsync(function () {
    var key = AnakinState.loadFromDocumentSettings();
    if (key) {
      elements.apiKeyInput.value = key;
      setStatus(elements.keyStatus, "API key loaded from this workbook.", "success");
    }
  });
}

function onSaveKey() {
  var value = elements.apiKeyInput.value.trim();
  AnakinState.setApiKey(value);
  setStatus(elements.keyStatus, value ? "API key saved." : "API key cleared.", "success");
}

function onScrape() {
  var url = elements.scrapeUrlInput.value.trim();
  if (!url) {
    setStatus(elements.resultStatus, "Enter a URL to scrape.", "error");
    return;
  }
  setBusy(elements.scrapeButton, true, "Scraping…", "Scrape into selected cell");
  setStatus(elements.resultStatus, "Scraping " + url + " …");
  Anakin.scrapeUrl({ url: url, useBrowser: elements.scrapeBrowserCheckbox.checked })
    .then(function (doc) {
      return writeScalarToSelectedCell(doc.markdown || doc.summary || "");
    })
    .then(function () {
      setStatus(elements.resultStatus, "Done — markdown written to the selected cell.", "success");
    })
    .catch(function (error) {
      setStatus(elements.resultStatus, error.message || String(error), "error");
    })
    .then(function () {
      setBusy(elements.scrapeButton, false, "Scraping…", "Scrape into selected cell");
    });
}

function onSearch() {
  var query = elements.searchQueryInput.value.trim();
  if (!query) {
    setStatus(elements.resultStatus, "Enter a search query.", "error");
    return;
  }
  var limit = parseInt(elements.searchLimitInput.value, 10) || 5;
  setBusy(elements.searchButton, true, "Searching…", "Search into selected cell");
  setStatus(elements.resultStatus, "Searching…");
  Anakin.search({ prompt: query, limit: limit })
    .then(function (result) {
      var rows = (result.results || []).map(function (item) {
        return [item.url || "", item.title || "", item.snippet || ""];
      });
      return writeMatrixToSelectedCell(rows.length ? rows : [["", "", "No results."]]);
    })
    .then(function () {
      setStatus(elements.resultStatus, "Done — results written starting at the selected cell.", "success");
    })
    .catch(function (error) {
      setStatus(elements.resultStatus, error.message || String(error), "error");
    })
    .then(function () {
      setBusy(elements.searchButton, false, "Searching…", "Search into selected cell");
    });
}

function onAgenticSearch() {
  var prompt = elements.agenticPromptInput.value.trim();
  if (!prompt) {
    setStatus(elements.resultStatus, "Enter a research question.", "error");
    return;
  }
  setBusy(elements.agenticButton, true, "Researching…", "Run into selected cell");
  setStatus(elements.resultStatus, "Running agentic research — this can take a while…");
  Anakin.agenticSearch({ prompt: prompt })
    .then(function (job) {
      var summary = (job.generatedJson && job.generatedJson.summary) || "";
      var structured = job.generatedJson && job.generatedJson.structured_data;
      if (structured) {
        return writeMatrixToSelectedCell([[summary, JSON.stringify(structured)]]);
      }
      return writeScalarToSelectedCell(summary);
    })
    .then(function () {
      setStatus(elements.resultStatus, "Done — summary written to the selected cell.", "success");
    })
    .catch(function (error) {
      setStatus(elements.resultStatus, error.message || String(error), "error");
    })
    .then(function () {
      setBusy(elements.agenticButton, false, "Researching…", "Run into selected cell");
    });
}

// ─── State-changing actions: task-pane buttons only, never Custom Functions.
// Each carries `destructiveHint: true` in Anakin's own MCP tool definitions
// (anakin-mcp/src/tools/*.ts) because it mutates account state or a
// third-party site — a live cell formula could silently re-run it on
// Excel's own recalculation. Every handler below confirms with the user
// before calling the API. See README.md's "Read vs. write" section.

function onWireWriteAction() {
  var actionId = elements.wireWriteActionIdInput.value.trim();
  if (!actionId) {
    setStatus(elements.resultStatus, "Enter an action_id.", "error");
    return;
  }
  var params;
  try {
    params = parseJsonField(elements.wireWriteParamsInput.value, "Params");
  } catch (error) {
    setStatus(elements.resultStatus, error.message, "error");
    return;
  }
  if (!window.confirm('Run Wire write action "' + actionId + '"? This changes state on the target site.')) {
    return;
  }
  var credentialId = elements.wireWriteCredentialIdInput.value.trim();
  setBusy(elements.wireWriteButton, true, "Running…", "Run write action");
  setStatus(elements.resultStatus, "Running Wire write action " + actionId + " …");
  var options = { actionId: actionId, params: params };
  if (credentialId) options.credentialId = credentialId;
  Anakin.wireRun(options)
    .then(function (job) {
      return writeScalarToSelectedCell(JSON.stringify(job.data !== undefined ? job.data : job));
    })
    .then(function () {
      setStatus(elements.resultStatus, "Done — result written to the selected cell.", "success");
    })
    .catch(function (error) {
      setStatus(elements.resultStatus, error.message || String(error), "error");
    })
    .then(function () {
      setBusy(elements.wireWriteButton, false, "Running…", "Run write action");
    });
}

function onWireLogin() {
  var catalogSlug = elements.wireLoginCatalogSlugInput.value.trim();
  if (!catalogSlug) {
    setStatus(elements.resultStatus, "Enter a catalog_slug.", "error");
    return;
  }
  var params;
  try {
    params = parseJsonField(elements.wireLoginParamsInput.value, "Login params");
  } catch (error) {
    setStatus(elements.resultStatus, error.message, "error");
    return;
  }
  if (!window.confirm('Sign in to "' + catalogSlug + '"? This stores an encrypted session.')) {
    return;
  }
  var identityName = elements.wireLoginIdentityNameInput.value.trim();
  setBusy(elements.wireLoginButton, true, "Signing in…", "Sign in");
  setStatus(elements.resultStatus, "Signing in to " + catalogSlug + " …");
  var options = { catalogSlug: catalogSlug };
  if (Object.keys(params).length) options.params = params;
  if (identityName) options.identityName = identityName;
  Anakin.wireLogin(options)
    .then(function (result) {
      return writeScalarToSelectedCell(JSON.stringify(result));
    })
    .then(function () {
      setStatus(elements.resultStatus, "Done — result written to the selected cell.", "success");
    })
    .catch(function (error) {
      setStatus(elements.resultStatus, error.message || String(error), "error");
    })
    .then(function () {
      setBusy(elements.wireLoginButton, false, "Signing in…", "Sign in");
    });
}

function onWireBuild() {
  var websiteUrl = elements.wireBuildUrlInput.value.trim();
  var goal = elements.wireBuildGoalInput.value.trim();
  if (!websiteUrl || !goal) {
    setStatus(elements.resultStatus, "Enter a website_url and a goal.", "error");
    return;
  }
  if (!window.confirm('Request a new Wire action for "' + websiteUrl + '"? This spends credits.')) {
    return;
  }
  setBusy(elements.wireBuildButton, true, "Requesting…", "Request build");
  setStatus(elements.resultStatus, "Requesting a new Wire action for " + websiteUrl + " …");
  Anakin.wireBuild({
    websiteUrl: websiteUrl,
    goal: goal,
    visibility: elements.wireBuildVisibilitySelect.value,
    force: elements.wireBuildForceCheckbox.checked,
  })
    .then(function (result) {
      return writeScalarToSelectedCell(JSON.stringify(result));
    })
    .then(function () {
      setStatus(elements.resultStatus, "Done — result written to the selected cell.", "success");
    })
    .catch(function (error) {
      setStatus(elements.resultStatus, error.message || String(error), "error");
    })
    .then(function () {
      setBusy(elements.wireBuildButton, false, "Requesting…", "Request build");
    });
}

function onMonitorCreate() {
  var url = elements.monitorCreateUrlInput.value.trim();
  var intervalMinutes = parseInt(elements.monitorCreateIntervalInput.value, 10);
  if (!url || !intervalMinutes) {
    setStatus(elements.resultStatus, "Enter a URL and a check interval (minutes).", "error");
    return;
  }
  var advanced;
  try {
    advanced = parseJsonField(elements.monitorCreateAdvancedInput.value, "Advanced options");
  } catch (error) {
    setStatus(elements.resultStatus, error.message, "error");
    return;
  }
  if (!window.confirm("Create a monitor checking this URL every " + intervalMinutes + " minutes? This is a recurring, credit-billed job.")) {
    return;
  }
  setBusy(elements.monitorCreateButton, true, "Creating…", "Create monitor");
  setStatus(elements.resultStatus, "Creating monitor…");
  var options = {
    url: url,
    intervalMinutes: intervalMinutes,
    scope: elements.monitorCreateScopeSelect.value,
    watchMode: elements.monitorCreateWatchmodeSelect.value,
    aiMode: elements.monitorCreateAimodeCheckbox.checked,
  };
  Object.keys(advanced).forEach(function (key) {
    options[key] = advanced[key];
  });
  Anakin.monitorCreate(options)
    .then(function (result) {
      return writeScalarToSelectedCell(JSON.stringify(result));
    })
    .then(function () {
      setStatus(elements.resultStatus, "Done — result written to the selected cell.", "success");
    })
    .catch(function (error) {
      setStatus(elements.resultStatus, error.message || String(error), "error");
    })
    .then(function () {
      setBusy(elements.monitorCreateButton, false, "Creating…", "Create monitor");
    });
}

function onMonitorControl() {
  var id = elements.monitorControlIdInput.value.trim();
  var action = elements.monitorControlActionSelect.value;
  if (!id) {
    setStatus(elements.resultStatus, "Enter a monitor ID.", "error");
    return;
  }
  var confirmMessage =
    action === "delete"
      ? 'Permanently delete monitor "' + id + '" and its history? This cannot be undone.'
      : 'Apply "' + action + '" to monitor "' + id + '"?';
  if (!window.confirm(confirmMessage)) {
    return;
  }
  setBusy(elements.monitorControlButton, true, "Applying…", "Apply");
  setStatus(elements.resultStatus, "Applying " + action + " to monitor " + id + " …");
  Anakin.monitorControl(id, action)
    .then(function (result) {
      return writeScalarToSelectedCell(JSON.stringify(result));
    })
    .then(function () {
      setStatus(elements.resultStatus, "Done — result written to the selected cell.", "success");
    })
    .catch(function (error) {
      setStatus(elements.resultStatus, error.message || String(error), "error");
    })
    .then(function () {
      setBusy(elements.monitorControlButton, false, "Applying…", "Apply");
    });
}

function onSessionDelete() {
  var id = elements.sessionDeleteIdInput.value.trim();
  if (!id) {
    setStatus(elements.resultStatus, "Enter a session ID.", "error");
    return;
  }
  if (!window.confirm('Permanently delete session "' + id + '"? Anything using it loses authenticated access.')) {
    return;
  }
  setBusy(elements.sessionDeleteButton, true, "Deleting…", "Delete session");
  setStatus(elements.resultStatus, "Deleting session " + id + " …");
  Anakin.sessionDelete(id)
    .then(function (result) {
      return writeScalarToSelectedCell(JSON.stringify(result));
    })
    .then(function () {
      setStatus(elements.resultStatus, "Done — session deleted.", "success");
    })
    .catch(function (error) {
      setStatus(elements.resultStatus, error.message || String(error), "error");
    })
    .then(function () {
      setBusy(elements.sessionDeleteButton, false, "Deleting…", "Delete session");
    });
}

function onBrowserTask() {
  var prompt = elements.browserTaskPromptInput.value.trim();
  if (!prompt) {
    setStatus(elements.resultStatus, "Enter a task prompt.", "error");
    return;
  }
  if (!window.confirm("Run this browser task? It drives a real browser and can change state on the target site, and may take up to ~5 minutes.")) {
    return;
  }
  var url = elements.browserTaskUrlInput.value.trim();
  var sessionId = elements.browserTaskSessionIdInput.value.trim();
  var maxSteps = parseInt(elements.browserTaskMaxStepsInput.value, 10);
  setBusy(elements.browserTaskButton, true, "Running…", "Run task into selected cell");
  setStatus(elements.resultStatus, "Running browser task — this can take a while…");
  var options = { prompt: prompt };
  if (url) options.url = url;
  if (sessionId) options.sessionId = sessionId;
  if (maxSteps) options.maxSteps = maxSteps;
  Anakin.browserTask(options)
    .then(function (result) {
      return writeScalarToSelectedCell(JSON.stringify(result));
    })
    .then(function () {
      setStatus(elements.resultStatus, "Done — result written to the selected cell.", "success");
    })
    .catch(function (error) {
      setStatus(elements.resultStatus, error.message || String(error), "error");
    })
    .then(function () {
      setBusy(elements.browserTaskButton, false, "Running…", "Run task into selected cell");
    });
}

Office.onReady(function (info) {
  if (info.host !== Office.HostType.Excel) {
    return;
  }

  elements = {
    apiKeyInput: document.getElementById("api-key-input"),
    saveKeyButton: document.getElementById("save-key-button"),
    keyStatus: document.getElementById("key-status"),
    scrapeUrlInput: document.getElementById("scrape-url-input"),
    scrapeBrowserCheckbox: document.getElementById("scrape-browser-checkbox"),
    scrapeButton: document.getElementById("scrape-button"),
    searchQueryInput: document.getElementById("search-query-input"),
    searchLimitInput: document.getElementById("search-limit-input"),
    searchButton: document.getElementById("search-button"),
    agenticPromptInput: document.getElementById("agentic-prompt-input"),
    agenticButton: document.getElementById("agentic-button"),

    wireWriteActionIdInput: document.getElementById("wire-write-action-id-input"),
    wireWriteParamsInput: document.getElementById("wire-write-params-input"),
    wireWriteCredentialIdInput: document.getElementById("wire-write-credential-id-input"),
    wireWriteButton: document.getElementById("wire-write-button"),

    wireLoginCatalogSlugInput: document.getElementById("wire-login-catalog-slug-input"),
    wireLoginParamsInput: document.getElementById("wire-login-params-input"),
    wireLoginIdentityNameInput: document.getElementById("wire-login-identity-name-input"),
    wireLoginButton: document.getElementById("wire-login-button"),

    wireBuildUrlInput: document.getElementById("wire-build-url-input"),
    wireBuildGoalInput: document.getElementById("wire-build-goal-input"),
    wireBuildVisibilitySelect: document.getElementById("wire-build-visibility-select"),
    wireBuildForceCheckbox: document.getElementById("wire-build-force-checkbox"),
    wireBuildButton: document.getElementById("wire-build-button"),

    monitorCreateUrlInput: document.getElementById("monitor-create-url-input"),
    monitorCreateIntervalInput: document.getElementById("monitor-create-interval-input"),
    monitorCreateScopeSelect: document.getElementById("monitor-create-scope-select"),
    monitorCreateWatchmodeSelect: document.getElementById("monitor-create-watchmode-select"),
    monitorCreateAimodeCheckbox: document.getElementById("monitor-create-aimode-checkbox"),
    monitorCreateAdvancedInput: document.getElementById("monitor-create-advanced-input"),
    monitorCreateButton: document.getElementById("monitor-create-button"),

    monitorControlIdInput: document.getElementById("monitor-control-id-input"),
    monitorControlActionSelect: document.getElementById("monitor-control-action-select"),
    monitorControlButton: document.getElementById("monitor-control-button"),

    sessionDeleteIdInput: document.getElementById("session-delete-id-input"),
    sessionDeleteButton: document.getElementById("session-delete-button"),

    browserTaskPromptInput: document.getElementById("browser-task-prompt-input"),
    browserTaskUrlInput: document.getElementById("browser-task-url-input"),
    browserTaskSessionIdInput: document.getElementById("browser-task-session-id-input"),
    browserTaskMaxStepsInput: document.getElementById("browser-task-max-steps-input"),
    browserTaskButton: document.getElementById("browser-task-button"),

    resultStatus: document.getElementById("result-status"),
  };

  initApiKeyField();

  elements.saveKeyButton.addEventListener("click", onSaveKey);
  elements.scrapeButton.addEventListener("click", onScrape);
  elements.searchButton.addEventListener("click", onSearch);
  elements.agenticButton.addEventListener("click", onAgenticSearch);

  elements.wireWriteButton.addEventListener("click", onWireWriteAction);
  elements.wireLoginButton.addEventListener("click", onWireLogin);
  elements.wireBuildButton.addEventListener("click", onWireBuild);
  elements.monitorCreateButton.addEventListener("click", onMonitorCreate);
  elements.monitorControlButton.addEventListener("click", onMonitorControl);
  elements.sessionDeleteButton.addEventListener("click", onSessionDelete);
  elements.browserTaskButton.addEventListener("click", onBrowserTask);
});
