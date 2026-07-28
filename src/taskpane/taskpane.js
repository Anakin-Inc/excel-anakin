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
    resultStatus: document.getElementById("result-status"),
  };

  initApiKeyField();

  elements.saveKeyButton.addEventListener("click", onSaveKey);
  elements.scrapeButton.addEventListener("click", onScrape);
  elements.searchButton.addEventListener("click", onSearch);
  elements.agenticButton.addEventListener("click", onAgenticSearch);
});
