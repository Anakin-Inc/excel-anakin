// Copyright Anakin (https://anakin.io).
//
// Shared global state for the SharedRuntime. taskpane.html loads this file,
// api.js, functions.js, and taskpane.js as plain <script> tags in that order
// (see taskpane.html) — under a SharedRuntime, all of them execute in the
// SAME window/JS context, so a single global object here is visible from
// both the task pane UI code (taskpane.js) and the custom functions
// (functions.js) without any message-passing. This is the same mechanism
// the real reference sample (OfficeDev/Office-Add-in-samples,
// excel-shared-runtime-scenario/src/utilities/globals.js) uses to share a
// "isConnected" flag between its task pane and its custom functions.
//
// Anakin's API key is entered once in the task pane and needs to be
// readable from ANAKIN.SCRAPE / ANAKIN.SEARCH / ANAKIN.AGENTICSEARCH when
// they run from a cell formula — this is why this add-in requires a shared
// runtime rather than the (simpler, default) isolated per-surface runtime.

/* global Office */

var AnakinState = (function () {
  var STORAGE_KEY = "anakin.apiKey";
  var apiKey = null;

  // Office.context.document.settings persists inside the workbook itself
  // (roams with the file, survives reloads) — the documented mechanism for
  // add-in state that needs to outlive a single session. localStorage is
  // used only as a same-session fallback before settings have loaded.
  try {
    apiKey = window.localStorage.getItem(STORAGE_KEY) || null;
  } catch (e) {
    // localStorage can throw in some sandboxed webviews; state.js must not
    // fail to load because of it.
    apiKey = null;
  }

  function getApiKey() {
    return apiKey;
  }

  function setApiKey(value) {
    apiKey = value || null;
    try {
      if (apiKey) {
        window.localStorage.setItem(STORAGE_KEY, apiKey);
      } else {
        window.localStorage.removeItem(STORAGE_KEY);
      }
    } catch (e) {
      // Ignore — localStorage is a convenience cache, not the source of
      // truth (document settings, written by taskpane.js, is).
    }
    if (typeof Office !== "undefined" && Office.context && Office.context.document) {
      Office.context.document.settings.set(STORAGE_KEY, apiKey);
      Office.context.document.settings.saveAsync();
    }
  }

  function loadFromDocumentSettings() {
    if (typeof Office !== "undefined" && Office.context && Office.context.document) {
      var fromDoc = Office.context.document.settings.get(STORAGE_KEY);
      if (fromDoc) {
        apiKey = fromDoc;
      }
    }
    return apiKey;
  }

  return {
    getApiKey: getApiKey,
    setApiKey: setApiKey,
    loadFromDocumentSettings: loadFromDocumentSettings,
  };
})();
