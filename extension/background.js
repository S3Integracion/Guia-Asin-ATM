const DEFAULT_SETTINGS = {
  language: "es",
  region: "mx",
  excelDoc1Url: "",
  excelDoc2Url: "",
  bolsas: [],
  sheetPrimary: "",
  sheetDuplicates: "Duplicados",
  assistedMode: true,
  validations: {
    dedupeGuide: true,
    dedupeAsin: true,
    accumulateQuantity: true,
    splitDuplicatesToSheet: true
  },
  logs: [],
  history: []
};

const mergeDefaults = (stored) => {
  if (!stored || typeof stored !== "object") {
    return { ...DEFAULT_SETTINGS };
  }
  return {
    ...DEFAULT_SETTINGS,
    ...stored,
    validations: {
      ...DEFAULT_SETTINGS.validations,
      ...(stored.validations || {})
    },
    logs: Array.isArray(stored.logs) ? stored.logs : [],
    history: Array.isArray(stored.history) ? stored.history : []
  };
};

chrome.runtime.onInstalled.addListener(async () => {
  const stored = await chrome.storage.local.get();
  const merged = mergeDefaults(stored);
  await chrome.storage.local.set(merged);
});

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (!message || typeof message !== "object") {
    return;
  }

  if (message.type === "LOG_EVENT") {
    chrome.storage.local.get({ logs: [] }).then(({ logs }) => {
      const next = [
        {
          id: crypto.randomUUID(),
          ts: Date.now(),
          level: message.level || "info",
          source: message.source || "popup",
          message: message.message || ""
        },
        ...logs
      ].slice(0, 200);
      chrome.storage.local.set({ logs: next });
    });
  }

  if (message.type === "CLEAR_LOGS") {
    chrome.storage.local.set({ logs: [] }).then(() => sendResponse({ ok: true }));
    return true;
  }

  if (message.type === "CLEAR_HISTORY") {
    chrome.storage.local.set({ history: [] }).then(() => sendResponse({ ok: true }));
    return true;
  }
});
