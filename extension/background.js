const DEFAULT_SETTINGS = {
  language: "es",
  region: "mx",
  excelDoc1Url: "",
  excelDoc2Url: "",
  bolsas: [],
  doc1Selection: {
    mode: "infer",
    range: "",
    column: "",
    manualColumn: "",
    selectionType: "",
    bagLabel: "",
    guides: [],
    invalidCount: 0,
    duplicateCount: 0,
    capturedAt: 0
  },
  doc2Selection: {
    startCell: "",
    startColumn: "",
    startRow: 1,
    manualColumn: "",
    approved: false,
    headersEnsured: false,
    nextIndex: 1,
    capturedAt: 0
  },
  dedupe: {
    guides: [],
    asinsByGuide: {},
    doc2Guides: [],
    doc2AsinsByGuide: {}
  },
  sheetPrimary: "",
  sheetDuplicates: "Duplicados",
  sheetDoc1: "",
  msAuth: {
    clientId: "",
    tenant: "common",
    accessToken: "",
    refreshToken: "",
    expiresAt: 0,
    account: ""
  },
  assistedMode: true,
  validations: {
    dedupeGuide: true,
    dedupeAsin: true,
    accumulateQuantity: true,
    splitDuplicatesToSheet: true
  },
  stats: {
    lastRun: "",
    lastError: ""
  },
  pendingOutput: {
    mainRows: [],
    manualRows: [],
    nextIndex: 1,
    createdAt: 0
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
    doc1Selection: {
      ...DEFAULT_SETTINGS.doc1Selection,
      ...(stored.doc1Selection || {})
    },
    doc2Selection: {
      ...DEFAULT_SETTINGS.doc2Selection,
      ...(stored.doc2Selection || {})
    },
    dedupe: {
      ...DEFAULT_SETTINGS.dedupe,
      ...(stored.dedupe || {})
    },
    msAuth: {
      ...DEFAULT_SETTINGS.msAuth,
      ...(stored.msAuth || {})
    },
    validations: {
      ...DEFAULT_SETTINGS.validations,
      ...(stored.validations || {})
    },
    stats: {
      ...DEFAULT_SETTINGS.stats,
      ...(stored.stats || {})
    },
    pendingOutput: {
      ...DEFAULT_SETTINGS.pendingOutput,
      ...(stored.pendingOutput || {})
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
