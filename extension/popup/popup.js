const guidesInput = document.getElementById("guidesInput");
const validCount = document.getElementById("validCount");
const duplicateCount = document.getElementById("duplicateCount");
const invalidCount = document.getElementById("invalidCount");
const statusText = document.getElementById("statusText");
const progressCount = document.getElementById("progressCount");
const progressTotal = document.getElementById("progressTotal");
const outputTsv = document.getElementById("outputTsv");
const manualTsv = document.getElementById("manualTsv");

const CACHE_KEY = "lookupCache";

let currentGuides = [];
let running = false;
let lastResults = [];

const SELLER_HOST_PATTERNS = [
  "sellercentral.amazon.com",
  "sellercentral.amazon.com.mx",
  "sellercentral.amazon.ca",
  "sellercentral.amazon.co.uk",
  "sellercentral.amazon.es",
  "sellercentral.amazon.de",
  "sellercentral.amazon.fr",
  "sellercentral.amazon.it",
  "sellercentral.amazon.co.jp",
  "sellercentral.amazon.com.au",
  "sellercentral.amazon.com.br",
  "sellercentral.amazon.in",
  "sellercentral.amazon.com.tr",
  "sellercentral.amazon.ae",
  "sellercentral.amazon.sa",
  "sellercentral.amazon.eg",
  "sellercentral.amazon.sg"
];

const setStatus = (message) => {
  statusText.textContent = message;
};

const saveCache = async () => {
  try {
    await chrome.storage.local.set({
      [CACHE_KEY]: {
        outputTsv: outputTsv.value || "",
        manualTsv: manualTsv.value || "",
        results: lastResults,
        updatedAt: Date.now()
      }
    });
  } catch (error) {
    // ignore
  }
};

const loadCache = async () => {
  try {
    const stored = await chrome.storage.local.get(CACHE_KEY);
    const cached = stored[CACHE_KEY];
    if (!cached) {
      return;
    }
    if (typeof cached.outputTsv === "string") {
      outputTsv.value = cached.outputTsv;
    }
    if (typeof cached.manualTsv === "string") {
      manualTsv.value = cached.manualTsv;
    }
    if (Array.isArray(cached.results)) {
      lastResults = cached.results;
    }
  } catch (error) {
    // ignore
  }
};

const normalizeGuideToken = (token) => {
  const trimmed = String(token || "").trim();
  if (!trimmed) {
    return "";
  }
  if (/^\d{3}-\d{7}-\d{7}$/.test(trimmed)) {
    return trimmed;
  }
  const digits = trimmed.replace(/[^\d]/g, "");
  if (digits.length >= 8 && digits.length <= 14) {
    return digits;
  }
  return "";
};

const parseGuides = (text) => {
  const rawTokens = String(text || "")
    .replace(/\r/g, "")
    .split(/[\s,;]+/)
    .filter(Boolean);
  const guides = [];
  const invalid = [];
  const seen = new Set();
  let duplicates = 0;

  rawTokens.forEach((token) => {
    const normalized = normalizeGuideToken(token);
    if (!normalized) {
      invalid.push(token);
      return;
    }
    if (seen.has(normalized)) {
      duplicates += 1;
      return;
    }
    seen.add(normalized);
    guides.push(normalized);
  });

  return {
    guides,
    invalidCount: invalid.length,
    duplicateCount: duplicates
  };
};

const updateGuideStats = () => {
  const parsed = parseGuides(guidesInput.value || "");
  currentGuides = parsed.guides;
  validCount.textContent = String(parsed.guides.length);
  invalidCount.textContent = String(parsed.invalidCount);
  duplicateCount.textContent = String(parsed.duplicateCount);
  progressCount.textContent = "0";
  progressTotal.textContent = String(parsed.guides.length);
};

const buildOutput = (results = []) => {
  lastResults = results;
  const mainRows = [];
  const manualRows = [];

  results.forEach((result) => {
    if (!result || !result.guide) {
      return;
    }
    if (result.status === "error") {
      manualRows.push({
        guide: result.guide,
        asin: "",
        quantity: "",
        reason: result.note || "Error"
      });
      return;
    }
    const items = Array.isArray(result.items) ? result.items : [];
    if (!items.length || result.status === "not_found") {
      manualRows.push({
        guide: result.guide,
        asin: "",
        quantity: "",
        reason: "No encontrado"
      });
      return;
    }

    if (items.length > 1) {
      items.forEach((item) => {
        manualRows.push({
          guide: result.guide,
          asin: item.asin || "",
          quantity: item.quantity || "",
          reason: "Varios ASIN"
        });
      });
      return;
    }

    const item = items[0];
    if (!item || !item.asin) {
      manualRows.push({
        guide: result.guide,
        asin: "",
        quantity: "",
        reason: "ASIN no detectado"
      });
      return;
    }

    if (!item.quantity) {
      manualRows.push({
        guide: result.guide,
        asin: item.asin,
        quantity: "",
        reason: "Cantidad no detectada"
      });
      return;
    }

    mainRows.push([
      item.asin,
      item.quantity,
      item.tracking || result.tracking || result.guide
    ]);
  });

  outputTsv.value = mainRows.map((row) => row.join("\t")).join("\n");
  manualTsv.value = manualRows
    .map((row) => [row.guide, row.asin, row.quantity, row.reason].join("\t"))
    .join("\n");
  saveCache();
};

const getActiveTab = async () => {
  const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
  return tabs[0] || null;
};

const ensureSellerTab = async () => {
  const tab = await getActiveTab();
  if (!tab || !tab.id) {
    setStatus("No se detecto una pestana activa.");
    return null;
  }
  const url = tab.url || "";
  const host = (() => {
    try {
      return new URL(url).hostname;
    } catch (error) {
      return "";
    }
  })();
  if (!SELLER_HOST_PATTERNS.includes(host)) {
    setStatus("Abre la pestana de Amazon Seller antes de iniciar.");
    return null;
  }
  return tab;
};

const sendToSeller = async (message) => {
  const tab = await ensureSellerTab();
  if (!tab) {
    return null;
  }
  try {
    await chrome.scripting.executeScript({
      target: { tabId: tab.id, allFrames: true },
      files: ["content/amazon_orders.js"]
    });
  } catch (error) {
    // ignore
  }
  try {
    return await chrome.tabs.sendMessage(tab.id, message);
  } catch (error) {
    setStatus("No se pudo comunicar con Amazon Seller.");
    return null;
  }
};

document.getElementById("clearGuides").addEventListener("click", () => {
  guidesInput.value = "";
  updateGuideStats();
});

document.getElementById("copyGuides").addEventListener("click", async () => {
  if (!currentGuides.length) {
    setStatus("No hay guias para copiar.");
    return;
  }
  await navigator.clipboard.writeText(currentGuides.join("\n"));
  setStatus("Guias copiadas.");
});

document.getElementById("startLookup").addEventListener("click", async () => {
  if (running) {
    return;
  }
  updateGuideStats();
  if (!currentGuides.length) {
    setStatus("Pega guias validas antes de iniciar.");
    return;
  }
  running = true;
  outputTsv.value = "";
  manualTsv.value = "";
  progressCount.textContent = "0";
  progressTotal.textContent = String(currentGuides.length);
  setStatus("Iniciando busqueda...");
  await chrome.storage.local.set({ cancelLookup: false });
  const response = await sendToSeller({
    type: "START_LOOKUP",
    guides: currentGuides
  });
  if (!response || !response.ok) {
    running = false;
    setStatus(response?.error || "No se pudo iniciar la busqueda.");
  }
});

document.getElementById("stopLookup").addEventListener("click", async () => {
  if (!running) {
    setStatus("No hay busqueda activa.");
    return;
  }
  await chrome.storage.local.set({ cancelLookup: true });
  await sendToSeller({ type: "CANCEL_LOOKUP" });
  setStatus("Cancelando...");
});

document.getElementById("copyOutput").addEventListener("click", async () => {
  if (!outputTsv.value) {
    setStatus("No hay resultados para copiar.");
    return;
  }
  await navigator.clipboard.writeText(outputTsv.value);
  setStatus("Resultado copiado.");
});

document.getElementById("clearOutput").addEventListener("click", () => {
  outputTsv.value = "";
  saveCache();
});

document.getElementById("copyManual").addEventListener("click", async () => {
  if (!manualTsv.value) {
    setStatus("No hay revision manual.");
    return;
  }
  await navigator.clipboard.writeText(manualTsv.value);
  setStatus("Revision copiada.");
});

document.getElementById("clearManual").addEventListener("click", () => {
  manualTsv.value = "";
  saveCache();
});

document.getElementById("clearCache").addEventListener("click", async () => {
  await chrome.storage.local.clear();
  outputTsv.value = "";
  manualTsv.value = "";
  lastResults = [];
  setStatus("Cache limpiada.");
});

let inputTimer = null;
guidesInput.addEventListener("input", () => {
  if (inputTimer) {
    clearTimeout(inputTimer);
  }
  inputTimer = setTimeout(() => {
    updateGuideStats();
  }, 180);
});

chrome.runtime.onMessage.addListener((message) => {
  if (!message || typeof message !== "object") {
    return;
  }
  if (message.type === "LOOKUP_PROGRESS") {
    const payload = message.payload || {};
    progressCount.textContent = String(payload.index || 0);
    progressTotal.textContent = String(payload.total || currentGuides.length);
    if (payload.guide) {
      setStatus(`Buscando ${payload.guide} (${payload.index}/${payload.total})`);
    }
    return;
  }
  if (message.type === "LOOKUP_COMPLETE") {
    running = false;
    progressCount.textContent = String(message.payload?.total || currentGuides.length);
    buildOutput(message.payload?.results || []);
    setStatus(message.payload?.canceled ? "Busqueda cancelada." : "Busqueda finalizada.");
    return;
  }
  if (message.type === "LOOKUP_ERROR") {
    running = false;
    setStatus(message.payload?.error || "Error en la busqueda.");
  }
});

document.addEventListener("DOMContentLoaded", () => {
  updateGuideStats();
  loadCache();
});
