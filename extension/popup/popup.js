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

const SELLER_URL_PATTERNS = [
  "*://sellercentral.amazon.com/*",
  "*://sellercentral.amazon.com.mx/*",
  "*://sellercentral.amazon.ca/*",
  "*://sellercentral.amazon.co.uk/*",
  "*://sellercentral.amazon.es/*",
  "*://sellercentral.amazon.de/*",
  "*://sellercentral.amazon.fr/*",
  "*://sellercentral.amazon.it/*",
  "*://sellercentral.amazon.co.jp/*",
  "*://sellercentral.amazon.com.au/*",
  "*://sellercentral.amazon.com.br/*",
  "*://sellercentral.amazon.in/*",
  "*://sellercentral.amazon.com.tr/*",
  "*://sellercentral.amazon.ae/*",
  "*://sellercentral.amazon.sa/*",
  "*://sellercentral.amazon.eg/*",
  "*://sellercentral.amazon.sg/*"
];

const STRINGS = {
  es: {
    appTitle: "Guia ASIN ATM",
    appSubtitle: "Sincroniza guias con Amazon Seller y Excel Online.",
    statusLabel: "Estado:",
    statusReady: "Listo",
    statusPending: "Pendiente",
    statusNeedsTabs: "Faltan pestanas",
    languageLabel: "Idioma",
    regionLabel: "Region Amazon",
    tabConfig: "Configuracion",
    tabValidation: "Validacion",
    tabExecution: "Ejecucion",
    tabLogs: "Logs",
    excelLinksTitle: "Archivos Excel",
    excelDoc1Label: "Documento 1 (solo lectura)",
    excelDoc1Placeholder: "Link del inventario",
    excelDoc2Label: "Documento 2 (salida)",
    excelDoc2Placeholder: "Link de resultados",
    sheetPrimaryLabel: "Hoja principal",
    sheetDupLabel: "Hoja duplicados",
    saveConfigButton: "Guardar configuracion",
    doc1SelectionTitle: "Seleccion Documento 1",
    doc1SelectionHint: "Selecciona la columna o bolsa en Excel y presiona Capturar.",
    doc1ModeLabel: "Modo de seleccion",
    doc1ModeInfer: "Inferir automaticamente",
    doc1ModeBolsa: "Bolsa",
    doc1ModeColumna: "Columna",
    captureDoc1Button: "Capturar seleccion",
    doc1RangeLabel: "Rango detectado",
    doc1TypeLabel: "Tipo",
    doc1GuidesLabel: "Guias listas",
    doc1InvalidLabel: "Guias invalidas",
    doc2SelectionTitle: "Seleccion Documento 2",
    doc2SelectionHint: "Selecciona la celda inicial (encabezado) en Excel y captura la salida.",
    captureDoc2Button: "Capturar salida",
    approveDoc2Button: "Aprobar seleccion",
    doc2StartLabel: "Celda inicial",
    doc2ManualLabel: "Columna revision",
    doc2ApprovalLabel: "Aprobacion:",
    approvalPending: "Pendiente",
    approvalReady: "Aprobado",
    bolsasTitle: "Columnas de Bolsas",
    bolsasHint: "Agrega encabezados o nombres de columna con numeros de guia.",
    bolsaPlaceholder: "Bolsa A",
    addBolsaButton: "Agregar",
    validationRulesTitle: "Reglas persistentes",
    dedupeGuideLabel: "No repetir numero de guia",
    dedupeAsinLabel: "No repetir ASIN",
    accumulateQtyLabel: "Acumular cantidad si ASIN ya existe",
    splitDuplicatesLabel: "Enviar duplicados a hoja separada",
    integrityTitle: "Estado de integridad",
    doc2ValidationTitle: "Validacion Documento 2",
    doc2ValidationHint: "Selecciona la tabla actual (D-G) y captura para validar duplicados.",
    captureDoc2ExistingButton: "Capturar datos actuales",
    uniqueGuidesLabel: "Guias unicas registradas",
    uniqueAsinsLabel: "ASIN unicos",
    lastRunLabel: "Ultima corrida",
    lastErrorLabel: "Ultimo error",
    executionModeTitle: "Modo de ejecucion",
    modeAssisted: "Asistido",
    modeDirect: "Directo",
    assistedTitle: "Flujo asistido",
    step1: "Verificar pestanas abiertas y sesion activa",
    step2: "Leer guias desde Documento 1",
    step3: "Buscar ASIN y cantidad en Amazon Seller",
    step4: "Consolidar y validar duplicados",
    step5: "Escribir resultados en Documento 2",
    startAssistedButton: "Iniciar flujo asistido",
    directTitle: "Panel directo",
    checkTabsButton: "Validar pestanas",
    scanBolsasButton: "Leer bolsas",
    scanAmazonButton: "Buscar en Amazon",
    writeExcelButton: "Escribir Excel",
    runAllButton: "Ejecutar todo",
    statusTitle: "Estado de sesion",
    sellerStatusLabel: "Amazon Seller abierto",
    doc1StatusLabel: "Documento 1 abierto",
    doc2StatusLabel: "Documento 2 abierto",
    refreshStatusButton: "Actualizar estado",
    logsTitle: "Logs recientes",
    clearLogsButton: "Limpiar logs",
    historyTitle: "Historial",
    clearHistoryButton: "Limpiar historial",
    logsEmpty: "Sin logs por ahora.",
    historyEmpty: "Sin historial por ahora."
  },
  en: {
    appTitle: "Guide ASIN ATM",
    appSubtitle: "Sync guides with Amazon Seller and Excel Online.",
    statusLabel: "Status:",
    statusReady: "Ready",
    statusPending: "Pending",
    statusNeedsTabs: "Tabs missing",
    languageLabel: "Language",
    regionLabel: "Amazon Region",
    tabConfig: "Setup",
    tabValidation: "Validation",
    tabExecution: "Run",
    tabLogs: "Logs",
    excelLinksTitle: "Excel files",
    excelDoc1Label: "Document 1 (read only)",
    excelDoc1Placeholder: "Inventory link",
    excelDoc2Label: "Document 2 (output)",
    excelDoc2Placeholder: "Results link",
    sheetPrimaryLabel: "Primary sheet",
    sheetDupLabel: "Duplicate sheet",
    saveConfigButton: "Save configuration",
    doc1SelectionTitle: "Document 1 selection",
    doc1SelectionHint: "Select the column or bag in Excel and press Capture.",
    doc1ModeLabel: "Selection mode",
    doc1ModeInfer: "Auto infer",
    doc1ModeBolsa: "Bag",
    doc1ModeColumna: "Column",
    captureDoc1Button: "Capture selection",
    doc1RangeLabel: "Detected range",
    doc1TypeLabel: "Type",
    doc1GuidesLabel: "Guides ready",
    doc1InvalidLabel: "Invalid guides",
    doc2SelectionTitle: "Document 2 selection",
    doc2SelectionHint: "Select the header start cell in Excel and capture the output.",
    captureDoc2Button: "Capture output",
    approveDoc2Button: "Approve selection",
    doc2StartLabel: "Start cell",
    doc2ManualLabel: "Manual column",
    doc2ApprovalLabel: "Approval:",
    approvalPending: "Pending",
    approvalReady: "Approved",
    bolsasTitle: "Bag columns",
    bolsasHint: "Add header names or column labels containing guide numbers.",
    bolsaPlaceholder: "Bag A",
    addBolsaButton: "Add",
    validationRulesTitle: "Persistent rules",
    dedupeGuideLabel: "Do not repeat guide number",
    dedupeAsinLabel: "Do not repeat ASIN",
    accumulateQtyLabel: "Increase quantity if ASIN exists",
    splitDuplicatesLabel: "Send duplicates to separate sheet",
    integrityTitle: "Integrity status",
    doc2ValidationTitle: "Document 2 validation",
    doc2ValidationHint: "Select the current table (D-G) and capture to validate duplicates.",
    captureDoc2ExistingButton: "Capture current data",
    uniqueGuidesLabel: "Unique guides stored",
    uniqueAsinsLabel: "Unique ASINs",
    lastRunLabel: "Last run",
    lastErrorLabel: "Last error",
    executionModeTitle: "Execution mode",
    modeAssisted: "Assisted",
    modeDirect: "Direct",
    assistedTitle: "Assisted flow",
    step1: "Check open tabs and active session",
    step2: "Read guides from Document 1",
    step3: "Find ASIN and quantity in Amazon Seller",
    step4: "Consolidate and validate duplicates",
    step5: "Write results in Document 2",
    startAssistedButton: "Start assisted flow",
    directTitle: "Direct panel",
    checkTabsButton: "Validate tabs",
    scanBolsasButton: "Read bags",
    scanAmazonButton: "Search Amazon",
    writeExcelButton: "Write Excel",
    runAllButton: "Run all",
    statusTitle: "Session status",
    sellerStatusLabel: "Amazon Seller open",
    doc1StatusLabel: "Document 1 open",
    doc2StatusLabel: "Document 2 open",
    refreshStatusButton: "Refresh status",
    logsTitle: "Recent logs",
    clearLogsButton: "Clear logs",
    historyTitle: "History",
    clearHistoryButton: "Clear history",
    logsEmpty: "No logs yet.",
    historyEmpty: "No history yet."
  }
};

let currentSettings = { ...DEFAULT_SETTINGS };

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
    stats: {
      ...DEFAULT_SETTINGS.stats,
      ...(stored.stats || {})
    },
    pendingOutput: {
      ...DEFAULT_SETTINGS.pendingOutput,
      ...(stored.pendingOutput || {})
    },
    validations: {
      ...DEFAULT_SETTINGS.validations,
      ...(stored.validations || {})
    },
    bolsas: Array.isArray(stored.bolsas) ? stored.bolsas : [],
    logs: Array.isArray(stored.logs) ? stored.logs : [],
    history: Array.isArray(stored.history) ? stored.history : []
  };
};

const getString = (key) => {
  const lang = currentSettings.language || "es";
  return (STRINGS[lang] && STRINGS[lang][key]) || STRINGS.es[key] || key;
};

const applyTranslations = () => {
  document.querySelectorAll("[data-i18n]").forEach((el) => {
    const key = el.getAttribute("data-i18n");
    el.textContent = getString(key);
  });
  document.querySelectorAll("[data-i18n-placeholder]").forEach((el) => {
    const key = el.getAttribute("data-i18n-placeholder");
    el.setAttribute("placeholder", getString(key));
  });
};

const logEvent = (level, message) => {
  chrome.runtime.sendMessage({
    type: "LOG_EVENT",
    level,
    source: "popup",
    message
  });
};

const setLastError = async (message) => {
  currentSettings.stats = {
    ...currentSettings.stats,
    lastError: message
  };
  await chrome.storage.local.set({ stats: currentSettings.stats });
  updateIntegrityStats();
};

const setStatusDot = (id, ready) => {
  const el = document.getElementById(id);
  if (!el) {
    return;
  }
  el.classList.toggle("ready", ready);
};

const normalizeUrl = (value) => {
  try {
    const url = new URL(value);
    return `${url.origin}${url.pathname}`.replace(/\/$/, "");
  } catch (error) {
    return "";
  }
};

const isSellerUrl = (value) => {
  try {
    const url = new URL(value);
    return url.hostname.startsWith("sellercentral.amazon.");
  } catch (error) {
    return false;
  }
};

const safeTabsQuery = async (queryInfo, fallbackFilter) => {
  try {
    return await chrome.tabs.query(queryInfo);
  } catch (error) {
    logEvent("error", `Filtro de tabs invalido: ${error.message || error}`);
    const tabs = await chrome.tabs.query({});
    return fallbackFilter ? tabs.filter(fallbackFilter) : tabs;
  }
};

const fallbackGetExcelSelection = async (tabId) => {
  try {
    const results = await chrome.scripting.executeScript({
      target: { tabId, allFrames: true },
      func: () => {
        const selectors = [
          'input[aria-label="Name box"]',
          'input[aria-label="Name Box"]',
          'input[aria-label*="Name"]',
          'input[aria-label*="Nombre"]',
          'input[title*="Name"]',
          'input[title*="Nombre"]',
          '[data-automation-id*="nameBox"] input',
          '[data-automation-id*="NameBox"] input',
          '#NameBox input',
          '#NameBox'
        ];
        const formulaSelectors = [
          'input[aria-label="Formula Bar"]',
          'input[aria-label="Barra de formulas"]',
          'input[aria-label*="Formula"]',
          'textarea[aria-label*="Formula"]',
          '[role="textbox"][aria-label*="Formula"]'
        ];
        const getValue = (el) =>
          el && (el.value || el.getAttribute("value") || el.textContent || "");
        const nameBox = selectors
          .map((selector) => document.querySelector(selector))
          .find(Boolean);
        const formulaBar = formulaSelectors
          .map((selector) => document.querySelector(selector))
          .find(Boolean);
        const range = String(getValue(nameBox)).trim();
        const activeValue = String(getValue(formulaBar)).trim();
        try {
          document.execCommand("copy");
        } catch (error) {
          return { range, activeValue };
        }
        return { range, activeValue };
      }
    });
    const hit = results.find((entry) => entry.result && entry.result.range);
    return hit ? hit.result : {};
  } catch (error) {
    return {};
  }
};

const isTabOpen = (tabs, targetUrl) => {
  if (!targetUrl) {
    return false;
  }
  const normalized = normalizeUrl(targetUrl);
  return tabs.some((tab) => {
    if (!tab.url) {
      return false;
    }
    if (normalized) {
      return tab.url.includes(normalized);
    }
    return tab.url.includes(targetUrl);
  });
};

const findExcelTab = async (targetUrl) => {
  if (!targetUrl) {
    return null;
  }
  const excelTabs = await safeTabsQuery({
    url: ["*://*.office.com/*", "*://*.live.com/*", "*://*.sharepoint.com/*"]
  });
  const normalized = normalizeUrl(targetUrl);
  return (
    excelTabs.find((tab) => tab.url && tab.url.includes(normalized)) ||
    excelTabs.find((tab) => tab.url && tab.url.includes(targetUrl)) ||
    null
  );
};

const columnLetterToIndex = (letters) => {
  let total = 0;
  const clean = letters.toUpperCase().replace(/[^A-Z]/g, "");
  for (let i = 0; i < clean.length; i += 1) {
    total = total * 26 + (clean.charCodeAt(i) - 64);
  }
  return total;
};

const columnIndexToLetter = (index) => {
  let current = index;
  let result = "";
  while (current > 0) {
    const mod = (current - 1) % 26;
    result = String.fromCharCode(65 + mod) + result;
    current = Math.floor((current - mod - 1) / 26);
  }
  return result || "";
};

const parseCellRef = (value) => {
  const match = String(value || "").toUpperCase().match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    return null;
  }
  return {
    column: match[1],
    row: Number.parseInt(match[2], 10)
  };
};

const normalizeGuideValue = (value) => {
  if (!value) {
    return "";
  }
  const trimmed = String(value).trim();
  if (!trimmed) {
    return "";
  }
  const compact = trimmed.replace(/\s+/g, "");
  if (/^\d+(?:-\d+)+$/.test(compact)) {
    return compact;
  }
  const noCommas = compact.replace(/,/g, "");
  if (/e\+?/i.test(noCommas)) {
    const num = Number(noCommas);
    if (Number.isFinite(num)) {
      return Math.trunc(num).toString();
    }
  }
  const digitsOnly = compact.replace(/[^\d]/g, "");
  if (digitsOnly.length >= 6) {
    return digitsOnly;
  }
  return "";
};

const parseGuidesFromClipboard = (text, mode) => {
  const rows = text.replace(/\r/g, "").split("\n");
  const firstColumn = rows.map((row) => row.split("\t")[0].trim());
  const bolsaLines = firstColumn.filter((line) => /bolsa/i.test(line));
  const hasBolsa = bolsaLines.length > 0;
  const inferredType = bolsaLines.length > 1 ? "columna" : "bolsa";
  let selectionType = mode === "infer" ? inferredType : mode;
  if (selectionType === "bolsa" && !hasBolsa) {
    selectionType = "columna";
  }

  const guides = [];
  const invalid = [];
  const seen = new Set();
  let duplicateCount = 0;
  let emptyRun = 0;
  let inBag = selectionType !== "bolsa";
  let bagLabel = "";

  for (const rawValue of firstColumn) {
    const value = rawValue.trim();
    if (!value) {
      emptyRun += 1;
      if (emptyRun >= 3) {
        break;
      }
      continue;
    }
    emptyRun = 0;
    if (/bolsa/i.test(value)) {
      if (!bagLabel) {
        bagLabel = value;
      }
      if (selectionType === "bolsa") {
        if (!inBag) {
          inBag = true;
          continue;
        }
        break;
      }
      continue;
    }
    if (selectionType === "bolsa" && !inBag) {
      continue;
    }
    const normalized = normalizeGuideValue(value);
    if (!normalized) {
      invalid.push(value);
      continue;
    }
    if (seen.has(normalized)) {
      duplicateCount += 1;
      continue;
    }
    seen.add(normalized);
    guides.push(normalized);
  }

  return {
    guides,
    invalidCount: invalid.length,
    duplicateCount,
    selectionType,
    bagLabel
  };
};

const parseDoc2ExistingFromClipboard = (text) => {
  const rows = text.replace(/\r/g, "").split("\n");
  const guides = new Set();
  const asinsByGuide = {};
  let maxIndex = 0;

  rows.forEach((row) => {
    const cols = row.split("\t").map((col) => col.trim());
    if (!cols.length || cols.every((col) => !col)) {
      return;
    }
    const lineText = cols.join(" ").toLowerCase();
    if (lineText.includes("asin") || lineText.includes("producto")) {
      return;
    }
    const asin = cols[1] || "";
    const guideRaw = cols[3] || "";
    const guide = normalizeGuideValue(guideRaw);
    const indexValue = Number.parseInt(cols[0], 10);
    if (Number.isFinite(indexValue) && indexValue > maxIndex) {
      maxIndex = indexValue;
    }
    if (!asin || !guide) {
      return;
    }
    guides.add(guide);
    if (!asinsByGuide[guide]) {
      asinsByGuide[guide] = [];
    }
    if (!asinsByGuide[guide].includes(asin)) {
      asinsByGuide[guide].push(asin);
    }
  });

  return {
    guides: Array.from(guides),
    asinsByGuide,
    maxIndex
  };
};

const findSellerTab = async () => {
  const tabs = await safeTabsQuery({ url: SELLER_URL_PATTERNS }, (tab) =>
    isSellerUrl(tab.url || "")
  );
  const active = tabs.find((tab) => tab.active);
  return active || tabs[0] || null;
};

const buildOutputRows = (results) => {
  const mainRows = [];
  const manualRows = [];
  const seenGuides = new Set([
    ...currentSettings.dedupe.guides,
    ...currentSettings.dedupe.doc2Guides
  ]);
  const localGuides = new Set();
  let nextIndex = currentSettings.doc2Selection.nextIndex || 1;

  const pushManual = (guide, asin, quantity, reason) => {
    manualRows.push({
      guide,
      asin,
      quantity,
      reason
    });
  };

  results.forEach((result) => {
    const guide = result.guide || "";
    if (!guide) {
      return;
    }
    if (currentSettings.validations.dedupeGuide) {
      if (seenGuides.has(guide) || localGuides.has(guide)) {
        pushManual(guide, "", "", "Guia duplicada");
        return;
      }
    }

    localGuides.add(guide);

    if (result.status !== "found" || !Array.isArray(result.items) || !result.items.length) {
      pushManual(guide, "", "", "No encontrado");
      return;
    }

    const grouped = new Map();
    result.items.forEach((item) => {
      if (!item || !item.asin) {
        return;
      }
      const key = item.asin;
      const qty = Number(item.quantity) || 1;
      grouped.set(key, (grouped.get(key) || 0) + qty);
    });

    const entries = Array.from(grouped.entries());
    if (!entries.length) {
      pushManual(guide, "", "", "ASIN no encontrado");
      return;
    }

    const hasMultipleAsins = entries.length > 1;
    const hasQtyOverOne = entries.some(([, qty]) => qty > 1);
    if (hasMultipleAsins || hasQtyOverOne) {
      const reason = hasMultipleAsins ? "Varios productos por guia" : "Cantidad mayor a 1";
      entries.forEach(([asin, qty]) => {
        pushManual(guide, asin, qty, reason);
      });
      return;
    }

    const [asin, qty] = entries[0];
    mainRows.push({
      index: nextIndex,
      asin,
      quantity: qty,
      guide
    });
    nextIndex += 1;
  });

  return {
    mainRows,
    manualRows,
    nextIndex
  };
};

const formatManualEntry = (entry) => {
  const guide = entry.guide || "-";
  const asin = entry.asin || "-";
  const qty = entry.quantity || "-";
  const reason = entry.reason || "Revisar";
  return `Guia: ${guide} | ASIN: ${asin} | Cantidad: ${qty} | Motivo: ${reason}`;
};

const refreshStatus = async () => {
  const sellerTabs = await safeTabsQuery({ url: SELLER_URL_PATTERNS }, (tab) =>
    isSellerUrl(tab.url || "")
  );
  const excelTabs = await safeTabsQuery({
    url: ["*://*.office.com/*", "*://*.live.com/*", "*://*.sharepoint.com/*"]
  });

  const sellerReady = sellerTabs.length > 0;
  const doc1Ready = isTabOpen(excelTabs, currentSettings.excelDoc1Url);
  const doc2Ready = isTabOpen(excelTabs, currentSettings.excelDoc2Url);

  setStatusDot("sellerStatus", sellerReady);
  setStatusDot("doc1Status", doc1Ready);
  setStatusDot("doc2Status", doc2Ready);

  const statusSession = document.getElementById("statusSession");
  if (sellerReady && doc1Ready && doc2Ready) {
    statusSession.textContent = getString("statusReady");
  } else if (sellerReady) {
    statusSession.textContent = getString("statusNeedsTabs");
  } else {
    statusSession.textContent = getString("statusPending");
  }
};

const renderBolsas = () => {
  const list = document.getElementById("bolsaList");
  list.innerHTML = "";
  if (currentSettings.bolsas.length === 0) {
    const empty = document.createElement("div");
    empty.className = "hint";
    empty.textContent = getString("bolsasHint");
    list.appendChild(empty);
    return;
  }
  currentSettings.bolsas.forEach((item) => {
    const tag = document.createElement("span");
    tag.className = "tag";
    tag.textContent = item;

    const remove = document.createElement("button");
    remove.type = "button";
    remove.textContent = "x";
    remove.addEventListener("click", () => {
      currentSettings.bolsas = currentSettings.bolsas.filter((value) => value !== item);
      chrome.storage.local.set({ bolsas: currentSettings.bolsas });
      renderBolsas();
    });

    tag.appendChild(remove);
    list.appendChild(tag);
  });
};

const renderLogList = (list, containerId, emptyKey) => {
  const container = document.getElementById(containerId);
  container.innerHTML = "";
  if (!list.length) {
    const empty = document.createElement("div");
    empty.className = "hint";
    empty.textContent = getString(emptyKey);
    container.appendChild(empty);
    return;
  }
  list.slice(0, 20).forEach((entry) => {
    const row = document.createElement("div");
    row.className = "log-entry";
    const message = document.createElement("div");
    message.textContent = entry.message || "";
    const meta = document.createElement("small");
    const time = entry.ts ? new Date(entry.ts).toLocaleString() : "-";
    meta.textContent = `${entry.level || "info"} - ${time} - ${entry.source || "app"}`;
    row.appendChild(message);
    row.appendChild(meta);
    container.appendChild(row);
  });
};

const updateDoc1Summary = () => {
  document.getElementById("doc1Range").textContent =
    currentSettings.doc1Selection.range || "-";
  document.getElementById("doc1Type").textContent =
    currentSettings.doc1Selection.selectionType || "-";
  document.getElementById("doc1Count").textContent =
    currentSettings.doc1Selection.guides.length || 0;
  document.getElementById("doc1InvalidCount").textContent =
    currentSettings.doc1Selection.invalidCount || 0;
  document.getElementById("doc1Mode").value = currentSettings.doc1Selection.mode || "infer";
};

const updateDoc2Summary = () => {
  document.getElementById("doc2StartCell").textContent =
    currentSettings.doc2Selection.startCell || "-";
  document.getElementById("doc2ManualCol").textContent =
    currentSettings.doc2Selection.manualColumn || "-";
  document.getElementById("doc2ApprovalStatus").textContent = currentSettings
    .doc2Selection.approved
    ? getString("approvalReady")
    : getString("approvalPending");
};

const updateIntegrityStats = () => {
  const uniqueGuideSet = new Set([
    ...currentSettings.dedupe.guides,
    ...currentSettings.dedupe.doc2Guides
  ]);
  const asinSet = new Set();
  const mergeAsins = (map) => {
    Object.values(map || {}).forEach((list) => {
      if (!Array.isArray(list)) {
        return;
      }
      list.forEach((asin) => asinSet.add(asin));
    });
  };
  mergeAsins(currentSettings.dedupe.asinsByGuide);
  mergeAsins(currentSettings.dedupe.doc2AsinsByGuide);

  document.getElementById("uniqueGuidesCount").textContent = uniqueGuideSet.size;
  document.getElementById("uniqueAsinsCount").textContent = asinSet.size;
  document.getElementById("lastRun").textContent = currentSettings.stats.lastRun || "-";
  document.getElementById("lastError").textContent = currentSettings.stats.lastError || "-";
};

const applySettingsToUI = () => {
  document.getElementById("languageSelect").value = currentSettings.language || "es";
  document.getElementById("regionSelect").value = currentSettings.region || "mx";
  document.getElementById("doc1Url").value = currentSettings.excelDoc1Url || "";
  document.getElementById("doc2Url").value = currentSettings.excelDoc2Url || "";
  document.getElementById("sheetPrimary").value = currentSettings.sheetPrimary || "";
  document.getElementById("sheetDuplicates").value =
    currentSettings.sheetDuplicates || "Duplicados";

  document.getElementById("dedupeGuide").checked = currentSettings.validations.dedupeGuide;
  document.getElementById("dedupeAsin").checked = currentSettings.validations.dedupeAsin;
  document.getElementById("accumulateQty").checked =
    currentSettings.validations.accumulateQuantity;
  document.getElementById("splitDuplicates").checked =
    currentSettings.validations.splitDuplicatesToSheet;

  document.querySelectorAll('input[name="mode"]').forEach((input) => {
    input.checked = input.value === (currentSettings.assistedMode ? "assisted" : "direct");
  });

  const assisted = document.getElementById("assistedMode");
  const direct = document.getElementById("directMode");
  assisted.classList.toggle("hidden", !currentSettings.assistedMode);
  direct.classList.toggle("hidden", currentSettings.assistedMode);

  renderBolsas();
  renderLogList(currentSettings.logs, "logList", "logsEmpty");
  renderLogList(currentSettings.history, "historyList", "historyEmpty");
  updateDoc1Summary();
  updateDoc2Summary();
  updateIntegrityStats();
  applyTranslations();
};

const saveConfig = async () => {
  const prevDoc1Url = currentSettings.excelDoc1Url;
  const prevDoc2Url = currentSettings.excelDoc2Url;
  currentSettings.excelDoc1Url = document.getElementById("doc1Url").value.trim();
  currentSettings.excelDoc2Url = document.getElementById("doc2Url").value.trim();
  currentSettings.sheetPrimary = document.getElementById("sheetPrimary").value.trim();
  currentSettings.sheetDuplicates = document
    .getElementById("sheetDuplicates")
    .value.trim();

  if (prevDoc1Url && prevDoc1Url !== currentSettings.excelDoc1Url) {
    currentSettings.doc1Selection = { ...DEFAULT_SETTINGS.doc1Selection };
  }
  if (prevDoc2Url && prevDoc2Url !== currentSettings.excelDoc2Url) {
    currentSettings.doc2Selection = { ...DEFAULT_SETTINGS.doc2Selection };
    currentSettings.pendingOutput = { ...DEFAULT_SETTINGS.pendingOutput };
    currentSettings.dedupe = {
      ...currentSettings.dedupe,
      doc2Guides: [],
      doc2AsinsByGuide: {}
    };
  }
  await chrome.storage.local.set({
    excelDoc1Url: currentSettings.excelDoc1Url,
    excelDoc2Url: currentSettings.excelDoc2Url,
    sheetPrimary: currentSettings.sheetPrimary,
    sheetDuplicates: currentSettings.sheetDuplicates,
    doc1Selection: currentSettings.doc1Selection,
    doc2Selection: currentSettings.doc2Selection,
    pendingOutput: currentSettings.pendingOutput,
    dedupe: currentSettings.dedupe
  });
  await refreshStatus();
  logEvent("info", "Configuracion guardada.");
  applySettingsToUI();
};

const captureDoc1Selection = async () => {
  if (!currentSettings.excelDoc1Url) {
    logEvent("error", "Ingresa el link del Documento 1 primero.");
    await setLastError("Documento 1 sin link.");
    return;
  }
  const tab = await findExcelTab(currentSettings.excelDoc1Url);
  if (!tab) {
    logEvent("error", "Documento 1 no esta abierto en Excel Online.");
    await setLastError("Documento 1 no esta abierto.");
    return;
  }

  let response = null;
  try {
    response = await chrome.tabs.sendMessage(tab.id, { type: "CAPTURE_SELECTION" });
  } catch (error) {
    response = null;
  }

  let clipboardText = (response && response.clipboardText) || "";
  if (!clipboardText) {
    try {
      clipboardText = await navigator.clipboard.readText();
    } catch (error) {
      clipboardText = "";
    }
  }
  if (!clipboardText) {
    logEvent(
      "error",
      "No se pudo leer la seleccion. Selecciona el rango y presiona Ctrl+C, luego Capturar."
    );
    await setLastError("No se pudo leer la seleccion.");
    return;
  }

  let range = (response && response.range) || "";
  let column = (response && response.column) || "";
  if (!range) {
    const fallback = await fallbackGetExcelSelection(tab.id);
    range = fallback.range || "";
    column = extractColumnFromRange(range);
  }

  const parsed = parseGuidesFromClipboard(
    clipboardText,
    currentSettings.doc1Selection.mode || "infer"
  );

  currentSettings.doc1Selection = {
    ...currentSettings.doc1Selection,
    range,
    column,
    selectionType: parsed.selectionType,
    bagLabel: parsed.bagLabel,
    guides: parsed.guides,
    invalidCount: parsed.invalidCount,
    duplicateCount: parsed.duplicateCount,
    capturedAt: Date.now()
  };

  await chrome.storage.local.set({ doc1Selection: currentSettings.doc1Selection });
  updateDoc1Summary();
  logEvent(
    "info",
    `Documento 1 capturado. Guias: ${parsed.guides.length}, invalidas: ${parsed.invalidCount}.`
  );
  if (!parsed.guides.length) {
    logEvent(
      "error",
      "No se detectaron guias. Verifica la seleccion y que este copiada."
    );
    await setLastError("No se detectaron guias.");
    return;
  }
  await setLastError("");
};

const captureDoc2Selection = async () => {
  if (!currentSettings.excelDoc2Url) {
    logEvent("error", "Ingresa el link del Documento 2 primero.");
    await setLastError("Documento 2 sin link.");
    return;
  }
  const tab = await findExcelTab(currentSettings.excelDoc2Url);
  if (!tab) {
    logEvent("error", "Documento 2 no esta abierto en Excel Online.");
    await setLastError("Documento 2 no esta abierto.");
    return;
  }

  let response = null;
  try {
    response = await chrome.tabs.sendMessage(tab.id, { type: "GET_ACTIVE_CELL" });
  } catch (error) {
    response = null;
  }

  let activeCell = response ? response.activeCell : "";
  let activeValue = response ? response.activeValue || "" : "";
  if (!activeCell) {
    const fallback = await fallbackGetExcelSelection(tab.id);
    activeCell = fallback.range || "";
    activeValue = fallback.activeValue || "";
  }
  const parsed = parseCellRef(activeCell);
  if (!parsed) {
    logEvent(
      "error",
      "No se pudo leer la celda. Haz clic en la celda inicial y presiona Capturar."
    );
    await setLastError("Seleccion invalida en Documento 2.");
    return;
  }

  const baseIndex = columnLetterToIndex(parsed.column);
  const manualColumn = columnIndexToLetter(baseIndex + 7);
  const headerPresent = /no\.?\s*producto/i.test(activeValue);

  currentSettings.doc2Selection = {
    ...currentSettings.doc2Selection,
    startCell: activeCell,
    startColumn: parsed.column,
    startRow: parsed.row,
    manualColumn,
    approved: false,
    headersEnsured: headerPresent,
    capturedAt: Date.now()
  };

  await chrome.storage.local.set({ doc2Selection: currentSettings.doc2Selection });
  updateDoc2Summary();
  logEvent("info", "Documento 2 capturado. Falta aprobar la seleccion.");
  await setLastError("");
};

const approveDoc2Selection = async () => {
  if (!currentSettings.doc2Selection.startCell) {
    logEvent("error", "Captura la celda inicial del Documento 2 primero.");
    return;
  }
  currentSettings.doc2Selection = {
    ...currentSettings.doc2Selection,
    approved: true
  };
  await chrome.storage.local.set({ doc2Selection: currentSettings.doc2Selection });
  updateDoc2Summary();
  logEvent("info", "Seleccion de Documento 2 aprobada.");
};

const captureDoc2Existing = async () => {
  if (!currentSettings.excelDoc2Url) {
    logEvent("error", "Ingresa el link del Documento 2 primero.");
    await setLastError("Documento 2 sin link.");
    return;
  }
  const tab = await findExcelTab(currentSettings.excelDoc2Url);
  if (!tab) {
    logEvent("error", "Documento 2 no esta abierto en Excel Online.");
    await setLastError("Documento 2 no esta abierto.");
    return;
  }

  let response = null;
  try {
    response = await chrome.tabs.sendMessage(tab.id, { type: "CAPTURE_SELECTION" });
  } catch (error) {
    response = null;
  }

  let clipboardText = (response && response.clipboardText) || "";
  if (!clipboardText) {
    try {
      clipboardText = await navigator.clipboard.readText();
    } catch (error) {
      clipboardText = "";
    }
  }
  if (!clipboardText) {
    logEvent(
      "error",
      "No se pudo leer la tabla. Selecciona la tabla y presiona Ctrl+C antes de capturar."
    );
    await setLastError("No se pudo leer la tabla.");
    return;
  }

  const parsed = parseDoc2ExistingFromClipboard(clipboardText);
  currentSettings.dedupe = {
    ...currentSettings.dedupe,
    doc2Guides: parsed.guides,
    doc2AsinsByGuide: parsed.asinsByGuide
  };
  if (parsed.maxIndex && parsed.maxIndex > 0) {
    currentSettings.doc2Selection.nextIndex = parsed.maxIndex + 1;
  }
  await chrome.storage.local.set({
    dedupe: currentSettings.dedupe,
    doc2Selection: currentSettings.doc2Selection
  });
  updateIntegrityStats();
  logEvent("info", `Documento 2 capturado. Guias existentes: ${parsed.guides.length}.`);
  await setLastError("");
};

const runAmazonLookup = async () => {
  const guides = currentSettings.doc1Selection.guides || [];
  if (!guides.length) {
    logEvent("error", "No hay guias capturadas en Documento 1.");
    await setLastError("Sin guias capturadas.");
    return;
  }
  const sellerTab = await findSellerTab();
  if (!sellerTab) {
    logEvent("error", "No hay pestana activa de Amazon Seller.");
    await setLastError("Amazon Seller no esta abierto.");
    return;
  }

  let response = null;
  try {
    response = await chrome.tabs.sendMessage(sellerTab.id, {
      type: "LOOKUP_GUIDES",
      guides
    });
  } catch (error) {
    logEvent("error", "No se pudo comunicar con Amazon Seller.");
    await setLastError("No se pudo comunicar con Amazon Seller.");
    return;
  }

  if (!response || !response.ok) {
    logEvent("error", "Amazon Seller no devolvio resultados.");
    await setLastError("Amazon Seller sin resultados.");
    return;
  }

  const output = buildOutputRows(response.results || []);
  currentSettings.pendingOutput = {
    mainRows: output.mainRows,
    manualRows: output.manualRows,
    nextIndex: output.nextIndex,
    createdAt: Date.now()
  };
  await chrome.storage.local.set({
    pendingOutput: currentSettings.pendingOutput
  });

  logEvent(
    "info",
    `Amazon listo. Principal: ${output.mainRows.length}, revision: ${output.manualRows.length}.`
  );
  await setLastError("");
};

const writeExcelOutput = async () => {
  if (!currentSettings.doc2Selection.approved) {
    logEvent("error", "Aprueba la seleccion del Documento 2 antes de escribir.");
    await setLastError("Seleccion de Documento 2 no aprobada.");
    return;
  }
  if (!currentSettings.pendingOutput.mainRows.length && !currentSettings.pendingOutput.manualRows.length) {
    logEvent("error", "No hay datos pendientes para escribir.");
    await setLastError("Sin datos pendientes para escribir.");
    return;
  }
  const tab = await findExcelTab(currentSettings.excelDoc2Url);
  if (!tab) {
    logEvent("error", "Documento 2 no esta abierto en Excel Online.");
    await setLastError("Documento 2 no esta abierto.");
    return;
  }

  const rows = [
    [
      "#No. Producto",
      "ASIN",
      "Cantidad",
      "No. Guia",
      "",
      "",
      "",
      "Revisar Manualmente"
    ]
  ];

  currentSettings.pendingOutput.mainRows.forEach((row) => {
    rows.push([
      String(row.index),
      row.asin,
      String(row.quantity),
      row.guide,
      "",
      "",
      "",
      ""
    ]);
  });

  currentSettings.pendingOutput.manualRows.forEach((entry) => {
    rows.push(["", "", "", "", "", "", "", formatManualEntry(entry)]);
  });

  const tsv = rows.map((row) => row.join("\t")).join("\n");
  let response = {};
  try {
    response = await chrome.tabs.sendMessage(tab.id, { type: "PASTE_TSV", tsv });
  } catch (error) {
    logEvent("error", "No se pudo pegar en Documento 2.");
    await setLastError("No se pudo pegar en Documento 2.");
    return;
  }

  if (!response.clipboardOk) {
    logEvent("error", "No se pudo copiar al portapapeles.");
    await setLastError("No se pudo copiar al portapapeles.");
    return;
  }
  if (!response.pasteOk) {
    logEvent("error", "Pegado automatico bloqueado. Pega manualmente con Ctrl+V.");
    await setLastError("Pegado automatico bloqueado.");
    return;
  }

  const updatedGuides = new Set(currentSettings.dedupe.guides);
  const updatedAsinsByGuide = { ...currentSettings.dedupe.asinsByGuide };
  const allEntries = [
    ...currentSettings.pendingOutput.mainRows.map((row) => ({
      guide: row.guide,
      asin: row.asin
    })),
    ...currentSettings.pendingOutput.manualRows.map((row) => ({
      guide: row.guide,
      asin: row.asin
    }))
  ];
  allEntries.forEach((entry) => {
    if (!entry.guide) {
      return;
    }
    updatedGuides.add(entry.guide);
    if (!updatedAsinsByGuide[entry.guide]) {
      updatedAsinsByGuide[entry.guide] = [];
    }
    if (entry.asin && !updatedAsinsByGuide[entry.guide].includes(entry.asin)) {
      updatedAsinsByGuide[entry.guide].push(entry.asin);
    }
  });

  currentSettings.dedupe = {
    ...currentSettings.dedupe,
    guides: Array.from(updatedGuides),
    asinsByGuide: updatedAsinsByGuide
  };
  currentSettings.doc2Selection = {
    ...currentSettings.doc2Selection,
    nextIndex: currentSettings.pendingOutput.nextIndex || currentSettings.doc2Selection.nextIndex
  };
  currentSettings.stats = {
    ...currentSettings.stats,
    lastRun: new Date().toLocaleString(),
    lastError: ""
  };
  const mainCount = currentSettings.pendingOutput.mainRows.length;
  const manualCount = currentSettings.pendingOutput.manualRows.length;
  currentSettings.pendingOutput = {
    mainRows: [],
    manualRows: [],
    nextIndex: currentSettings.doc2Selection.nextIndex,
    createdAt: 0
  };
  const historyEntry = {
    id: crypto.randomUUID(),
    ts: Date.now(),
    message: `Escritura completada. Principal: ${mainCount}, revision: ${manualCount}.`
  };
  currentSettings.history = [historyEntry, ...currentSettings.history].slice(0, 200);

  await chrome.storage.local.set({
    dedupe: currentSettings.dedupe,
    stats: currentSettings.stats,
    pendingOutput: currentSettings.pendingOutput,
    doc2Selection: currentSettings.doc2Selection,
    history: currentSettings.history
  });
  updateIntegrityStats();
  logEvent("info", "Documento 2 actualizado correctamente.");
};

const runAllWorkflow = async () => {
  await refreshStatus();
  if (!currentSettings.doc1Selection.guides.length) {
    await captureDoc1Selection();
  }
  await runAmazonLookup();
  await writeExcelOutput();
};

const initTabs = () => {
  const tabs = document.querySelectorAll(".tab");
  tabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      const target = tab.getAttribute("data-tab");
      tabs.forEach((btn) => btn.classList.remove("active"));
      tab.classList.add("active");
      document.querySelectorAll(".panel").forEach((panel) => {
        panel.classList.toggle("active", panel.id === `panel-${target}`);
      });
    });
  });
};

const initEvents = () => {
  document.getElementById("saveConfig").addEventListener("click", saveConfig);
  document.getElementById("addBolsa").addEventListener("click", () => {
    const input = document.getElementById("bolsaInput");
    const value = input.value.trim();
    if (!value) {
      return;
    }
    const exists = currentSettings.bolsas.some(
      (item) => item.toLowerCase() === value.toLowerCase()
    );
    if (!exists) {
      currentSettings.bolsas = [...currentSettings.bolsas, value];
      chrome.storage.local.set({ bolsas: currentSettings.bolsas });
    }
    input.value = "";
    renderBolsas();
  });

  document.getElementById("languageSelect").addEventListener("change", (event) => {
    currentSettings.language = event.target.value;
    chrome.storage.local.set({ language: currentSettings.language });
    applySettingsToUI();
    refreshStatus();
  });

  document.getElementById("regionSelect").addEventListener("change", (event) => {
    currentSettings.region = event.target.value;
    chrome.storage.local.set({ region: currentSettings.region });
  });

  document.getElementById("doc1Mode").addEventListener("change", (event) => {
    currentSettings.doc1Selection.mode = event.target.value;
    chrome.storage.local.set({ doc1Selection: currentSettings.doc1Selection });
  });

  document.getElementById("captureDoc1").addEventListener("click", captureDoc1Selection);
  document.getElementById("captureDoc2").addEventListener("click", captureDoc2Selection);
  document.getElementById("approveDoc2").addEventListener("click", approveDoc2Selection);
  document
    .getElementById("captureDoc2Existing")
    .addEventListener("click", captureDoc2Existing);

  document.querySelectorAll('input[name="mode"]').forEach((input) => {
    input.addEventListener("change", (event) => {
      currentSettings.assistedMode = event.target.value === "assisted";
      chrome.storage.local.set({ assistedMode: currentSettings.assistedMode });
      applySettingsToUI();
    });
  });

  document.getElementById("dedupeGuide").addEventListener("change", (event) => {
    currentSettings.validations.dedupeGuide = event.target.checked;
    chrome.storage.local.set({ validations: currentSettings.validations });
  });
  document.getElementById("dedupeAsin").addEventListener("change", (event) => {
    currentSettings.validations.dedupeAsin = event.target.checked;
    chrome.storage.local.set({ validations: currentSettings.validations });
  });
  document.getElementById("accumulateQty").addEventListener("change", (event) => {
    currentSettings.validations.accumulateQuantity = event.target.checked;
    chrome.storage.local.set({ validations: currentSettings.validations });
  });
  document.getElementById("splitDuplicates").addEventListener("change", (event) => {
    currentSettings.validations.splitDuplicatesToSheet = event.target.checked;
    chrome.storage.local.set({ validations: currentSettings.validations });
  });

  document.getElementById("refreshStatus").addEventListener("click", refreshStatus);
  document.getElementById("checkTabs").addEventListener("click", refreshStatus);

  document.getElementById("startAssisted").addEventListener("click", () => {
    runAllWorkflow();
  });

  document.getElementById("scanBolsas").addEventListener("click", captureDoc1Selection);

  document.getElementById("scanAmazon").addEventListener("click", () => {
    runAmazonLookup();
  });

  document.getElementById("writeExcel").addEventListener("click", () => {
    writeExcelOutput();
  });

  document.getElementById("runAll").addEventListener("click", () => {
    runAllWorkflow();
  });

  document.getElementById("clearLogs").addEventListener("click", async () => {
    await chrome.runtime.sendMessage({ type: "CLEAR_LOGS" });
    currentSettings.logs = [];
    renderLogList(currentSettings.logs, "logList", "logsEmpty");
  });

  document.getElementById("clearHistory").addEventListener("click", async () => {
    await chrome.runtime.sendMessage({ type: "CLEAR_HISTORY" });
    currentSettings.history = [];
    renderLogList(currentSettings.history, "historyList", "historyEmpty");
  });

  chrome.storage.onChanged.addListener((changes) => {
    if (changes.logs) {
      currentSettings.logs = changes.logs.newValue || [];
      renderLogList(currentSettings.logs, "logList", "logsEmpty");
    }
    if (changes.history) {
      currentSettings.history = changes.history.newValue || [];
      renderLogList(currentSettings.history, "historyList", "historyEmpty");
    }
    if (changes.doc1Selection) {
      currentSettings.doc1Selection = changes.doc1Selection.newValue;
      updateDoc1Summary();
    }
    if (changes.doc2Selection) {
      currentSettings.doc2Selection = changes.doc2Selection.newValue;
      updateDoc2Summary();
    }
    if (changes.dedupe) {
      currentSettings.dedupe = changes.dedupe.newValue;
      updateIntegrityStats();
    }
    if (changes.stats) {
      currentSettings.stats = changes.stats.newValue;
      updateIntegrityStats();
    }
    if (changes.pendingOutput) {
      currentSettings.pendingOutput = changes.pendingOutput.newValue || {
        mainRows: [],
        manualRows: [],
        nextIndex: 1,
        createdAt: 0
      };
    }
  });
};

const init = async () => {
  initTabs();
  initEvents();

  const stored = await chrome.storage.local.get();
  currentSettings = mergeDefaults(stored);
  applySettingsToUI();
  await refreshStatus();
};

document.addEventListener("DOMContentLoaded", init);
