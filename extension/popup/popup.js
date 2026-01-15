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

const refreshStatus = async () => {
  const sellerTabs = await chrome.tabs.query({
    url: ["*://sellercentral.amazon.*/*"]
  });
  const excelTabs = await chrome.tabs.query({
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
    meta.textContent = `${entry.level || "info"} • ${time} • ${entry.source || "app"}`;
    row.appendChild(message);
    row.appendChild(meta);
    container.appendChild(row);
  });
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
  applyTranslations();
};

const saveConfig = async () => {
  currentSettings.excelDoc1Url = document.getElementById("doc1Url").value.trim();
  currentSettings.excelDoc2Url = document.getElementById("doc2Url").value.trim();
  currentSettings.sheetPrimary = document.getElementById("sheetPrimary").value.trim();
  currentSettings.sheetDuplicates = document
    .getElementById("sheetDuplicates")
    .value.trim();
  await chrome.storage.local.set({
    excelDoc1Url: currentSettings.excelDoc1Url,
    excelDoc2Url: currentSettings.excelDoc2Url,
    sheetPrimary: currentSettings.sheetPrimary,
    sheetDuplicates: currentSettings.sheetDuplicates
  });
  await refreshStatus();
  chrome.runtime.sendMessage({
    type: "LOG_EVENT",
    level: "info",
    source: "popup",
    message: "Configuracion guardada."
  });
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
    chrome.runtime.sendMessage({
      type: "LOG_EVENT",
      level: "info",
      source: "popup",
      message: "Flujo asistido pendiente de implementar."
    });
  });

  document.getElementById("scanBolsas").addEventListener("click", () => {
    chrome.runtime.sendMessage({
      type: "LOG_EVENT",
      level: "info",
      source: "popup",
      message: "Lectura de bolsas pendiente de implementar."
    });
  });

  document.getElementById("scanAmazon").addEventListener("click", () => {
    chrome.runtime.sendMessage({
      type: "LOG_EVENT",
      level: "info",
      source: "popup",
      message: "Busqueda en Amazon pendiente de implementar."
    });
  });

  document.getElementById("writeExcel").addEventListener("click", () => {
    chrome.runtime.sendMessage({
      type: "LOG_EVENT",
      level: "info",
      source: "popup",
      message: "Escritura en Excel pendiente de implementar."
    });
  });

  document.getElementById("runAll").addEventListener("click", () => {
    chrome.runtime.sendMessage({
      type: "LOG_EVENT",
      level: "info",
      source: "popup",
      message: "Ejecucion completa pendiente de implementar."
    });
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
