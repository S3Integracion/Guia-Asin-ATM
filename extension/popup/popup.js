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
  sheetDoc1: "",
  sheetPrimary: "",
  sheetDuplicates: "Duplicados",
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

const EXCEL_URL_PATTERNS = [
  "*://*.office.com/*",
  "*://*.live.com/*",
  "*://onedrive.live.com/*",
  "*://*.sharepoint.com/*",
  "*://officeapps.live.com/*",
  "*://*.officeapps.live.com/*",
  "*://excel.officeapps.live.com/*",
  "*://excel.office.com/*",
  "*://excel.cloud.microsoft.com/*",
  "*://excel.cloud.microsoft/*"
];

const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
const GRAPH_SCOPES = [
  "offline_access",
  "Files.ReadWrite.All",
  "Sites.ReadWrite.All",
  "User.Read"
];
const MS_TOKEN_BUFFER_MS = 2 * 60 * 1000;

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
    sheetDoc1Label: "Hoja Documento 1",
    sheetPrimaryLabel: "Hoja principal",
    sheetDupLabel: "Hoja duplicados",
    msAuthTitle: "Microsoft Graph",
    msClientIdLabel: "Client ID",
    msTenantLabel: "Tenant",
    msConnectButton: "Conectar Microsoft",
    msStatusIdle: "Sin sesion",
    msStatusOk: "Conectado",
    msStatusExpired: "Sesion expirada",
    msAuthHint: "La sesion se guarda localmente en la extension.",
    saveConfigButton: "Guardar configuracion",
    doc1SelectionTitle: "Seleccion Documento 1",
    doc1SelectionHint:
      "Selecciona la columna o bolsa en Excel y presiona Capturar, o escribe la columna manual.",
    doc1ModeLabel: "Modo de seleccion",
    doc1ModeInfer: "Inferir automaticamente",
    doc1ModeBolsa: "Bolsa",
    doc1ModeColumna: "Columna",
    doc1ManualColLabel: "Columna manual (opcional)",
    captureDoc1Button: "Capturar seleccion",
    doc1RangeLabel: "Rango detectado",
    doc1TypeLabel: "Tipo",
    doc1GuidesLabel: "Guias listas",
    doc1InvalidLabel: "Guias invalidas",
    doc2SelectionTitle: "Seleccion Documento 2",
    doc2SelectionHint:
      "Selecciona la celda inicial (encabezado) en Excel y captura la salida.",
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
    doc2ValidationHint:
      "Lee la tabla actual (D-G) para validar duplicados.",
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
    sheetDoc1Label: "Document 1 sheet",
    sheetPrimaryLabel: "Primary sheet",
    sheetDupLabel: "Duplicate sheet",
    msAuthTitle: "Microsoft Graph",
    msClientIdLabel: "Client ID",
    msTenantLabel: "Tenant",
    msConnectButton: "Connect Microsoft",
    msStatusIdle: "Signed out",
    msStatusOk: "Connected",
    msStatusExpired: "Session expired",
    msAuthHint: "The session is stored locally in the extension.",
    saveConfigButton: "Save configuration",
    doc1SelectionTitle: "Document 1 selection",
    doc1SelectionHint:
      "Select the column or bag in Excel and press Capture, or type the column manually.",
    doc1ModeLabel: "Selection mode",
    doc1ModeInfer: "Auto infer",
    doc1ModeBolsa: "Bag",
    doc1ModeColumna: "Column",
    doc1ManualColLabel: "Manual column (optional)",
    captureDoc1Button: "Capture selection",
    doc1RangeLabel: "Detected range",
    doc1TypeLabel: "Type",
    doc1GuidesLabel: "Guides ready",
    doc1InvalidLabel: "Invalid guides",
    doc2SelectionTitle: "Document 2 selection",
    doc2SelectionHint:
      "Select the header start cell in Excel and capture the output.",
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
    doc2ValidationHint:
      "Read the current table (D-G) to validate duplicates.",
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
    msAuth: {
      ...DEFAULT_SETTINGS.msAuth,
      ...(stored.msAuth || {})
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

const normalizeRangeToken = (value) => {
  const raw = String(value || "").trim();
  if (!raw) {
    return "";
  }
  const noAbs = raw.replace(/\$/g, "");
  const bangIndex = noAbs.lastIndexOf("!");
  return bangIndex >= 0 ? noAbs.slice(bangIndex + 1).trim() : noAbs;
};

const isRangeRef = (value) => {
  const cleaned = normalizeRangeToken(value).toUpperCase();
  if (!cleaned) {
    return false;
  }
  if (/^[A-Z]{1,3}\d{1,7}$/.test(cleaned)) {
    return true;
  }
  if (/^[A-Z]{1,3}\d{1,7}:[A-Z]{1,3}\d{1,7}$/.test(cleaned)) {
    return true;
  }
  if (/^[A-Z]{1,3}:[A-Z]{1,3}$/.test(cleaned)) {
    return true;
  }
  if (/^\d+:\d+$/.test(cleaned)) {
    return true;
  }
  return false;
};

const isLikelyExcelUrl = (value) => {
  if (!value) {
    return false;
  }
  return /excel|officeapps|office|onedrive|sharepoint/i.test(value);
};

const extractDocSignature = (value) => {
  if (!value) {
    return "";
  }
  try {
    const url = new URL(value);
    const params = [
      "docid",
      "resid",
      "id",
      "itemid",
      "driveitemid",
      "fileid"
    ];
    const lowered = new Map();
    for (const [key, val] of url.searchParams.entries()) {
      lowered.set(key.toLowerCase(), val);
    }
    for (const key of params) {
      const hit = lowered.get(key);
      if (hit) {
        return hit;
      }
    }
    const cid = lowered.get("cid");
    const itemId = lowered.get("id");
    if (cid && itemId) {
      return `${cid}:${itemId}`;
    }
    if (url.hostname.endsWith("1drv.ms")) {
      return url.pathname.replace(/\//g, "");
    }
  } catch (error) {
    return "";
  }
  return "";
};

const isSellerUrl = (value) => {
  try {
    const url = new URL(value);
    return url.hostname.startsWith("sellercentral.amazon.");
  } catch (error) {
    return false;
  }
};

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const safeTabsQuery = async (queryInfo, fallbackFilter) => {
  try {
    return await chrome.tabs.query(queryInfo);
  } catch (error) {
    logEvent("error", `Filtro de tabs invalido: ${error.message || error}`);
    const tabs = await chrome.tabs.query({});
    return fallbackFilter ? tabs.filter(fallbackFilter) : tabs;
  }
};

const safeSendMessage = async (tabId, message, options = {}) => {
  try {
    return await chrome.tabs.sendMessage(tabId, message, options);
  } catch (error) {
    return null;
  }
};

const readClipboardText = async (attempts = 2, delayMs = 120) => {
  for (let i = 0; i < attempts; i += 1) {
    try {
      const text = await navigator.clipboard.readText();
      if (text) {
        return text;
      }
    } catch (error) {
      // Ignore clipboard errors; we retry briefly.
    }
    if (i < attempts - 1) {
      await sleep(delayMs);
    }
  }
  return "";
};

const readExcelClipboardCache = async (tabId) => {
  try {
    const results = await chrome.scripting.executeScript({
      target: { tabId, allFrames: true },
      func: () => ({
        cachedClipboard: window.__GA_LAST_CLIPBOARD__ || "",
        frameUrl: location.href
      })
    });
    const hit = results.find(
      (entry) => entry.result && entry.result.cachedClipboard
    );
    return hit && hit.result ? hit.result.cachedClipboard : "";
  } catch (error) {
    return "";
  }
};

const base64UrlEncode = (buffer) => {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  bytes.forEach((byte) => {
    binary += String.fromCharCode(byte);
  });
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
};

const generateCodeVerifier = () => {
  const bytes = new Uint8Array(32);
  crypto.getRandomValues(bytes);
  return base64UrlEncode(bytes);
};

const sha256 = async (plain) => {
  const encoder = new TextEncoder();
  const data = encoder.encode(plain);
  return crypto.subtle.digest("SHA-256", data);
};

const createCodeChallenge = async (verifier) => {
  const hash = await sha256(verifier);
  return base64UrlEncode(hash);
};

const getMicrosoftRedirectUrl = () => chrome.identity.getRedirectURL("msauth");

const updateMicrosoftStatus = () => {
  const status = document.getElementById("msStatus");
  if (!status) {
    return;
  }
  const expiresAt = currentSettings.msAuth.expiresAt || 0;
  const hasToken = Boolean(currentSettings.msAuth.accessToken);
  const isValid = hasToken && Date.now() < expiresAt - MS_TOKEN_BUFFER_MS;
  if (!hasToken && !currentSettings.msAuth.refreshToken) {
    status.textContent = getString("msStatusIdle");
    return;
  }
  if (!isValid) {
    status.textContent = getString("msStatusExpired");
    return;
  }
  status.textContent = getString("msStatusOk");
};

const saveMicrosoftAuth = async (payload) => {
  currentSettings.msAuth = {
    ...currentSettings.msAuth,
    ...payload
  };
  await chrome.storage.local.set({ msAuth: currentSettings.msAuth });
  updateMicrosoftStatus();
};

const exchangeMicrosoftToken = async (params) => {
  const tenant = currentSettings.msAuth.tenant || "common";
  const response = await fetch(
    `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams(params).toString()
    }
  );
  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Token error: ${response.status} ${text}`);
  }
  return response.json();
};

const startMicrosoftAuth = async () => {
  if (!currentSettings.msAuth.clientId) {
    logEvent("error", "Ingresa el Client ID de Microsoft en Configuracion.");
    await setLastError("Client ID de Microsoft no configurado.");
    return false;
  }
  const verifier = generateCodeVerifier();
  const challenge = await createCodeChallenge(verifier);
  const redirectUrl = getMicrosoftRedirectUrl();
  const tenant = currentSettings.msAuth.tenant || "common";
  const scope = GRAPH_SCOPES.join(" ");
  const authUrl = new URL(
    `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/authorize`
  );
  authUrl.searchParams.set("client_id", currentSettings.msAuth.clientId);
  authUrl.searchParams.set("response_type", "code");
  authUrl.searchParams.set("redirect_uri", redirectUrl);
  authUrl.searchParams.set("response_mode", "query");
  authUrl.searchParams.set("scope", scope);
  authUrl.searchParams.set("code_challenge_method", "S256");
  authUrl.searchParams.set("code_challenge", challenge);

  let responseUrl = "";
  try {
    responseUrl = await chrome.identity.launchWebAuthFlow({
      url: authUrl.toString(),
      interactive: true
    });
  } catch (error) {
    logEvent("error", "No se pudo completar el login de Microsoft.");
    await setLastError("Login de Microsoft cancelado.");
    return false;
  }

  if (!responseUrl) {
    logEvent("error", "Login de Microsoft no devolvio respuesta.");
    await setLastError("Login de Microsoft fallido.");
    return false;
  }

  const parsed = new URL(responseUrl);
  const code = parsed.searchParams.get("code");
  if (!code) {
    logEvent("error", "Microsoft no devolvio codigo de autorizacion.");
    await setLastError("Login de Microsoft fallido.");
    return false;
  }

  const token = await exchangeMicrosoftToken({
    client_id: currentSettings.msAuth.clientId,
    scope,
    grant_type: "authorization_code",
    code,
    redirect_uri: redirectUrl,
    code_verifier: verifier
  });

  const expiresAt = Date.now() + (token.expires_in || 3600) * 1000;
  await saveMicrosoftAuth({
    accessToken: token.access_token,
    refreshToken: token.refresh_token || currentSettings.msAuth.refreshToken,
    expiresAt
  });
  logEvent("info", "Sesion Microsoft conectada.");
  await setLastError("");
  return true;
};

const refreshMicrosoftToken = async () => {
  if (!currentSettings.msAuth.refreshToken) {
    return null;
  }
  const scope = GRAPH_SCOPES.join(" ");
  const token = await exchangeMicrosoftToken({
    client_id: currentSettings.msAuth.clientId,
    scope,
    grant_type: "refresh_token",
    refresh_token: currentSettings.msAuth.refreshToken
  });
  const expiresAt = Date.now() + (token.expires_in || 3600) * 1000;
  await saveMicrosoftAuth({
    accessToken: token.access_token,
    refreshToken: token.refresh_token || currentSettings.msAuth.refreshToken,
    expiresAt
  });
  return token.access_token;
};

const getMicrosoftAccessToken = async () => {
  if (!currentSettings.msAuth.clientId) {
    return "";
  }
  const now = Date.now();
  if (
    currentSettings.msAuth.accessToken &&
    now < (currentSettings.msAuth.expiresAt || 0) - MS_TOKEN_BUFFER_MS
  ) {
    return currentSettings.msAuth.accessToken;
  }
  const refreshed = await refreshMicrosoftToken();
  return refreshed || "";
};

const graphRequest = async (path, options = {}) => {
  const token = await getMicrosoftAccessToken();
  if (!token) {
    throw new Error("Microsoft no autenticado.");
  }
  const response = await fetch(`${GRAPH_BASE_URL}${path}`, {
    ...options,
    headers: {
      ...(options.headers || {}),
      Authorization: `Bearer ${token}`
    }
  });
  if (response.status === 401) {
    const refreshed = await refreshMicrosoftToken();
    if (!refreshed) {
      throw new Error("Microsoft token expirado.");
    }
    const retry = await fetch(`${GRAPH_BASE_URL}${path}`, {
      ...options,
      headers: {
        ...(options.headers || {}),
        Authorization: `Bearer ${refreshed}`
      }
    });
    if (!retry.ok) {
      const text = await retry.text();
      throw new Error(`Graph error: ${retry.status} ${text}`);
    }
    return retry.json();
  }
  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Graph error: ${response.status} ${text}`);
  }
  if (response.status === 204) {
    return {};
  }
  return response.json();
};

const base64UrlEncodeText = (text) => {
  const encoder = new TextEncoder();
  return base64UrlEncode(encoder.encode(text));
};

const buildShareId = (url) => `u!${base64UrlEncodeText(url)}`;

const graphGetWorksheets = async (shareId) =>
  graphRequest(`/shares/${shareId}/driveItem/workbook/worksheets`);

const graphGetWorksheet = async (shareId, sheetName) => {
  const list = await graphGetWorksheets(shareId);
  const sheets = list.value || [];
  if (!sheets.length) {
    throw new Error("No se encontraron hojas en el workbook.");
  }
  if (!sheetName) {
    return sheets[0];
  }
  const byName = sheets.find(
    (sheet) => sheet.name.toLowerCase() === sheetName.toLowerCase()
  );
  return byName || sheets[0];
};

const graphGetRange = async (shareId, sheetId, address) => {
  const safeAddress = String(address || "").replace(/'/g, "''");
  const encoded = encodeURIComponent(safeAddress);
  return graphRequest(
    `/shares/${shareId}/driveItem/workbook/worksheets/${sheetId}/range(address='${encoded}')`
  );
};

const graphGetUsedRange = async (shareId, sheetId) =>
  graphRequest(
    `/shares/${shareId}/driveItem/workbook/worksheets/${sheetId}/usedRange(valuesOnly=true)`
  );

  const graphUpdateRange = async (shareId, sheetId, address, values) => {
    const safeAddress = String(address || "").replace(/'/g, "''");
    const encoded = encodeURIComponent(safeAddress);
    return graphRequest(
      `/shares/${shareId}/driveItem/workbook/worksheets/${sheetId}/range(address='${encoded}')`,
    {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ values })
      }
    );
  };

  const isPersonalTenant = (tenant) => {
    const normalized = String(tenant || "").trim().toLowerCase();
    return normalized === "consumers" || normalized === "consumer";
  };

  const isGraphMsaError = (error) => {
    const message = String(error && error.message ? error.message : error || "")
      .toLowerCase()
      .replace(/\s+/g, " ");
    return (
      message.includes("msa account") ||
      message.includes("msa accounts") ||
      (message.includes("microsoft.excel") && message.includes("not supported")) ||
      (message.includes("addressurl") && message.includes("microsoft.excel"))
    );
  };

const sendExcelMessage = async (tabId, frameId, message) => {
  if (Number.isInteger(frameId)) {
    return safeSendMessage(tabId, message, { frameId });
  }
  return safeSendMessage(tabId, message);
};

const ensureExcelReady = async (tabId) => {
  const injected = await ensureContentScript(tabId, "content/excel_online.js");
  if (injected) {
    return true;
  }
  return probeExcelUi(tabId);
};

const collectExcelCapture = async (tabId, initialClipboardText = "") => {
  const fallback = await fallbackGetExcelSelection(tabId);
  const frameId = Number.isInteger(fallback.frameId) ? fallback.frameId : null;
  let response = null;
  if (frameId !== null) {
    response = await sendExcelMessage(tabId, frameId, { type: "CAPTURE_SELECTION" });
  }
  if (!response) {
    response = await safeSendMessage(tabId, { type: "CAPTURE_SELECTION" });
  }

  let range =
    (response && response.range) ||
    fallback.range ||
    fallback.cachedRange ||
    fallback.cachedCell ||
    "";
  let column = (response && response.column) || extractColumnFromRange(range);
  let clipboardText =
    (response && response.clipboardText) ||
    fallback.clipboardText ||
    fallback.cachedClipboard ||
    initialClipboardText ||
    "";

  if (!clipboardText) {
    const copyResponse =
      (frameId !== null &&
        (await sendExcelMessage(tabId, frameId, { type: "COPY_SELECTION" }))) ||
      (await safeSendMessage(tabId, { type: "COPY_SELECTION" }));
    if (copyResponse && copyResponse.ok) {
      await sleep(160);
      clipboardText = (await readExcelClipboardCache(tabId)) || "";
      if (!clipboardText) {
        clipboardText = await readClipboardText(3, 160);
      }
    }
  }

  if (!clipboardText) {
    clipboardText = initialClipboardText || (await readClipboardText());
  }

  return {
    range,
    column,
    clipboardText,
    fallback,
    frameId,
    response
  };
};

const getExcelActiveCell = async (tabId, frameId) => {
  let response = null;
  if (Number.isInteger(frameId)) {
    response = await sendExcelMessage(tabId, frameId, { type: "GET_ACTIVE_CELL" });
  }
  if (!response) {
    response = await safeSendMessage(tabId, { type: "GET_ACTIVE_CELL" });
  }
  return response;
};

  const getExcelSelectionRange = async (tabId, frameId) => {
    let response = null;
    if (Number.isInteger(frameId)) {
      response = await sendExcelMessage(tabId, frameId, { type: "GET_SELECTION_RANGE" });
    }
  if (!response) {
    response = await safeSendMessage(tabId, { type: "GET_SELECTION_RANGE" });
    }
    return response;
  };

  const readColumnValuesFromUi = async (tabId, frameId, options) => {
    const payload = { type: "READ_COLUMN_VALUES", options };
    let response = null;
    if (Number.isInteger(frameId)) {
      response = await sendExcelMessage(tabId, frameId, payload);
    }
    if (!response) {
      response = await safeSendMessage(tabId, payload);
    }
    return response || { ok: false, reason: "no_response" };
  };

  const setExcelSelection = async (tabId, frameId, range) => {
    const payload = { type: "SET_SELECTION", range };
    let response = null;
  if (Number.isInteger(frameId)) {
    response = await sendExcelMessage(tabId, frameId, payload);
  }
  if (!response) {
    response = await safeSendMessage(tabId, payload);
  }
  if (response && response.ok) {
    return response;
  }
  try {
    const results = await chrome.scripting.executeScript({
      target: { tabId, allFrames: true },
      args: [range],
      func: (targetRange) => {
        const normalizeToken = (value) => {
          const raw = String(value || "").trim();
          if (!raw) {
            return "";
          }
          const noAbs = raw.replace(/\$/g, "");
          const bangIndex = noAbs.lastIndexOf("!");
          return bangIndex >= 0 ? noAbs.slice(bangIndex + 1).trim() : noAbs;
        };
        const selectors = [
          'input[aria-label="Name box"]',
          'input[aria-label="Name Box"]',
          'input[aria-label*="Name"]',
          'input[aria-label*="Nombre"]',
          '[data-automation-id*="nameBox"] input',
          '[data-automation-id*="NameBox"] input',
          '[data-automation-id="formulaNameBox"]',
          '#NameBox input',
          '#NameBox',
          '[role="combobox"][aria-label*="Name"]',
          '[role="combobox"][aria-label*="Nombre"]'
        ];
        const findNameBox = () => {
          for (const selector of selectors) {
            const hit = document.querySelector(selector);
            if (hit) {
              return hit.querySelector("input, textarea") || hit;
            }
          }
          return null;
        };
        const input = findNameBox();
        const clean = normalizeToken(targetRange);
        if (!input || !clean) {
          return { ok: false, reason: "name_box_missing" };
        }
        input.focus();
        if ("value" in input) {
          input.value = clean;
        } else {
          input.textContent = clean;
        }
        input.dispatchEvent(new Event("input", { bubbles: true }));
        input.dispatchEvent(new Event("change", { bubbles: true }));
        input.dispatchEvent(
          new KeyboardEvent("keydown", { key: "Enter", code: "Enter", bubbles: true })
        );
        input.dispatchEvent(
          new KeyboardEvent("keyup", { key: "Enter", code: "Enter", bubbles: true })
        );
        return { ok: true };
      }
    });
    const hit = results.find((entry) => entry.result && entry.result.ok);
    if (hit && hit.result) {
      return hit.result;
    }
  } catch (error) {
    return response;
  }
  return response;
};

const ensureContentScript = async (tabId, filePath) => {
  const ping = await safeSendMessage(tabId, { type: "PING" });
  if (ping) {
    return true;
  }
  try {
    await chrome.scripting.executeScript({
      target: { tabId, allFrames: true },
      files: [filePath]
    });
    await sleep(200);
    const retry = await safeSendMessage(tabId, { type: "PING" });
    return Boolean(retry);
  } catch (error) {
    return false;
  }
};

const probeExcelUi = async (tabId) => {
  try {
    const results = await chrome.scripting.executeScript({
      target: { tabId, allFrames: true },
      func: () => {
        const selectors = [
          'input[aria-label="Name box"]',
          'input[aria-label="Name Box"]',
          'input[aria-label*="Name"]',
          'input[aria-label*="Nombre"]',
          'div[contenteditable="true"][aria-label*="Name"]',
          'div[contenteditable="true"][aria-label*="Nombre"]',
          'input[aria-label="Cuadro de nombre"]',
          'input[aria-label="Cuadro de nombres"]',
          'input[title*="Name"]',
          'input[title*="Nombre"]',
          '[data-automation-id*="nameBox"]',
          '[data-automation-id*="NameBox"]',
          '[data-automation-id="formulaNameBox"]',
          '[data-testid*="nameBox"]'
        ];
        const formulaSelectors = [
          'input[aria-label="Formula Bar"]',
          'input[aria-label="Barra de formulas"]',
          'input[aria-label*="Formula"]',
          'textarea[aria-label*="Formula"]',
          '[role="textbox"][aria-label*="Formula"]',
          'div[contenteditable="true"][aria-label*="Formula"]',
          '[data-automation-id*="formulaBar"]',
          '[data-testid*="formulaBar"]'
        ];
        const deepQuery = (selectorList) => {
          const roots = [document];
          const visited = new Set();
          while (roots.length) {
            const root = roots.shift();
            if (!root || visited.has(root)) {
              continue;
            }
            visited.add(root);
            for (const selector of selectorList) {
              const hits = root.querySelectorAll(selector);
              for (const hit of hits) {
                if (hit) {
                  return hit;
                }
              }
            }
            const nodes = root.querySelectorAll ? root.querySelectorAll("*") : [];
            nodes.forEach((node) => {
              if (node.shadowRoot) {
                roots.push(node.shadowRoot);
              }
            });
          }
          return null;
        };
        const hasNameBox = Boolean(deepQuery(selectors));
        const hasFormula = Boolean(deepQuery(formulaSelectors));
        const hasGrid = Boolean(document.querySelector('[role="grid"]'));
        return hasNameBox || hasFormula || hasGrid;
      }
    });
    return results.some((entry) => entry.result === true);
  } catch (error) {
    return false;
  }
};

const fallbackGetExcelSelection = async (tabId) => {
  try {
    const results = await chrome.scripting.executeScript({
      target: { tabId, allFrames: true },
      func: async () => {
        const selectors = [
          'input[aria-label="Name box"]',
          'input[aria-label="Name Box"]',
          'input[aria-label*="Name"]',
          'input[aria-label*="Nombre"]',
          'div[contenteditable="true"][aria-label*="Name"]',
          'div[contenteditable="true"][aria-label*="Nombre"]',
          'input[aria-label="Cuadro de nombre"]',
          'input[aria-label="Cuadro de nombres"]',
          'input[title*="Name"]',
          'input[title*="Nombre"]',
          '[data-automation-id*="nameBox"] input',
          '[data-automation-id*="NameBox"] input',
          '[data-automation-id="formulaNameBox"]',
          '#NameBox input',
          '#NameBox',
          '[role="combobox"][aria-label*="Name"]',
          '[role="combobox"][aria-label*="Nombre"]',
          '[role="textbox"][aria-label*="Name"]',
          '[role="textbox"][aria-label*="Nombre"]',
          '[data-testid*="nameBox"]'
        ];
        const formulaSelectors = [
          'input[aria-label="Formula Bar"]',
          'input[aria-label="Barra de formulas"]',
          'input[aria-label*="Formula"]',
          'textarea[aria-label*="Formula"]',
          '[role="textbox"][aria-label*="Formula"]',
          'div[contenteditable="true"][aria-label*="Formula"]',
          '[data-automation-id*="formulaBar"] input',
          '[data-automation-id*="formulaBar"] textarea',
          '[data-testid*="formulaBar"]',
          'input[aria-label]',
          'textarea[aria-label]',
          '[role="textbox"][aria-label]',
          'div[contenteditable="true"][aria-label]'
        ];

        const getValue = (el) =>
          el && (el.value || el.getAttribute("value") || el.textContent || "");

        const buildLabelText = (el) => {
          const parts = [];
          const nodes = [el, el ? el.parentElement : null];
          nodes.forEach((node) => {
            if (!node) {
              return;
            }
            parts.push(
              node.getAttribute("aria-label"),
              node.getAttribute("title"),
              node.getAttribute("data-automation-id"),
              node.getAttribute("data-testid"),
              node.getAttribute("data-test-id"),
              node.id
            );
          });
          return parts
            .filter(Boolean)
            .join(" ")
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .toLowerCase();
        };
        const isLikelyNameBox = (el) => {
          if (!el) {
            return false;
          }
          const label = buildLabelText(el);
          if (!label) {
            return false;
          }
          if (label.includes("font") || label.includes("fuente")) {
            return false;
          }
          return (
            label.includes("name box") ||
            label.includes("namebox") ||
            label.includes("cuadro de nombre") ||
            label.includes("cuadro de nombres") ||
            label.includes("formulanamebox")
          );
        };
        const isLikelyFormulaBar = (el) => {
          if (!el) {
            return false;
          }
          const label = buildLabelText(el);
          if (!label) {
            return false;
          }
          return (
            label.includes("formula bar") ||
            label.includes("barra de formulas") ||
            label.includes("formulabar")
          );
        };
        const findDeep = (selectorList, predicate) => {
          const roots = [document];
          const visited = new Set();
          while (roots.length) {
            const root = roots.shift();
            if (!root || visited.has(root)) {
              continue;
            }
            visited.add(root);
            for (const selector of selectorList) {
              const hits = root.querySelectorAll(selector);
              for (const hit of hits) {
                if (hit && (!predicate || predicate(hit))) {
                  return hit;
                }
              }
            }
            const nodes = root.querySelectorAll ? root.querySelectorAll("*") : [];
            nodes.forEach((node) => {
              if (node.shadowRoot) {
                roots.push(node.shadowRoot);
              }
            });
          }
          return null;
        };

        const GRID_SELECTORS = [
          '[role="grid"]',
          '[data-automation-id="grid"]',
          '[data-automation-id*="grid"]',
          '[data-testid*="grid"]',
          '[data-test-id*="grid"]'
        ];
        const findGrid = () => findDeep(GRID_SELECTORS, () => true);
        const nameBox = findDeep(selectors, isLikelyNameBox);
        const formulaBar = findDeep(formulaSelectors, isLikelyFormulaBar);
        const grid = findGrid();
        const hasGrid = Boolean(grid);
        const hasUi = Boolean(nameBox || formulaBar || hasGrid);
        const normalizeToken = (value) => {
          const raw = String(value || "").trim();
          if (!raw) {
            return "";
          }
          const noAbs = raw.replace(/\$/g, "");
          const bangIndex = noAbs.lastIndexOf("!");
          return bangIndex >= 0 ? noAbs.slice(bangIndex + 1).trim() : noAbs;
        };
        const isRangeRef = (value) => {
          const cleaned = normalizeToken(value).toUpperCase();
          if (!cleaned) {
            return false;
          }
          if (/^[A-Z]{1,3}\d{1,7}$/.test(cleaned)) {
            return true;
          }
          if (/^[A-Z]{1,3}\d{1,7}:[A-Z]{1,3}\d{1,7}$/.test(cleaned)) {
            return true;
          }
          if (/^[A-Z]{1,3}:[A-Z]{1,3}$/.test(cleaned)) {
            return true;
          }
          if (/^\d+:\d+$/.test(cleaned)) {
            return true;
          }
          return false;
        };
        const columnIndexToLetter = (index) => {
          let current = Number(index);
          let result = "";
          while (current > 0) {
            const mod = (current - 1) % 26;
            result = String.fromCharCode(65 + mod) + result;
            current = Math.floor((current - mod - 1) / 26);
          }
          return result;
        };
        const extractCellRefFromLabel = (label) => {
          const match = String(label || "").match(/\b([A-Z]{1,3}\d{1,7})\b/);
          return match ? match[1] : "";
        };
        const cellRefFromElement = (cell) => {
          if (!cell) {
            return "";
          }
          const colIndex = Number.parseInt(
            cell.getAttribute("aria-colindex"),
            10
          );
          const rowIndex = Number.parseInt(
            cell.getAttribute("aria-rowindex"),
            10
          );
          if (Number.isFinite(colIndex) && Number.isFinite(rowIndex)) {
            return `${columnIndexToLetter(colIndex)}${rowIndex}`;
          }
          return extractCellRefFromLabel(cell.getAttribute("aria-label") || "");
        };
        const getActiveCellFromGrid = () => {
          const container =
            grid || findDeep(['[aria-activedescendant]'], () => true);
          if (!container) {
            return "";
          }
          const activeId = container.getAttribute("aria-activedescendant");
          if (activeId) {
            const activeCell = document.getElementById(activeId);
            const ref = cellRefFromElement(activeCell);
            if (ref) {
              return ref;
            }
          }
          const selectedCell =
            container.querySelector('[role="gridcell"][aria-selected="true"]') ||
            container.querySelector('[role="gridcell"][aria-selected="mixed"]') ||
            container.querySelector('[role="gridcell"][tabindex="0"]');
          const selectedRef = cellRefFromElement(selectedCell);
          if (selectedRef) {
            return selectedRef;
          }
          return extractCellRefFromLabel(
            container.getAttribute("aria-label") || ""
          );
        };
        let range = normalizeToken(getValue(nameBox));
        if (!isRangeRef(range)) {
          range = "";
        }
        if (!range) {
          range = getActiveCellFromGrid();
        }
        if (!range) {
          const active = document.activeElement;
          const label = active ? active.getAttribute("aria-label") || "" : "";
          range = extractCellRefFromLabel(label);
        }
        const activeValue = String(getValue(formulaBar)).trim();
        let copyOk = false;
        try {
          if (grid && grid.focus) {
            grid.focus();
          }
          copyOk = document.execCommand("copy");
        } catch (error) {
          copyOk = false;
        }
        let clipboardText = "";
        try {
          clipboardText = await navigator.clipboard.readText();
        } catch (error) {
          clipboardText = "";
        }
        const cachedRange = window.__GA_LAST_RANGE__ || "";
        const cachedCell = window.__GA_LAST_CELL__ || "";
        const cachedClipboard = window.__GA_LAST_CLIPBOARD__ || "";
        return {
          range,
          activeValue,
          copyOk,
          clipboardText,
          hasUi,
          hasGrid,
          cachedRange,
          cachedCell,
          cachedClipboard,
          frameUrl: location.href
        };
      }
    });
    const pickBest = (entries) => {
      let best = null;
      let bestScore = -1;
      entries.forEach((entry) => {
        if (!entry || !entry.result) {
          return;
        }
        const result = entry.result;
        const rangeCandidate =
          result.range || result.cachedRange || result.cachedCell || "";
        const score =
          (isRangeRef(rangeCandidate) ? 6 : 0) +
          (result.hasGrid ? 3 : 0) +
          (result.clipboardText ? 1 : 0) +
          (result.hasUi ? 1 : 0);
        if (score > bestScore) {
          bestScore = score;
          best = entry;
        }
      });
      return best;
    };
    const hit = pickBest(results);
    if (!hit || !hit.result) {
      return {};
    }
    return {
      ...hit.result,
      frameId: hit.frameId
    };
  } catch (error) {
    return {};
  }
};

const fallbackPasteExcelTsv = async (tabId, tsv) => {
  try {
    const results = await chrome.scripting.executeScript({
      target: { tabId, allFrames: true },
      args: [tsv],
      func: (payload) => {
        let clipboardOk = false;
        let pasteOk = false;
        try {
          navigator.clipboard.writeText(payload);
          clipboardOk = true;
        } catch (error) {
          clipboardOk = false;
        }
        try {
          pasteOk = document.execCommand("paste");
        } catch (error) {
          pasteOk = false;
        }
        return { clipboardOk, pasteOk };
      }
    });
    const hit = results.find((entry) => entry.result && entry.result.pasteOk);
    if (hit && hit.result) {
      return { ...hit.result };
    }
    return { clipboardOk: false, pasteOk: false };
  } catch (error) {
    return { clipboardOk: false, pasteOk: false };
  }
};

const isTabOpen = (tabs, targetUrl) => {
  if (!targetUrl) {
    return false;
  }
  const normalized = normalizeUrl(targetUrl);
  const signature = extractDocSignature(targetUrl);
  return tabs.some((tab) => {
    if (!tab.url) {
      return false;
    }
    if (normalized && tab.url.includes(normalized)) {
      return true;
    }
    if (signature && tab.url.includes(signature)) {
      return true;
    }
    if (!normalized && tab.url.includes(targetUrl)) {
      return true;
    }
    return false;
  });
};

const findExcelTab = async (targetUrl) => {
  const excelTabs = await safeTabsQuery({
    url: EXCEL_URL_PATTERNS
  });
  if (!targetUrl) {
    return excelTabs.find((tab) => tab.active) || excelTabs[0] || null;
  }
  const normalized = normalizeUrl(targetUrl);
  const directMatch =
    excelTabs.find((tab) => tab.url && tab.url.includes(normalized)) ||
    excelTabs.find((tab) => tab.url && tab.url.includes(targetUrl));
  if (directMatch) {
    return directMatch;
  }
  const signature = extractDocSignature(targetUrl);
  if (signature) {
    const signatureMatch = excelTabs.find(
      (tab) => tab.url && tab.url.includes(signature)
    );
    if (signatureMatch) {
      return signatureMatch;
    }
  }
  const [activeTab] = await chrome.tabs.query({
    active: true,
    currentWindow: true
  });
  if (activeTab) {
    if (await probeExcelUi(activeTab.id)) {
      logEvent("info", "Pestana activa detectada como Excel Online.");
      return activeTab;
    }
    if (isLikelyExcelUrl(activeTab.url || "")) {
      logEvent("info", "Usando pestana activa con URL de Excel Online.");
      return activeTab;
    }
  }
  for (const tab of excelTabs) {
    if (await probeExcelUi(tab.id)) {
      logEvent("info", "Pestana Excel detectada por UI.");
      return tab;
    }
  }
  const allTabs = await chrome.tabs.query({ currentWindow: true });
  const candidates = allTabs.filter((tab) => isLikelyExcelUrl(tab.url || ""));
  for (const tab of candidates) {
    if (await probeExcelUi(tab.id)) {
      logEvent("info", "Pestana Excel detectada por coincidencia de UI.");
      return tab;
    }
  }
  return null;
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

const extractColumnFromRange = (range) => {
  const normalized = normalizeRangeToken(range);
  if (!normalized || !isRangeRef(normalized)) {
    return "";
  }
  const cleaned = String(normalized).toUpperCase().split(":")[0];
  const match = cleaned.match(/^([A-Z]+)\d*/);
  return match ? match[1] : "";
};

const parseCellRef = (value) => {
  const cleaned = normalizeRangeToken(value).toUpperCase().split(":")[0];
  if (!cleaned) {
    return null;
  }
  const match = cleaned.match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    return null;
  }
  return {
    column: match[1],
    row: Number.parseInt(match[2], 10)
  };
};

const parseRangeBounds = (address) => {
  const token = normalizeRangeToken(address).toUpperCase().replace(/\$/g, "");
  if (!token) {
    return null;
  }
  const parts = token.split(":");
  const start = parts[0];
  const end = parts[1] || parts[0];
  const startMatch = start.match(/^([A-Z]+)(\d+)?$/);
  const endMatch = end.match(/^([A-Z]+)(\d+)?$/);
  if (!startMatch || !endMatch) {
    return null;
  }
  return {
    startColumn: startMatch[1],
    endColumn: endMatch[1],
    startRow: startMatch[2] ? Number.parseInt(startMatch[2], 10) : null,
    endRow: endMatch[2] ? Number.parseInt(endMatch[2], 10) : null
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

const parseGuidesFromColumn = (values, mode) => {
  const cleanedValues = values.map((value) => String(value || "").trim());
  const bolsaLines = cleanedValues.filter((line) => /bolsa/i.test(line));
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

  for (const rawValue of cleanedValues) {
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

const parseGuidesFromClipboard = (text, mode) => {
  const rows = text.replace(/\r/g, "").split("\n");
  const parsedRows = rows.map((row) => row.split("\t"));
  const columnCount = parsedRows.reduce(
    (max, row) => Math.max(max, row.length),
    0
  );
  if (columnCount <= 1) {
    const columnValues = parsedRows.map((row) => row[0] || "");
    const result = parseGuidesFromColumn(columnValues, mode);
    return { ...result, columnOffset: 0 };
  }

  let best = null;
  for (let col = 0; col < columnCount; col += 1) {
    const columnValues = parsedRows.map((row) => row[col] || "");
    const result = parseGuidesFromColumn(columnValues, mode);
    if (!best) {
      best = { ...result, columnOffset: col };
      continue;
    }
    if (result.guides.length > best.guides.length) {
      best = { ...result, columnOffset: col };
      continue;
    }
    if (
      result.guides.length === best.guides.length &&
      result.invalidCount < best.invalidCount
    ) {
      best = { ...result, columnOffset: col };
    }
  }

  if (!best) {
    return {
      guides: [],
      invalidCount: 0,
      duplicateCount: 0,
      selectionType: mode === "bolsa" ? "bolsa" : "columna",
      bagLabel: "",
      columnOffset: 0
    };
  }

  return best;
};

const extractColumnValuesFromRange = (rangeData, column) => {
  if (!rangeData || !Array.isArray(rangeData.values)) {
    return { values: [], startRow: 1 };
  }
  const bounds = parseRangeBounds(rangeData.address || "");
  if (!bounds) {
    const values = rangeData.values.map((row) => row[0] || "");
    return { values, startRow: 1 };
  }
  const startRow = bounds.startRow || 1;
  if (!column) {
    const values = rangeData.values.map((row) => row[0] || "");
    return { values, startRow };
  }
  const startColIndex = columnLetterToIndex(bounds.startColumn);
  const targetIndex = columnLetterToIndex(column);
  const offset = targetIndex - startColIndex;
  if (offset < 0) {
    return { values: [], startRow };
  }
  const values = rangeData.values.map((row) => row[offset] || "");
  return { values, startRow };
};

const loadColumnValuesFromGraph = async (docUrl, sheetName, column) => {
  const shareId = buildShareId(docUrl);
  const sheet = await graphGetWorksheet(shareId, sheetName);
  try {
    const range = await graphGetRange(shareId, sheet.id, `${column}:${column}`);
    const { values, startRow } = extractColumnValuesFromRange(range, column);
    if (values.length) {
      return { values, startRow, sheet };
    }
  } catch (error) {
    logEvent("info", `Graph: fallo lectura directa de columna. ${error.message || error}`);
  }

  const used = await graphGetUsedRange(shareId, sheet.id);
  const { values, startRow } = extractColumnValuesFromRange(used, column);
  return { values, startRow, sheet };
};

const parseGuidesFromColumnValues = (values, startRow, selectionBounds, mode) => {
  const rows = values.map((value, index) => ({
    row: startRow + index,
    value
  }));
  if (!rows.length) {
    return {
      guides: [],
      invalidCount: 0,
      duplicateCount: 0,
      selectionType: mode === "infer" ? "columna" : mode,
      bagLabel: ""
    };
  }
  const bagRows = rows.filter((row) => /bolsa/i.test(String(row.value || "")));
  const hasRowBounds =
    selectionBounds && selectionBounds.startRow && selectionBounds.endRow;
  const rowSpan = hasRowBounds
    ? Math.abs(selectionBounds.endRow - selectionBounds.startRow) + 1
    : null;
  const bagHeadersInSelection = hasRowBounds
    ? bagRows.filter(
        (row) =>
          row.row >= selectionBounds.startRow &&
          row.row <= selectionBounds.endRow
      )
    : [];

  let selectionType = mode === "infer" ? "columna" : mode;
  if (mode === "infer") {
    if (selectionBounds && !selectionBounds.startRow) {
      selectionType = "columna";
    } else if (bagHeadersInSelection.length > 1 || (rowSpan && rowSpan > 40)) {
      selectionType = "columna";
    } else {
      selectionType = "bolsa";
    }
  }
  if (selectionType === "bolsa" && bagRows.length === 0) {
    selectionType = "columna";
  }

  const guides = [];
  const invalid = [];
  const seen = new Set();
  let duplicateCount = 0;
  let bagLabel = "";
  let emptyRun = 0;

  const collectRange = (startIndex, endIndex) => {
    emptyRun = 0;
    for (let idx = startIndex; idx <= endIndex; idx += 1) {
      const rawValue = rows[idx] ? rows[idx].value : "";
      const value = String(rawValue || "").trim();
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
  };

  if (selectionType === "columna") {
    const firstBag = bagRows.length ? bagRows[0].row : rows[0]?.row || startRow;
    if (bagRows.length && !bagLabel) {
      bagLabel = String(bagRows[0].value || "").trim();
    }
    const startIndex = rows.findIndex((row) => row.row === firstBag) + 1;
    const safeStart = startIndex > 0 ? startIndex : 0;
    collectRange(safeStart, rows.length - 1);
  } else {
    let anchorRow = selectionBounds ? selectionBounds.startRow : null;
    if (!anchorRow && bagRows.length) {
      anchorRow = bagRows[0].row;
    }
    if (!anchorRow) {
      collectRange(0, rows.length - 1);
    } else {
      const prevBag = bagRows
        .filter((row) => row.row <= anchorRow)
        .slice(-1)[0];
      const nextBag = bagRows.find((row) => row.row > anchorRow);
      const bagStartRow = prevBag ? prevBag.row : bagRows[0]?.row || anchorRow;
      const bagEndRow = nextBag ? nextBag.row - 1 : rows[rows.length - 1]?.row;
      const startIndex = rows.findIndex((row) => row.row === bagStartRow) + 1;
      const endHit = rows.findIndex((row) => row.row === bagEndRow);
      const endIndex = endHit >= 0 ? endHit : rows.length - 1;
      const bagHeader = rows.find((row) => row.row === bagStartRow);
      if (bagHeader && bagHeader.value && !bagLabel) {
        bagLabel = String(bagHeader.value).trim();
      }
      collectRange(startIndex > 0 ? startIndex : 0, endIndex);
    }
  }

  return {
    guides,
    invalidCount: invalid.length,
    duplicateCount,
    selectionType,
    bagLabel
  };
};

const parseDoc2ExistingFromValues = (values) => {
  const guides = new Set();
  const asinsByGuide = {};
  let maxIndex = 0;

  values.forEach((row) => {
    const cols = row.map((col) => String(col || "").trim());
    if (!cols.length || cols.every((col) => !col)) {
      return;
    }
    const lineText = cols.join(" ").toLowerCase();
    if (lineText.includes("asin") || lineText.includes("producto")) {
      return;
    }
    const indexValue = Number.parseInt(cols[0], 10);
    if (Number.isFinite(indexValue) && indexValue > maxIndex) {
      maxIndex = indexValue;
    }
    const asin = cols[1] || "";
    const guideRaw = cols[3] || "";
    const guide = normalizeGuideValue(guideRaw);
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
    const hasMultipleItems =
      Array.isArray(result.items) && result.items.length > 1;
    const hasQtyOverOne = entries.some(([, qty]) => qty > 1);
    if (hasMultipleAsins || hasMultipleItems || hasQtyOverOne) {
      const reason =
        hasMultipleAsins || hasMultipleItems
          ? "Varios productos por guia"
          : "Cantidad mayor a 1";
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
    url: EXCEL_URL_PATTERNS
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
  const range = normalizeRangeToken(currentSettings.doc1Selection.range);
  document.getElementById("doc1Range").textContent =
    isRangeRef(range) ? range : "-";
  document.getElementById("doc1Type").textContent =
    currentSettings.doc1Selection.selectionType || "-";
  document.getElementById("doc1Count").textContent =
    currentSettings.doc1Selection.guides.length || 0;
  document.getElementById("doc1InvalidCount").textContent =
    currentSettings.doc1Selection.invalidCount || 0;
  document.getElementById("doc1Mode").value = currentSettings.doc1Selection.mode || "infer";
  document.getElementById("doc1ManualCol").value =
    currentSettings.doc1Selection.manualColumn || "";
};

const updateDoc2Summary = () => {
  const startCell = normalizeRangeToken(currentSettings.doc2Selection.startCell);
  const validCell = /^[A-Z]{1,3}\d{1,7}$/.test(startCell);
  document.getElementById("doc2StartCell").textContent = validCell
    ? startCell
    : "-";
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
  document.getElementById("sheetDoc1").value = currentSettings.sheetDoc1 || "";
  document.getElementById("sheetPrimary").value = currentSettings.sheetPrimary || "";
  document.getElementById("sheetDuplicates").value =
    currentSettings.sheetDuplicates || "Duplicados";
  document.getElementById("msClientId").value = currentSettings.msAuth.clientId || "";
  document.getElementById("msTenant").value = currentSettings.msAuth.tenant || "common";

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
  updateMicrosoftStatus();
};

const saveConfig = async () => {
  const prevDoc1Url = currentSettings.excelDoc1Url;
  const prevDoc2Url = currentSettings.excelDoc2Url;
  const prevSheetDoc1 = currentSettings.sheetDoc1;
  const prevClientId = currentSettings.msAuth.clientId;
  const prevTenant = currentSettings.msAuth.tenant;
  currentSettings.excelDoc1Url = document.getElementById("doc1Url").value.trim();
  currentSettings.excelDoc2Url = document.getElementById("doc2Url").value.trim();
  currentSettings.sheetDoc1 = document.getElementById("sheetDoc1").value.trim();
  currentSettings.sheetPrimary = document.getElementById("sheetPrimary").value.trim();
  currentSettings.sheetDuplicates = document
    .getElementById("sheetDuplicates")
    .value.trim();
  currentSettings.msAuth.clientId = document.getElementById("msClientId").value.trim();
  currentSettings.msAuth.tenant =
    document.getElementById("msTenant").value.trim() || "common";

  if (prevDoc1Url && prevDoc1Url !== currentSettings.excelDoc1Url) {
    currentSettings.doc1Selection = { ...DEFAULT_SETTINGS.doc1Selection };
  }
  if (prevSheetDoc1 && prevSheetDoc1 !== currentSettings.sheetDoc1) {
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
  if (
    prevClientId &&
    (prevClientId !== currentSettings.msAuth.clientId ||
      prevTenant !== currentSettings.msAuth.tenant)
  ) {
    currentSettings.msAuth.accessToken = "";
    currentSettings.msAuth.refreshToken = "";
    currentSettings.msAuth.expiresAt = 0;
  }
  await chrome.storage.local.set({
    excelDoc1Url: currentSettings.excelDoc1Url,
    excelDoc2Url: currentSettings.excelDoc2Url,
    sheetDoc1: currentSettings.sheetDoc1,
    sheetPrimary: currentSettings.sheetPrimary,
    sheetDuplicates: currentSettings.sheetDuplicates,
    msAuth: currentSettings.msAuth,
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

  const ready = await ensureExcelReady(tab.id);
  if (!ready) {
    logEvent("error", "Excel Online no esta listo para capturar.");
    await setLastError("Excel Online no esta listo.");
    return;
  }

  const fallback = await fallbackGetExcelSelection(tab.id);
  const response = await getExcelSelectionRange(tab.id, fallback.frameId);
  let range =
    (response && response.range) ||
    fallback.range ||
    fallback.cachedRange ||
    fallback.cachedCell ||
    "";
  let column = (response && response.column) || extractColumnFromRange(range);
  const manualColumn = String(currentSettings.doc1Selection.manualColumn || "").trim();
  if (manualColumn) {
    const cleaned = manualColumn.toUpperCase().replace(/[^A-Z]/g, "");
    if (cleaned) {
      column = cleaned;
    }
  }
  range = normalizeRangeToken(range);
  if (!isRangeRef(range)) {
    range = "";
    if (!manualColumn) {
      column = "";
    }
  }
  if (!range && fallback.hasUi && !manualColumn) {
    logEvent(
      "error",
      "Excel detectado, pero no se pudo leer la celda activa. Haz clic en una celda y vuelve a intentar."
    );
  }

  if (!column) {
    logEvent("error", "No se pudo detectar la columna en Documento 1.");
    await setLastError("Columna no detectada.");
    return;
  }

  let valuesData = null;
  const selectionBounds = parseRangeBounds(range);
  const canUseGraph =
    Boolean(currentSettings.msAuth.clientId) &&
    !isPersonalTenant(currentSettings.msAuth.tenant);

  if (canUseGraph) {
    let token = await getMicrosoftAccessToken();
    if (!token) {
      const connected = await startMicrosoftAuth();
      if (connected) {
        token = await getMicrosoftAccessToken();
      }
    }
    if (token) {
      try {
        valuesData = await loadColumnValuesFromGraph(
          currentSettings.excelDoc1Url,
          currentSettings.sheetDoc1,
          column
        );
        logEvent(
          "info",
          `Graph: columna ${column} con ${valuesData.values.length} filas leidas.`
        );
      } catch (error) {
        if (isGraphMsaError(error)) {
          logEvent(
            "info",
            "Microsoft Graph no soporta cuentas personales. Usando lectura por UI."
          );
        } else {
          logEvent(
            "error",
            `No se pudo leer Excel via Microsoft Graph: ${error.message || error}`
          );
        }
      }
    }
  } else if (!currentSettings.msAuth.clientId) {
    logEvent("info", "Microsoft Graph no configurado. Usando lectura por UI.");
  } else if (isPersonalTenant(currentSettings.msAuth.tenant)) {
    logEvent("info", "Cuenta personal detectada. Usando lectura por UI.");
  }

  if (!valuesData) {
    const startRow = selectionBounds && selectionBounds.startRow
      ? Math.max(1, selectionBounds.startRow - 20)
      : 1;
    const endRow = selectionBounds && selectionBounds.endRow
      ? selectionBounds.endRow
      : 0;
    const uiResult = await readColumnValuesFromUi(tab.id, fallback.frameId, {
      column,
      startRow,
      endRow,
      maxRows: 1200,
      stopAfterEmptyRun: 3
    });
    if (!uiResult || !uiResult.ok) {
      logEvent("error", "No se pudo leer la columna desde Excel Online.");
      await setLastError("No se pudo leer la columna en Excel Online.");
      return;
    }
    valuesData = {
      values: uiResult.values || [],
      startRow: uiResult.startRow || startRow
    };
    logEvent(
      "info",
      `UI: columna ${column} con ${valuesData.values.length} filas leidas.`
    );
  }

  const parsed = parseGuidesFromColumnValues(
    valuesData.values || [],
    valuesData.startRow || 1,
    selectionBounds,
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
      "No se detectaron guias. Verifica que la columna tenga datos."
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

  const ready = await ensureExcelReady(tab.id);
  if (!ready) {
    logEvent("error", "Excel Online no esta listo para capturar.");
    await setLastError("Excel Online no esta listo.");
    return;
  }

  const fallback = await fallbackGetExcelSelection(tab.id);
  const response = await getExcelActiveCell(tab.id, fallback.frameId);

  let activeCell = response ? response.activeCell : "";
  let activeValue = response ? response.activeValue || "" : "";
  if (!activeCell) {
    activeCell = fallback.range || fallback.cachedRange || fallback.cachedCell || "";
    activeValue = fallback.activeValue || "";
  }
  const normalizedCell = normalizeRangeToken(activeCell);
  const parsed = parseCellRef(normalizedCell);
  if (!parsed) {
    logEvent(
      "error",
      `Debug Excel: range=${fallback.range || "-"} cached=${fallback.cachedRange || "-"} cell=${fallback.cachedCell || "-"} ui=${fallback.hasUi ? "si" : "no"} grid=${fallback.hasGrid ? "si" : "no"}`
    );
    logEvent(
      "error",
      "No se pudo leer la celda. Haz clic en la celda inicial y presiona Capturar."
    );
    if (fallback.hasUi) {
      logEvent(
        "error",
        "Excel detectado, pero no se pudo leer la celda activa. Confirma que la celda esta seleccionada."
      );
    }
    await setLastError("Seleccion invalida en Documento 2.");
    return;
  }

  const baseIndex = columnLetterToIndex(parsed.column);
  const manualColumn = columnIndexToLetter(baseIndex + 7);
  const headerPresent = /no\.?\s*producto/i.test(activeValue);

  currentSettings.doc2Selection = {
    ...currentSettings.doc2Selection,
    startCell: normalizedCell || activeCell,
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
  if (!currentSettings.doc2Selection.startColumn) {
    logEvent("error", "Captura la celda inicial del Documento 2 primero.");
    await setLastError("Documento 2 sin celda inicial.");
    return;
  }
  const tab = await findExcelTab(currentSettings.excelDoc2Url);
  if (!tab) {
    logEvent("error", "Documento 2 no esta abierto en Excel Online.");
    await setLastError("Documento 2 no esta abierto.");
    return;
  }
  const ready = await ensureExcelReady(tab.id);
  if (!ready) {
    logEvent("error", "Excel Online no esta listo para capturar.");
    await setLastError("Excel Online no esta listo.");
    return;
  }

  let usedRange = null;
  const canUseGraph =
    Boolean(currentSettings.msAuth.clientId) &&
    !isPersonalTenant(currentSettings.msAuth.tenant);
  if (canUseGraph) {
    let token = await getMicrosoftAccessToken();
    if (!token) {
      const connected = await startMicrosoftAuth();
      if (connected) {
        token = await getMicrosoftAccessToken();
      }
    }
    if (token) {
      try {
        const shareId = buildShareId(currentSettings.excelDoc2Url);
        const sheet = await graphGetWorksheet(shareId, currentSettings.sheetPrimary);
        usedRange = await graphGetUsedRange(shareId, sheet.id);
      } catch (error) {
        if (isGraphMsaError(error)) {
          logEvent(
            "info",
            "Microsoft Graph no soporta cuentas personales. Usando lectura por UI."
          );
        } else {
          logEvent(
            "error",
            `No se pudo leer Documento 2 con Graph: ${error.message || error}`
          );
        }
      }
    }
  } else if (!currentSettings.msAuth.clientId) {
    logEvent("info", "Microsoft Graph no configurado. Usando lectura por UI.");
  } else if (isPersonalTenant(currentSettings.msAuth.tenant)) {
    logEvent("info", "Cuenta personal detectada. Usando lectura por UI.");
  }

  let rows = [];
  if (usedRange) {
    const bounds = parseRangeBounds(usedRange.address || "");
    if (!bounds || !Array.isArray(usedRange.values)) {
      logEvent("error", "Documento 2 sin datos detectados.");
      await setLastError("Documento 2 sin datos.");
      return;
    }
    const usedStartIndex = columnLetterToIndex(bounds.startColumn);
    const targetIndex = columnLetterToIndex(currentSettings.doc2Selection.startColumn);
    const offset = targetIndex - usedStartIndex;
    rows = usedRange.values.map((row) => {
      const rowData = [];
      for (let i = 0; i < 4; i += 1) {
        const value = row[offset + i];
        rowData.push(value === undefined ? "" : value);
      }
      return rowData;
    });
  } else {
    const fallback = await fallbackGetExcelSelection(tab.id);
    const frameId = Number.isInteger(fallback.frameId) ? fallback.frameId : null;
    const startColumn = currentSettings.doc2Selection.startColumn;
    const startRow = currentSettings.doc2Selection.startRow || 1;
    const guideColumn = columnIndexToLetter(
      columnLetterToIndex(startColumn) + 3
    );
    const guideResult = await readColumnValuesFromUi(tab.id, frameId, {
      column: guideColumn,
      startRow,
      maxRows: 1200,
      stopAfterEmptyRun: 3
    });
    if (!guideResult || !guideResult.ok) {
      logEvent("error", "No se pudo leer la columna de guia en Documento 2.");
      await setLastError("No se pudo leer Documento 2 por UI.");
      return;
    }
    const guideValues = guideResult.values || [];
    let lastIndex = -1;
    guideValues.forEach((value, index) => {
      if (String(value || "").trim()) {
        lastIndex = index;
      }
    });
    const totalRows = lastIndex >= 0 ? lastIndex + 1 : 0;
    if (!totalRows) {
      rows = [];
    } else {
      const endRow = startRow + totalRows - 1;
      const indexColumn = startColumn;
      const asinColumn = columnIndexToLetter(
        columnLetterToIndex(startColumn) + 1
      );
      const indexResult = await readColumnValuesFromUi(tab.id, frameId, {
        column: indexColumn,
        startRow,
        endRow,
        maxRows: totalRows
      });
      const asinResult = await readColumnValuesFromUi(tab.id, frameId, {
        column: asinColumn,
        startRow,
        endRow,
        maxRows: totalRows
      });
      const indexValues = indexResult && indexResult.ok ? indexResult.values || [] : [];
      const asinValues = asinResult && asinResult.ok ? asinResult.values || [] : [];
      rows = new Array(totalRows).fill(null).map((_, idx) => [
        indexValues[idx] || "",
        asinValues[idx] || "",
        "",
        guideValues[idx] || ""
      ]);
    }
    logEvent(
      "info",
      `UI: Documento 2 leido con ${rows.length} filas detectadas.`
    );
  }

  const parsed = parseDoc2ExistingFromValues(rows);
    currentSettings.dedupe = {
      ...currentSettings.dedupe,
      doc2Guides: parsed.guides,
      doc2AsinsByGuide: parsed.asinsByGuide
    };
  if (parsed.guides.length || parsed.maxIndex) {
    currentSettings.doc2Selection.headersEnsured = true;
  }
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
  logEvent("info", `Iniciando busqueda en Amazon. Guias: ${guides.length}.`);
  const sellerTab = await findSellerTab();
  if (!sellerTab) {
    logEvent("error", "No hay pestana activa de Amazon Seller.");
    await setLastError("Amazon Seller no esta abierto.");
    return;
  }

  const ready = await ensureContentScript(sellerTab.id, "content/amazon_seller.js");
  if (!ready) {
    logEvent("error", "Amazon Seller no esta listo para recibir comandos.");
    await setLastError("Amazon Seller no esta listo.");
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
  if (!output.mainRows.length && !output.manualRows.length) {
    logEvent("error", "Amazon no devolvio items para procesar.");
    await setLastError("Amazon sin items.");
    return;
  }
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
  if (
    !currentSettings.pendingOutput.mainRows.length &&
    !currentSettings.pendingOutput.manualRows.length
  ) {
    logEvent("error", "No hay datos pendientes para escribir.");
    await setLastError("Sin datos pendientes para escribir.");
    return;
  }
  logEvent("info", "Preparando escritura en Excel Online.");
  if (!currentSettings.doc2Selection.startColumn) {
    logEvent("error", "No se detecto columna inicial en Documento 2.");
    await setLastError("Seleccion de Documento 2 incompleta.");
    return;
  }
  const tab = await findExcelTab(currentSettings.excelDoc2Url);
  if (!tab) {
    logEvent("error", "Documento 2 no esta abierto en Excel Online.");
    await setLastError("Documento 2 no esta abierto.");
    return;
  }
  const canUseGraph =
    Boolean(currentSettings.msAuth.clientId) &&
    !isPersonalTenant(currentSettings.msAuth.tenant);

  const startColumn = currentSettings.doc2Selection.startColumn;
  const startRow = currentSettings.doc2Selection.startRow || 1;
  const manualColumn =
    currentSettings.doc2Selection.manualColumn ||
    columnIndexToLetter(columnLetterToIndex(startColumn) + 7);
  const manualOffset =
    columnLetterToIndex(manualColumn) - columnLetterToIndex(startColumn);
  const width = manualOffset + 1;
  const headersNeeded = !currentSettings.doc2Selection.headersEnsured;

  const rows = [];
  if (headersNeeded) {
    const headerRow = new Array(width).fill("");
    headerRow[0] = "#No. Producto";
    headerRow[1] = "ASIN";
    headerRow[2] = "Cantidad";
    headerRow[3] = "No. Guia";
    headerRow[manualOffset] = "Revisar Manualmente";
    rows.push(headerRow);
  }

  currentSettings.pendingOutput.mainRows.forEach((row) => {
    const rowData = new Array(width).fill("");
    rowData[0] = String(row.index);
    rowData[1] = row.asin;
    rowData[2] = String(row.quantity);
    rowData[3] = row.guide;
    rows.push(rowData);
  });

  currentSettings.pendingOutput.manualRows.forEach((entry) => {
    const rowData = new Array(width).fill("");
    rowData[manualOffset] = formatManualEntry(entry);
    rows.push(rowData);
  });

  if (!rows.length) {
    logEvent("error", "No hay filas preparadas para escribir.");
    await setLastError("Sin filas para escribir.");
    return;
  }

  const nextIndex =
    currentSettings.pendingOutput.nextIndex ||
    currentSettings.doc2Selection.nextIndex ||
    1;
  const targetRow = headersNeeded ? startRow : startRow + nextIndex;
  const endColumn = columnIndexToLetter(
    columnLetterToIndex(startColumn) + width - 1
  );
  const endRow = targetRow + rows.length - 1;
  const address = `${startColumn}${targetRow}:${endColumn}${endRow}`;

  let wroteOk = false;
  if (canUseGraph) {
    let token = await getMicrosoftAccessToken();
    if (!token) {
      const connected = await startMicrosoftAuth();
      if (connected) {
        token = await getMicrosoftAccessToken();
      }
    }
    if (token) {
      try {
        const shareId = buildShareId(currentSettings.excelDoc2Url);
        const sheet = await graphGetWorksheet(shareId, currentSettings.sheetPrimary);
        await graphUpdateRange(shareId, sheet.id, address, rows);
        wroteOk = true;
        logEvent("info", "Resultados escritos en Excel via Graph.");
      } catch (error) {
        if (isGraphMsaError(error)) {
          logEvent(
            "info",
            "Microsoft Graph no soporta cuentas personales. Usando escritura por UI."
          );
        } else {
          logEvent(
            "error",
            `No se pudo escribir en Graph: ${error.message || error}`
          );
        }
      }
    }
  } else if (!currentSettings.msAuth.clientId) {
    logEvent("info", "Microsoft Graph no configurado. Usando escritura por UI.");
  } else if (isPersonalTenant(currentSettings.msAuth.tenant)) {
    logEvent("info", "Cuenta personal detectada. Usando escritura por UI.");
  }

  if (!wroteOk) {
    const ready = await ensureExcelReady(tab.id);
    if (!ready) {
      logEvent("error", "Excel Online no esta listo para escribir.");
      await setLastError("Excel Online no esta listo.");
      return;
    }
    const fallback = await fallbackGetExcelSelection(tab.id);
    const frameId = Number.isInteger(fallback.frameId) ? fallback.frameId : null;
    const selection = await setExcelSelection(tab.id, frameId, address);
    if (!selection || !selection.ok) {
      logEvent("error", "No se pudo seleccionar el rango en Excel.");
      await setLastError("No se pudo seleccionar el rango en Excel.");
      return;
    }
    const tsv = rows
      .map((row) => row.map((cell) => String(cell ?? "")).join("\t"))
      .join("\n");
    let pasteResponse = null;
    if (Number.isInteger(frameId)) {
      pasteResponse = await sendExcelMessage(tab.id, frameId, {
        type: "PASTE_TSV",
        tsv
      });
    }
    if (!pasteResponse) {
      pasteResponse = await safeSendMessage(tab.id, { type: "PASTE_TSV", tsv });
    }
    if (!pasteResponse || !pasteResponse.ok) {
      const fallbackPaste = await fallbackPasteExcelTsv(tab.id, tsv);
      if (!fallbackPaste || !fallbackPaste.pasteOk) {
        logEvent("error", "No se pudo pegar datos en Excel.");
        await setLastError("No se pudo pegar datos en Excel.");
        return;
      }
    }
    wroteOk = true;
    logEvent("info", "Resultados escritos en Excel via UI.");
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
    manualColumn,
    headersEnsured: true,
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
  logEvent("info", "Iniciando ejecucion del flujo.");
  await setLastError("");
  try {
    await refreshStatus();
    if (!currentSettings.doc1Selection.guides.length) {
      await captureDoc1Selection();
    }
    if (!currentSettings.doc1Selection.guides.length) {
      logEvent("error", "No se pudieron capturar guias para iniciar el flujo.");
      await setLastError("No se pudieron capturar guias.");
      return;
    }
    await runAmazonLookup();
    await writeExcelOutput();
    logEvent("info", "Flujo completado.");
  } catch (error) {
    logEvent("error", `Fallo en ejecucion: ${error.message || error}`);
    await setLastError("Fallo en ejecucion.");
  }
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
  document.getElementById("msConnect").addEventListener("click", () => {
    currentSettings.msAuth.clientId =
      document.getElementById("msClientId").value.trim();
    currentSettings.msAuth.tenant =
      document.getElementById("msTenant").value.trim() || "common";
    chrome.storage.local.set({ msAuth: currentSettings.msAuth });
    startMicrosoftAuth();
  });
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
  document.getElementById("doc1ManualCol").addEventListener("input", (event) => {
    currentSettings.doc1Selection.manualColumn = event.target.value.trim().toUpperCase();
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
    if (changes.msAuth) {
      currentSettings.msAuth = changes.msAuth.newValue || currentSettings.msAuth;
      updateMicrosoftStatus();
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
