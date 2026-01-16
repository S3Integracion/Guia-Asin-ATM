(() => {
  const NAME_BOX_SELECTORS = [
    'input[aria-label="Name box"]',
    'input[aria-label="Name Box"]',
    'input[aria-label*="Name"]',
    'input[title*="Name"]',
    'input[aria-label="Cuadro de nombre"]',
    'input[aria-label="Cuadro de nombres"]',
    'input[aria-label*="Nombre"]',
    'input[title*="Nombre"]',
    'input[aria-label="Nombre"]',
    'input[data-automation-id="formulaNameBox"]',
    '[data-automation-id="nameBox"] input',
    '[data-automation-id="NameBox"] input',
    '[data-automation-id*="nameBox"] input',
    '[data-automation-id*="NameBox"] input',
    '#NameBox input',
    '#NameBox',
    '[role="combobox"][aria-label*="Name"]',
    '[role="combobox"][aria-label*="Nombre"]'
  ];

  const FORMULA_BAR_SELECTORS = [
    'input[aria-label="Formula Bar"]',
    'input[aria-label="Barra de formulas"]',
    'input[aria-label*="Formula"]',
    'textarea[aria-label*="Formula"]',
    '[role="textbox"][aria-label*="Formula"]',
    '[role="textbox"][aria-label*="formula"]',
    'textarea[aria-label="Formula Bar"]',
    'textarea[aria-label="Barra de formulas"]'
  ];

  const logEvent = (level, message) => {
    chrome.runtime.sendMessage({
      type: "LOG_EVENT",
      level,
      source: "excel_online",
      message
    });
  };

  const findNameBox = () => {
    for (const selector of NAME_BOX_SELECTORS) {
      const el = document.querySelector(selector);
      if (el) {
        return el;
      }
    }
    const byAutomation = document.querySelector('[data-automation-id*="nameBox"]');
    if (byAutomation) {
      return byAutomation.querySelector("input") || byAutomation;
    }
    return null;
  };

  const findFormulaBar = () => {
    for (const selector of FORMULA_BAR_SELECTORS) {
      const el = document.querySelector(selector);
      if (el) {
        return el;
      }
    }
    const byAutomation = document.querySelector('[data-automation-id*="formulaBar"]');
    if (byAutomation) {
      return byAutomation.querySelector("input, textarea") || byAutomation;
    }
    return null;
  };

  const getSelectionRange = () => {
    const nameBox = findNameBox();
    if (!nameBox) {
      return "";
    }
    const value =
      nameBox.value ||
      nameBox.getAttribute("value") ||
      nameBox.textContent ||
      "";
    return String(value).trim();
  };

  const getActiveCellValue = () => {
    const formulaBar = findFormulaBar();
    if (!formulaBar) {
      return "";
    }
    const value =
      formulaBar.value ||
      formulaBar.getAttribute("value") ||
      formulaBar.textContent ||
      "";
    return String(value);
  };

  const extractColumnFromRange = (range) => {
    if (!range) {
      return "";
    }
    const match = range.toUpperCase().match(/^([A-Z]+)\d+/);
    return match ? match[1] : "";
  };

  const tryReadClipboard = async () => {
    try {
      return await navigator.clipboard.readText();
    } catch (error) {
      return "";
    }
  };

  const tryCopySelection = async () => {
    try {
      document.execCommand("copy");
      return await tryReadClipboard();
    } catch (error) {
      return "";
    }
  };

  const hasExcelUi = () => !!findNameBox() || !!findFormulaBar();

  const waitForExcelUi = async (timeoutMs = 1500) => {
    if (hasExcelUi()) {
      return true;
    }
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      await new Promise((resolve) => setTimeout(resolve, 150));
      if (hasExcelUi()) {
        return true;
      }
    }
    return false;
  };

  const captureSelection = async () => {
    await waitForExcelUi();
    const range = getSelectionRange();
    const column = extractColumnFromRange(range);
    let clipboardText = await tryReadClipboard();
    if (!clipboardText) {
      clipboardText = await tryCopySelection();
    }
    return {
      range,
      column,
      clipboardText
    };
  };

  const pasteTsv = async (tsv) => {
    if (!tsv) {
      return { ok: false, reason: "empty" };
    }
    let clipboardOk = false;
    try {
      await navigator.clipboard.writeText(tsv);
      clipboardOk = true;
    } catch (error) {
      clipboardOk = false;
    }
    let pasteOk = false;
    try {
      pasteOk = document.execCommand("paste");
    } catch (error) {
      pasteOk = false;
    }
    return {
      ok: clipboardOk && pasteOk,
      clipboardOk,
      pasteOk
    };
  };

  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (!message || typeof message !== "object") {
      return;
    }

    if (message.type === "CAPTURE_SELECTION") {
    if (!hasExcelUi()) {
      return;
    }
    captureSelection().then(sendResponse);
    return true;
    }

    if (message.type === "GET_ACTIVE_CELL") {
      if (!hasExcelUi()) {
        return;
      }
      const activeCell = getSelectionRange();
      const activeValue = getActiveCellValue();
      sendResponse({ activeCell, activeValue });
      return;
    }

    if (message.type === "PASTE_TSV") {
      if (!hasExcelUi()) {
        return;
      }
      pasteTsv(message.tsv || "").then(sendResponse);
      return true;
    }
  });

  const ready = () => {
    if (hasExcelUi()) {
      logEvent("info", "Excel Online content script listo.");
    }
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", ready, { once: true });
  } else {
    ready();
  }
})();
