(() => {
  if (window.__GA_EXCEL_LOADED__) {
    return;
  }
  window.__GA_EXCEL_LOADED__ = true;
  const NAME_BOX_SELECTORS = [
    'input[aria-label="Name box"]',
    'input[aria-label="Name Box"]',
    'input[aria-label*="Name"]',
    'input[title*="Name"]',
    'div[contenteditable="true"][aria-label*="Name"]',
    'input[aria-label="Cuadro de nombre"]',
    'input[aria-label="Cuadro de nombres"]',
    'input[aria-label*="Nombre"]',
    'input[title*="Nombre"]',
    'div[contenteditable="true"][aria-label*="Nombre"]',
    'input[aria-label="Nombre"]',
    'input[data-automation-id="formulaNameBox"]',
    '[data-automation-id="nameBox"] input',
    '[data-automation-id="NameBox"] input',
    '[data-automation-id*="nameBox"] input',
    '[data-automation-id*="NameBox"] input',
    '[data-automation-id*="nameBox"]',
    '[data-automation-id*="NameBox"]',
    '#NameBox input',
    '#NameBox',
    '[role="combobox"][aria-label*="Name"]',
    '[role="combobox"][aria-label*="Nombre"]',
    '[role="textbox"][aria-label*="Name"]',
    '[role="textbox"][aria-label*="Nombre"]'
  ];

  const FORMULA_BAR_SELECTORS = [
    'input[aria-label="Formula Bar"]',
    'input[aria-label="Barra de formulas"]',
    'input[aria-label*="Formula"]',
    'textarea[aria-label*="Formula"]',
    '[role="textbox"][aria-label*="Formula"]',
    '[role="textbox"][aria-label*="formula"]',
    'textarea[aria-label="Formula Bar"]',
    'textarea[aria-label="Barra de formulas"]',
    'div[contenteditable="true"][aria-label*="Formula"]',
    '[data-automation-id*="formulaBar"] input',
    '[data-automation-id*="formulaBar"] textarea',
    '[data-automation-id*="FormulaBar"] input',
    '[data-automation-id*="FormulaBar"] textarea',
    'input[aria-label]',
    'textarea[aria-label]',
    '[role="textbox"][aria-label]',
    'div[contenteditable="true"][aria-label]'
  ];

  const logEvent = (level, message) => {
    chrome.runtime.sendMessage({
      type: "LOG_EVENT",
      level,
      source: "excel_online",
      message
    });
  };

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

  const deepQuerySelector = (selectorList, predicate) => {
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

  const resolveInput = (el) => {
    if (!el) {
      return null;
    }
    return el.querySelector("input, textarea") || el;
  };

  const GRID_SELECTORS = [
    '[role="grid"]',
    '[data-automation-id="grid"]',
    '[data-automation-id*="grid"]',
    '[data-testid*="grid"]',
    '[data-test-id*="grid"]'
  ];

  const findNameBox = () =>
    resolveInput(deepQuerySelector(NAME_BOX_SELECTORS, isLikelyNameBox));

  const findFormulaBar = () =>
    resolveInput(deepQuerySelector(FORMULA_BAR_SELECTORS, isLikelyFormulaBar));

  const findGrid = () => deepQuerySelector(GRID_SELECTORS);

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
    const colIndex = Number.parseInt(cell.getAttribute("aria-colindex"), 10);
    const rowIndex = Number.parseInt(cell.getAttribute("aria-rowindex"), 10);
    if (Number.isFinite(colIndex) && Number.isFinite(rowIndex)) {
      return `${columnIndexToLetter(colIndex)}${rowIndex}`;
    }
    return extractCellRefFromLabel(cell.getAttribute("aria-label") || "");
  };

  const getActiveCellFromGrid = () => {
    const grid = findGrid();
    const container =
      grid || deepQuerySelector(['[aria-activedescendant]'], () => true);
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
    return extractCellRefFromLabel(container.getAttribute("aria-label") || "");
  };

  const getSelectionRange = () => {
    const nameBox = findNameBox();
    if (nameBox) {
      const value =
        nameBox.value ||
        nameBox.getAttribute("value") ||
        nameBox.textContent ||
        "";
      const clean = normalizeRangeToken(value);
      if (isRangeRef(clean)) {
        return clean;
      }
    }
    const gridRef = getActiveCellFromGrid();
    if (gridRef) {
      return gridRef;
    }
    const active = document.activeElement;
    const label = active ? active.getAttribute("aria-label") || "" : "";
    return extractCellRefFromLabel(label);
  };

  const getStartCellFromRange = (range) => {
    const token = normalizeRangeToken(range);
    const match = token.match(/^([A-Z]{1,3}\d{1,7})/);
    return match ? match[1] : "";
  };

  let lastKnownRange = "";
  let lastKnownCell = "";
  let lastKnownAt = 0;
  let lastClipboardText = "";
  let lastClipboardAt = 0;

  const updateSelectionCache = () => {
    const range = getSelectionRange();
    if (!range || !isRangeRef(range)) {
      return;
    }
    lastKnownRange = range;
    lastKnownCell = getStartCellFromRange(range);
    lastKnownAt = Date.now();
    window.__GA_LAST_RANGE__ = lastKnownRange;
    window.__GA_LAST_CELL__ = lastKnownCell;
    window.__GA_LAST_AT__ = lastKnownAt;
  };

  const updateClipboardCache = async () => {
    try {
      const text = await navigator.clipboard.readText();
      if (!text) {
        return;
      }
      lastClipboardText = text;
      lastClipboardAt = Date.now();
      window.__GA_LAST_CLIPBOARD__ = lastClipboardText;
      window.__GA_LAST_CLIPBOARD_AT__ = lastClipboardAt;
    } catch (error) {
      // Ignore clipboard errors; user gesture is required.
    }
  };

  const scheduleClipboardRead = () => {
    setTimeout(() => {
      updateClipboardCache();
    }, 60);
  };

  const handleCopyEvent = (event) => {
    try {
      const text =
        event && event.clipboardData
          ? event.clipboardData.getData("text/plain")
          : "";
      if (text) {
        lastClipboardText = text;
        lastClipboardAt = Date.now();
        window.__GA_LAST_CLIPBOARD__ = lastClipboardText;
        window.__GA_LAST_CLIPBOARD_AT__ = lastClipboardAt;
      }
    } catch (error) {
      // Ignore clipboard event errors.
    }
    scheduleClipboardRead();
  };

  const getActiveCellValue = () => {
    const formulaBar = findFormulaBar();
    if (!formulaBar) {
      return "";
    }
    const value =
      formulaBar.value ||
      formulaBar.getAttribute("value") ||
      formulaBar.innerText ||
      formulaBar.textContent ||
      "";
    return String(value);
  };

  const extractColumnFromRange = (range) => {
    const cleaned = normalizeRangeToken(range);
    if (!cleaned) {
      return "";
    }
    const match = cleaned.toUpperCase().match(/^([A-Z]+)\d*/);
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
      const grid = findGrid();
      if (grid && grid.focus) {
        grid.focus();
      }
      document.execCommand("copy");
      return await tryReadClipboard();
    } catch (error) {
      return "";
    }
  };

  const copySelectionToClipboard = async () => {
    await waitForExcelUi();
    let copyOk = false;
    try {
      const grid = findGrid();
      if (grid && grid.focus) {
        grid.focus();
      }
      copyOk = document.execCommand("copy");
    } catch (error) {
      copyOk = false;
    }
    if (copyOk) {
      scheduleClipboardRead();
    }
    return { ok: copyOk };
  };

  const setSelectionRange = async (range) => {
    const target = normalizeRangeToken(range);
    if (!target) {
      return { ok: false, reason: "empty_range" };
    }
    await waitForExcelUi();
    const nameBox = findNameBox();
    if (!nameBox) {
      return { ok: false, reason: "name_box_missing" };
    }
    const input = resolveInput(nameBox);
    if (!input || typeof input.focus !== "function") {
      return { ok: false, reason: "name_box_invalid" };
    }
    input.focus();
    if ("value" in input) {
      input.value = target;
    } else {
      input.textContent = target;
    }
    input.dispatchEvent(new Event("input", { bubbles: true }));
    input.dispatchEvent(new Event("change", { bubbles: true }));
    input.dispatchEvent(
      new KeyboardEvent("keydown", { key: "Enter", code: "Enter", bubbles: true })
    );
    input.dispatchEvent(
      new KeyboardEvent("keyup", { key: "Enter", code: "Enter", bubbles: true })
    );
    scheduleCacheUpdate();
    return { ok: true };
  };

  const hasExcelUi = () =>
    !!findNameBox() || !!findFormulaBar() || !!findGrid();

  const waitForExcelUi = async (timeoutMs = 6000) => {
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

  const scheduleCacheUpdate = () => {
    setTimeout(updateSelectionCache, 0);
  };

  const startSelectionWatcher = () => {
    const attachObserver = () => {
      const grid = findGrid();
      if (!grid) {
        return false;
      }
      const observer = new MutationObserver(() => {
        scheduleCacheUpdate();
      });
      observer.observe(grid, {
        attributes: true,
        subtree: true,
        attributeFilter: ["aria-activedescendant", "aria-selected", "tabindex"]
      });
      return true;
    };

    if (!attachObserver()) {
      let attempts = 0;
      const timer = setInterval(() => {
        attempts += 1;
        if (attachObserver() || attempts >= 10) {
          clearInterval(timer);
        }
      }, 500);
    }

    document.addEventListener("mouseup", scheduleCacheUpdate, true);
    document.addEventListener("keyup", scheduleCacheUpdate, true);
    document.addEventListener("pointerup", scheduleCacheUpdate, true);
    document.addEventListener("focusin", scheduleCacheUpdate, true);
    document.addEventListener("selectionchange", scheduleCacheUpdate, true);
    document.addEventListener("copy", handleCopyEvent, true);
    document.addEventListener(
      "keydown",
      (event) => {
        if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "c") {
          scheduleClipboardRead();
        }
      },
      true
    );
  };

  const captureSelection = async () => {
    await waitForExcelUi();
    let range = getSelectionRange();
    if (!range) {
      await new Promise((resolve) => setTimeout(resolve, 300));
      range = getSelectionRange();
    }
    if (!range && lastKnownRange) {
      range = lastKnownRange;
    }
    const column = extractColumnFromRange(range);
    let clipboardText = await tryReadClipboard();
    if (!clipboardText) {
      clipboardText = await tryCopySelection();
    }
    if (!clipboardText && lastClipboardText) {
      clipboardText = lastClipboardText;
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

    if (message.type === "PING") {
      updateSelectionCache();
      sendResponse({
        ok: hasExcelUi(),
        range: lastKnownRange || "",
        cell: lastKnownCell || "",
        clipboard: lastClipboardText || "",
        clipboardAt: lastClipboardAt || 0,
        url: location.href
      });
      return;
    }

    if (message.type === "CAPTURE_SELECTION") {
      updateSelectionCache();
      captureSelection().then((payload) => {
        sendResponse({ ...payload, ok: hasExcelUi() });
      });
      return true;
    }

    if (message.type === "COPY_SELECTION") {
      copySelectionToClipboard().then(sendResponse);
      return true;
    }

    if (message.type === "GET_ACTIVE_CELL") {
      (async () => {
        await waitForExcelUi();
        updateSelectionCache();
        const activeCell = getSelectionRange() || lastKnownCell || lastKnownRange;
        const activeValue = getActiveCellValue();
        sendResponse({ activeCell, activeValue, ok: Boolean(activeCell) });
      })();
      return true;
    }

    if (message.type === "SET_SELECTION") {
      setSelectionRange(message.range || "").then(sendResponse);
      return true;
    }

    if (message.type === "PASTE_TSV") {
      if (!hasExcelUi()) {
        sendResponse({ ok: false, reason: "no_ui" });
        return;
      }
      pasteTsv(message.tsv || "").then(sendResponse);
      return true;
    }
  });

  const ready = () => {
    startSelectionWatcher();
    scheduleCacheUpdate();
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
