(() => {
  if (window.__GA_AMAZON_LOADED__) {
    return;
  }
  window.__GA_AMAZON_LOADED__ = true;

  const FILTER_LABELS = {
    order: ["Order ID", "Order id", "ID de pedido", "ID del pedido", "Pedido"],
    tracking: [
      "Tracking ID",
      "Tracking Id",
      "ID de rastreo",
      "ID de seguimiento",
      "Guia"
    ]
  };

  const SEARCH_INPUT_SELECTORS = [
    'input[type="search"]',
    'input[placeholder*="Search"]',
    'input[placeholder*="Buscar"]',
    'input[aria-label*="Search"]',
    'input[aria-label*="Buscar"]',
    'input[data-test-id*="search"]',
    'input[data-testid*="search"]',
    'input[id*="search"]',
    'input[name*="search"]',
    'input[role="searchbox"]',
    '[role="search"] input',
    '[data-test-id*="order-search"] input',
    '[data-testid*="order-search"] input'
  ];

  const FILTER_BUTTON_SELECTORS = [
    'button[aria-haspopup="listbox"]',
    'button[aria-haspopup="menu"]',
    'button[aria-label*="Filter"]',
    'button[aria-label*="Filtro"]',
    '[data-test-id*="search-filter"]',
    '[data-testid*="search-filter"]'
  ];

  const ORDER_ROW_SELECTORS = [
    '[data-test-id*="order-row"]',
    '[data-testid*="order-row"]',
    'table[data-test-id*="orders"] tbody tr',
    'table[data-testid*="orders"] tbody tr',
    "table tbody tr"
  ];

  const ORDER_DETAIL_SELECTORS = [
    '[data-test-id*="order-detail"]',
    '[data-testid*="order-detail"]',
    '[data-test-id*="order-details"]',
    '[data-testid*="order-details"]',
    '[aria-label*="Order details"]',
    '[aria-label*="Detalles del pedido"]',
    '[data-test-id*="order-details-panel"]',
    '[data-testid*="order-details-panel"]'
  ];

  const ORDER_DETAIL_CLOSE_SELECTORS = [
    'button[aria-label*="Close"]',
    'button[aria-label*="Cerrar"]',
    'button[aria-label*="Back"]',
    'button[aria-label*="Regresar"]',
    'a[aria-label*="Back"]',
    'a[aria-label*="Regresar"]',
    '[data-test-id*="close"]',
    '[data-testid*="close"]'
  ];

  const ORDER_LINK_SELECTORS = [
    'a[href*="order"]',
    'a[href*="orders-v3"]',
    'a[data-test-id*="order-id"]',
    'a[data-testid*="order-id"]'
  ];

  const NO_RESULTS_TEXT = [
    "no orders found",
    "no se encontraron pedidos",
    "no se encontraron ordenes",
    "sin resultados",
    "0 orders",
    "0 pedidos"
  ];

  const logEvent = (level, message) => {
    chrome.runtime.sendMessage({
      type: "LOG_EVENT",
      level,
      source: "amazon_seller",
      message
    });
  };

  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  const toLower = (value) => String(value || "").trim().toLowerCase();

  const isVisible = (el) => el && el.offsetParent !== null;

  const findSearchInput = () => {
    for (const selector of SEARCH_INPUT_SELECTORS) {
      const el = document.querySelector(selector);
      if (el && isVisible(el)) {
        return el;
      }
    }
    return null;
  };

  const findFilterButton = () => {
    for (const selector of FILTER_BUTTON_SELECTORS) {
      const el = document.querySelector(selector);
      if (el && isVisible(el)) {
        return el;
      }
    }
    return null;
  };

  const getOrderRows = () => {
    for (const selector of ORDER_ROW_SELECTORS) {
      const rows = Array.from(document.querySelectorAll(selector)).filter(
        (row) => row.querySelector("td") || row.getAttribute("data-test-id")
      );
      if (rows.length) {
        return rows;
      }
    }
    return [];
  };

  const hasNoResults = () => {
    const bodyText = toLower(document.body ? document.body.textContent : "");
    return NO_RESULTS_TEXT.some((text) => bodyText.includes(text));
  };

  const waitForResults = async (timeoutMs = 15000) => {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      const rows = getOrderRows();
      if (rows.length > 0) {
        return "rows";
      }
      if (hasNoResults()) {
        return "empty";
      }
      await sleep(300);
    }
    return "timeout";
  };

  const extractAsinsFromText = (text) => {
    const matches = String(text || "").match(/[A-Z0-9]{10}/g);
    if (!matches) {
      return [];
    }
    return Array.from(new Set(matches));
  };

  const extractAsinsFromRow = (row) => {
    const asins = new Set();
    const candidates = row.querySelectorAll("a, span, div");
    candidates.forEach((el) => {
      const text = (el.textContent || "").trim();
      if (/^[A-Z0-9]{10}$/.test(text)) {
        asins.add(text);
      }
    });
    if (asins.size === 0) {
      extractAsinsFromText(row.textContent || "").forEach((asin) => asins.add(asin));
    }
    return Array.from(asins);
  };

  const extractQuantityFromText = (text) => {
    const match = String(text || "").match(
      /(?:Qty|Quantity|Cantidad)\s*[:\s]\s*(\d+)/i
    );
    if (match) {
      return Number.parseInt(match[1], 10);
    }
    const numMatch = String(text || "").match(/\b(\d+)\b/);
    if (numMatch) {
      const number = Number.parseInt(numMatch[1], 10);
      if (Number.isFinite(number) && number > 0 && number < 1000) {
        return number;
      }
    }
    return 1;
  };

  const extractQuantityFromRow = (row) => extractQuantityFromText(row.textContent || "");

  const parseRowItems = (row) => {
    const asins = extractAsinsFromRow(row);
    const qty = extractQuantityFromRow(row);
    return asins.map((asin) => ({ asin, quantity: qty }));
  };

  const clickMatchingMenuItem = (labels) => {
    const targets = Array.from(
      document.querySelectorAll('[role="option"], [role="menuitem"], button, li')
    );
    const labelSet = labels.map((label) => label.toLowerCase());
    const match = targets.find((el) => {
      if (el.offsetParent === null) {
        return false;
      }
      const text = (el.textContent || "").trim().toLowerCase();
      return text && labelSet.some((label) => text === label || text.includes(label));
    });
    if (match) {
      match.click();
      return true;
    }
    return false;
  };

  const selectSearchFilter = async (type) => {
    const labelCandidates = type === "order" ? FILTER_LABELS.order : FILTER_LABELS.tracking;
    const selects = Array.from(document.querySelectorAll("select"));
    const selectHit = selects.find((select) => {
      const options = Array.from(select.options || []);
      return options.some((opt) => {
        const text = (opt.textContent || "").toLowerCase();
        return labelCandidates.some((label) => text.includes(label.toLowerCase()));
      });
    });
    if (selectHit) {
      const option = Array.from(selectHit.options).find((opt) => {
        const text = (opt.textContent || "").toLowerCase();
        return labelCandidates.some((label) => text.includes(label.toLowerCase()));
      });
      if (option) {
        selectHit.value = option.value;
        selectHit.dispatchEvent(new Event("change", { bubbles: true }));
        return true;
      }
    }
    const button = findFilterButton();
    if (!button) {
      return false;
    }
    button.click();
    await sleep(200);
    return clickMatchingMenuItem(labelCandidates);
  };

  const triggerSearch = (input) => {
    input.dispatchEvent(new Event("input", { bubbles: true }));
    input.dispatchEvent(
      new KeyboardEvent("keydown", { key: "Enter", code: "Enter", bubbles: true })
    );
    const searchButton =
      document.querySelector('button[aria-label*="Search"]') ||
      document.querySelector('button[aria-label*="Buscar"]') ||
      document.querySelector('[data-test-id*="search"] button') ||
      document.querySelector('[data-testid*="search"] button');
    if (searchButton) {
      searchButton.click();
    }
  };

  const findOrderDetailRoot = () => {
    for (const selector of ORDER_DETAIL_SELECTORS) {
      const el = document.querySelector(selector);
      if (el) {
        return el;
      }
    }
    return null;
  };

  const waitForOrderDetail = async (timeoutMs = 8000) => {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      const root = findOrderDetailRoot();
      if (root) {
        return root;
      }
      await sleep(300);
    }
    return null;
  };

  const parseOrderDetailsTable = (root) => {
    const tables = Array.from(root.querySelectorAll("table"));
    for (const table of tables) {
      const rows = Array.from(table.querySelectorAll("tr"));
      if (!rows.length) {
        continue;
      }
      const headerCells = Array.from(rows[0].querySelectorAll("th, td"));
      const headers = headerCells.map((cell) => toLower(cell.textContent));
      const asinIndex = headers.findIndex((text) => text.includes("asin"));
      const qtyIndex = headers.findIndex(
        (text) => text.includes("qty") || text.includes("quantity") || text.includes("cantidad")
      );
      if (asinIndex === -1) {
        continue;
      }
      const items = [];
      rows.slice(1).forEach((row) => {
        const cells = Array.from(row.querySelectorAll("td, th"));
        const asin = cells[asinIndex] ? cells[asinIndex].textContent.trim() : "";
        if (!asin || !/^[A-Z0-9]{10}$/.test(asin)) {
          return;
        }
        const qtyText = cells[qtyIndex] ? cells[qtyIndex].textContent : "";
        const qty = extractQuantityFromText(qtyText);
        items.push({ asin, quantity: qty });
      });
      if (items.length) {
        return items;
      }
    }
    return [];
  };

  const parseOrderDetails = () => {
    const root = findOrderDetailRoot() || document;
    let items = parseOrderDetailsTable(root);
    if (items.length) {
      return items;
    }

    const itemBlocks = Array.from(
      root.querySelectorAll('[data-test-id*="item"], [data-testid*="item"]')
    );
    itemBlocks.forEach((block) => {
      const text = block.textContent || "";
      const asins = extractAsinsFromText(text);
      if (!asins.length) {
        return;
      }
      const qty = extractQuantityFromText(text);
      asins.forEach((asin) => items.push({ asin, quantity: qty }));
    });
    if (items.length) {
      return items;
    }

    const fallbackAsins = extractAsinsFromText(root.textContent || "");
    if (fallbackAsins.length) {
      return fallbackAsins.map((asin) => ({ asin, quantity: 1 }));
    }
    return [];
  };

  const openOrderDetailsFromRow = async (row) => {
    for (const selector of ORDER_LINK_SELECTORS) {
      const link = row.querySelector(selector);
      if (link) {
        link.click();
        const detail = await waitForOrderDetail();
        return Boolean(detail);
      }
    }
    row.click();
    const detail = await waitForOrderDetail();
    return Boolean(detail);
  };

  const closeOrderDetails = async () => {
    for (const selector of ORDER_DETAIL_CLOSE_SELECTORS) {
      const button = document.querySelector(selector);
      if (button) {
        button.click();
        await sleep(400);
        return true;
      }
    }
    if (window.history.length > 1) {
      window.history.back();
      await sleep(800);
      return true;
    }
    return false;
  };

  const ensureOrdersUi = () => {
    return Boolean(findSearchInput());
  };

  const isLoginPage = () => {
    const url = toLower(location.href);
    if (url.includes("signin") || url.includes("login")) {
      return true;
    }
    const passwordField = document.querySelector('input[type="password"]');
    return Boolean(passwordField);
  };

  const searchGuide = async (guide) => {
    if (isLoginPage()) {
      logEvent("error", "Sesion de Amazon Seller no detectada.");
      return { guide, status: "error", note: "login_required" };
    }
    const input = findSearchInput();
    if (!input) {
      logEvent("error", "No se encontro el campo de busqueda en Amazon.");
      return { guide, status: "error", note: "search_input_not_found" };
    }
    if (!ensureOrdersUi()) {
      logEvent("error", "No se detecto la vista de ordenes.");
      return { guide, status: "error", note: "orders_ui_not_found" };
    }

    const searchType = guide.includes("-") ? "order" : "tracking";
    const filterOk = await selectSearchFilter(searchType);
    if (!filterOk) {
      logEvent("error", `No se pudo cambiar el filtro a ${searchType}.`);
    }
    input.value = guide;
    triggerSearch(input);
    const status = await waitForResults();
    if (status === "timeout") {
      logEvent("error", "Tiempo de espera agotado al cargar resultados.");
      return { guide, status: "error", note: "results_timeout" };
    }
    if (status === "empty") {
      return { guide, status: "not_found", items: [] };
    }

    const rows = getOrderRows();
    if (!rows.length) {
      return { guide, status: "not_found", items: [] };
    }

    const items = [];
    for (const row of rows) {
      let rowItems = [];
      const opened = await openOrderDetailsFromRow(row);
      if (opened) {
        await sleep(300);
        rowItems = parseOrderDetails();
        await closeOrderDetails();
      }
      if (!rowItems.length) {
        rowItems = parseRowItems(row);
      }
      rowItems.forEach((item) => {
        if (item && item.asin) {
          items.push(item);
        }
      });
      await sleep(250);
    }

    if (!items.length) {
      return { guide, status: "not_found", items: [] };
    }
    return { guide, status: "found", items };
  };

  const lookupGuides = async (guides) => {
    const results = [];
    for (const guide of guides) {
      logEvent("info", `Buscando guia ${guide}...`);
      const result = await searchGuide(guide);
      results.push(result);
      await sleep(400);
    }
    return results;
  };

  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (!message || typeof message !== "object") {
      return;
    }

    if (message.type === "PING") {
      sendResponse({ ok: true, url: location.href });
      return;
    }

    if (message.type === "LOOKUP_GUIDES") {
      lookupGuides(message.guides || []).then((results) => sendResponse({ ok: true, results }));
      return true;
    }
  });

  const ready = () => {
    logEvent("info", "Amazon Seller content script listo.");
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", ready, { once: true });
  } else {
    ready();
  }
})();
