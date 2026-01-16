(() => {
  const FILTER_LABELS = {
    order: [
      "Order ID",
      "Order id",
      "ID de pedido",
      "ID del pedido",
      "Pedido"
    ],
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
    'input[data-test-id*="search"]'
  ];

  const FILTER_BUTTON_SELECTORS = [
    'button[aria-haspopup="listbox"]',
    'button[aria-haspopup="menu"]',
    'button[aria-label*="Filter"]',
    'button[aria-label*="Filtro"]',
    '[data-test-id*="search-filter"]'
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

  const findSearchInput = () => {
    for (const selector of SEARCH_INPUT_SELECTORS) {
      const el = document.querySelector(selector);
      if (el) {
        return el;
      }
    }
    return null;
  };

  const findFilterButton = () => {
    for (const selector of FILTER_BUTTON_SELECTORS) {
      const el = document.querySelector(selector);
      if (el) {
        return el;
      }
    }
    return null;
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
      return text && labelSet.some((label) => text === label);
    });
    if (match) {
      match.click();
      return true;
    }
    return false;
  };

  const selectSearchFilter = async (type) => {
    const button = findFilterButton();
    if (!button) {
      return false;
    }
    button.click();
    await sleep(200);
    const labels = type === "order" ? FILTER_LABELS.order : FILTER_LABELS.tracking;
    if (clickMatchingMenuItem(labels)) {
      return true;
    }
    return false;
  };

  const triggerSearch = (input) => {
    input.dispatchEvent(new Event("input", { bubbles: true }));
    input.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", code: "Enter", bubbles: true }));
    const searchButton =
      document.querySelector('button[aria-label*="Search"]') ||
      document.querySelector('button[aria-label*="Buscar"]') ||
      document.querySelector('[data-test-id*="search"] button');
    if (searchButton) {
      searchButton.click();
    }
  };

  const waitForResults = async () => {
    const start = Date.now();
    while (Date.now() - start < 12000) {
      const rows = document.querySelectorAll('[data-test-id*="order-row"], tr');
      if (rows.length > 0) {
        return true;
      }
      await sleep(300);
    }
    return false;
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
      const match = (row.textContent || "").match(/[A-Z0-9]{10}/g);
      if (match) {
        match.forEach((asin) => asins.add(asin));
      }
    }
    return Array.from(asins);
  };

  const extractQuantityFromRow = (row) => {
    const text = row.textContent || "";
    const match = text.match(/(?:Qty|Cantidad)\s*[:\s]\s*(\d+)/i);
    if (match) {
      return Number.parseInt(match[1], 10);
    }
    const numMatch = text.match(/\b(\d+)\b/);
    if (numMatch) {
      const number = Number.parseInt(numMatch[1], 10);
      if (Number.isFinite(number) && number > 0 && number < 1000) {
        return number;
      }
    }
    return 1;
  };

  const parseResults = () => {
    const rows = Array.from(document.querySelectorAll('[data-test-id*="order-row"], tr'));
    const items = [];
    rows.forEach((row) => {
      const asins = extractAsinsFromRow(row);
      if (asins.length === 0) {
        return;
      }
      const qty = extractQuantityFromRow(row);
      asins.forEach((asin) => {
        items.push({ asin, quantity: qty });
      });
    });
    return items;
  };

  const searchGuide = async (guide) => {
    const input = findSearchInput();
    if (!input) {
      return { guide, status: "error", note: "search_input_not_found" };
    }
    const searchType = guide.includes("-") ? "order" : "tracking";
    await selectSearchFilter(searchType);
    input.value = guide;
    triggerSearch(input);
    const ready = await waitForResults();
    if (!ready) {
      return { guide, status: "error", note: "results_timeout" };
    }
    const items = parseResults();
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
