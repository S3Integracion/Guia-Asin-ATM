(() => {
  if (window.__GA_AMAZON_ORDERS__) {
    return;
  }
  window.__GA_AMAZON_ORDERS__ = true;

  const FILTER_LABELS = {
    order: ["Order ID", "Order Id", "ID de pedido", "ID del pedido", "Pedido"],
    tracking: [
      "Tracking ID",
      "Tracking Id",
      "ID de seguimiento",
      "ID de rastreo",
      "Guia",
      "Guía"
    ]
  };

  const SEARCH_INPUT_SELECTORS = [
    'input[type="search"]',
    'input[type="text"]',
    'input[placeholder*="Search"]',
    'input[placeholder*="Buscar"]',
    'input[placeholder*="tools"]',
    'input[placeholder*="herramient"]',
    'input[aria-label*="Search"]',
    'input[aria-label*="Buscar"]',
    'input[aria-label*="Search orders"]',
    'input[aria-label*="Buscar pedidos"]',
    'input[data-test-id*="search"]',
    'input[data-testid*="search"]',
    'input[id*="search"]',
    'input[name*="search"]',
    'input[role="searchbox"]',
    '[role="search"] input'
  ];

  const SEARCH_INPUT_EXCLUDE = [
    "search for tools",
    "buscar herramientas",
    "buscar herramientas, ayuda",
    "search for tools, help"
  ];

  const SEARCH_BUTTON_SELECTORS = [
    'button[type="submit"]',
    'button[aria-label*="Search"]',
    'button[aria-label*="Buscar"]',
    'button:contains("Search")',
    'button:contains("Buscar")'
  ];

  const ORDER_ROW_SELECTORS = [
    '[data-test-id*="order-row"]',
    '[data-testid*="order-row"]',
    'table[data-test-id*="orders"] tbody tr',
    'table[data-testid*="orders"] tbody tr',
    "table tbody tr"
  ];

  const NO_RESULTS_TEXT = [
    "no orders found",
    "no se encontraron pedidos",
    "no se encontraron ordenes",
    "no se encontraron órdenes",
    "sin resultados",
    "0 orders"
  ];

  let cancelRequested = false;

  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  const isVisible = (el) => el && el.offsetParent !== null;

  const toLower = (value) => String(value || "").trim().toLowerCase();

  const deepQueryAll = (selectorList) => {
    const results = [];
    const roots = [document];
    const visited = new Set();
    while (roots.length) {
      const root = roots.shift();
      if (!root || visited.has(root)) {
        continue;
      }
      visited.add(root);
      for (const selector of selectorList) {
        root.querySelectorAll?.(selector).forEach((el) => results.push(el));
      }
      root.querySelectorAll?.("*").forEach((node) => {
        if (node.shadowRoot) {
          roots.push(node.shadowRoot);
        }
      });
    }
    return results;
  };

  const isExcludedInput = (input) => {
    const text = toLower(
      input?.getAttribute("placeholder") ||
        input?.getAttribute("aria-label") ||
        input?.getAttribute("name") ||
        ""
    );
    return SEARCH_INPUT_EXCLUDE.some((phrase) => text.includes(phrase));
  };

  const findSearchInput = () => {
    const inputs = deepQueryAll(SEARCH_INPUT_SELECTORS)
      .filter((el) => el && isVisible(el))
      .filter((el) => !isExcludedInput(el));
    if (!inputs.length) {
      return null;
    }

    const filterSelect = findFilterSelect();
    if (filterSelect) {
      const scope = filterSelect.closest("form, div, section") || filterSelect.parentElement;
      if (scope) {
        const scopedInput = inputs.find((input) => scope.contains(input));
        if (scopedInput) {
          return scopedInput;
        }
      }
    }

    const scored = inputs.map((input) => {
      let score = 0;
      let node = input;
      for (let i = 0; i < 6 && node; i += 1) {
        const text = toLower(node.textContent || "");
        if (
          text.includes("manage orders") ||
          text.includes("order id") ||
          text.includes("tracking id") ||
          text.includes("id de pedido") ||
          text.includes("id de seguimiento")
        ) {
          score += 2;
        }
        if (node.querySelector?.("select")) {
          score += 2;
        }
        if (node.querySelector?.("button")) {
          score += 1;
        }
        node = node.parentElement;
      }
      return { input, score };
    });
    scored.sort((a, b) => b.score - a.score);
    return scored[0].input;
  };

  const findSearchButton = (input) => {
    const container = input ? input.closest("form, div, section") : null;
    if (container) {
      const buttons = Array.from(container.querySelectorAll("button, input[type=\"submit\"]"));
      const hit = buttons.find((btn) =>
        /search|buscar/i.test(btn.textContent || "")
      );
      if (hit && isVisible(hit)) {
        return hit;
      }
      const inputBtn = buttons.find((btn) =>
        /search|buscar/i.test(btn.getAttribute("value") || "")
      );
      if (inputBtn && isVisible(inputBtn)) {
        return inputBtn;
      }
    }
    const buttons = Array.from(document.querySelectorAll("button, input[type=\"submit\"]")).filter(
      isVisible
    );
    return buttons.find((btn) => /search|buscar/i.test(btn.textContent || "")) || null;
  };

  const findFilterSelect = () => {
    const selects = Array.from(document.querySelectorAll("select"));
    return selects.find((select) => {
      const options = Array.from(select.options || []);
      return options.some((opt) => {
        const text = toLower(opt.textContent || "");
        return (
          FILTER_LABELS.order.some((label) => text.includes(toLower(label))) ||
          FILTER_LABELS.tracking.some((label) => text.includes(toLower(label)))
        );
      });
    });
  };

  const clickMatchingMenuItem = (labels) => {
    const labelSet = labels.map((label) => toLower(label));
    const candidates = Array.from(
      document.querySelectorAll('[role="option"], [role="menuitem"], button, li')
    );
    const match = candidates.find((el) => {
      if (!isVisible(el)) {
        return false;
      }
      const text = toLower(el.textContent || "");
      return text && labelSet.some((label) => text.includes(label));
    });
    if (match) {
      match.click();
      return true;
    }
    return false;
  };

  const selectSearchFilter = async (type) => {
    const labelCandidates = type === "order" ? FILTER_LABELS.order : FILTER_LABELS.tracking;
    const select = findFilterSelect();
    if (select) {
      const option = Array.from(select.options || []).find((opt) => {
        const text = toLower(opt.textContent || "");
        return labelCandidates.some((label) => text.includes(toLower(label)));
      });
      if (option) {
        select.value = option.value;
        select.dispatchEvent(new Event("change", { bubbles: true }));
        return true;
      }
    }
    const buttons = Array.from(document.querySelectorAll("button")).filter(isVisible);
    const filterButton = buttons.find((btn) =>
      labelCandidates.some((label) =>
        toLower(btn.textContent || "").includes(toLower(label))
      )
    );
    if (filterButton) {
      filterButton.click();
      return true;
    }
    const dropdown = buttons.find((btn) => {
      const hasPopup = btn.getAttribute("aria-haspopup");
      if (!hasPopup) {
        return false;
      }
      const text = toLower(btn.textContent || "");
      return (
        text.includes("order id") ||
        text.includes("tracking id") ||
        text.includes("id de pedido") ||
        text.includes("id de seguimiento")
      );
    });
    if (!dropdown) {
      return false;
    }
    dropdown.click();
    await sleep(200);
    return clickMatchingMenuItem(labelCandidates);
  };

  const triggerSearch = (input, value) => {
    if (!input) {
      return;
    }
    input.focus();
    input.value = "";
    input.dispatchEvent(new Event("input", { bubbles: true }));
    input.value = value;
    input.dispatchEvent(new Event("input", { bubbles: true }));
    input.dispatchEvent(
      new KeyboardEvent("keydown", { key: "Enter", code: "Enter", bubbles: true })
    );
    input.dispatchEvent(
      new KeyboardEvent("keyup", { key: "Enter", code: "Enter", bubbles: true })
    );
    const button = findSearchButton(input);
    if (button) {
      button.click();
    }
  };

  const getOrderRows = () => {
    for (const selector of ORDER_ROW_SELECTORS) {
      const rows = Array.from(document.querySelectorAll(selector)).filter((row) =>
        /asin/i.test(row.textContent || "")
      );
      if (rows.length) {
        return rows;
      }
    }
    return [];
  };

  const hasNoResults = () => {
    const text = toLower(document.body?.textContent || "");
    return NO_RESULTS_TEXT.some((phrase) => text.includes(phrase));
  };

  const getRowsSnapshot = () => {
    const rows = getOrderRows();
    if (!rows.length) {
      return "";
    }
    return rows.map((row) => toLower(row.textContent || "")).join(" | ");
  };

  const waitForResults = async (guide, timeoutMs = 20000) => {
    const start = Date.now();
    const initialRows = getRowsSnapshot();
    while (Date.now() - start < timeoutMs) {
      if (hasNoResults()) {
        return "empty";
      }
      const rows = getOrderRows();
      if (rows.length) {
        const currentRows = getRowsSnapshot();
        if (currentRows && currentRows !== initialRows) {
          return "rows";
        }
        if (guide) {
          const bodyText = toLower(document.body?.textContent || "");
          if (bodyText.includes(toLower(guide))) {
            return "rows";
          }
        }
      }
      await sleep(300);
    }
    return "timeout";
  };

  const extractAsinsFromText = (text) => {
    const matches = Array.from(
      String(text || "").matchAll(/ASIN\s*[:#]?\s*([A-Z0-9]{10})/gi)
    ).map((match) => match[1]);
    if (matches.length) {
      return Array.from(new Set(matches));
    }
    const fallback = Array.from(
      String(text || "").matchAll(/\b[A-Z][A-Z0-9]{9}\b/g)
    ).map((match) => match[0]);
    return Array.from(new Set(fallback));
  };

  const extractQuantityMatches = (text) => {
    return Array.from(
      String(text || "").matchAll(
        /(?:Qty|Quantity|Cantidad)\s*[:\s]\s*(\d+)/gi
      )
    ).map((match) => Number.parseInt(match[1], 10));
  };

  const extractTrackingFromText = (text) => {
    const match = String(text || "").match(
      /(?:Tracking\s*ID|ID\s*de\s*seguimiento|ID\s*de\s*rastreo|Guia|Guía)\s*[:\s]\s*([A-Z0-9-]+)/i
    );
    return match ? match[1] : "";
  };

  const parseItemsFromText = (text, guide) => {
    const asins = extractAsinsFromText(text);
    const qtyMatches = extractQuantityMatches(text);
    const tracking = extractTrackingFromText(text) || guide;

    if (!asins.length) {
      return { items: [], tracking };
    }

    let quantities = [];
    if (qtyMatches.length === asins.length) {
      quantities = qtyMatches;
    } else if (qtyMatches.length === 1) {
      quantities = asins.map(() => qtyMatches[0]);
    } else {
      quantities = asins.map(() => null);
    }

    const items = asins.map((asin, index) => ({
      asin,
      quantity: quantities[index] || null,
      tracking
    }));

    return { items, tracking };
  };

  const extractItemsFromPage = (guide) => {
    const rows = getOrderRows();
    if (rows.length) {
      const items = [];
      let tracking = "";
      rows.forEach((row) => {
        const parsed = parseItemsFromText(row.textContent || "", guide);
        if (parsed.tracking && !tracking) {
          tracking = parsed.tracking;
        }
        parsed.items.forEach((item) => items.push(item));
      });
      const merged = new Map();
      items.forEach((item) => {
        if (!item.asin) {
          return;
        }
        if (!merged.has(item.asin)) {
          merged.set(item.asin, { ...item });
          return;
        }
        const existing = merged.get(item.asin);
        if (existing && item.quantity && existing.quantity) {
          existing.quantity += item.quantity;
        }
      });
      return { items: Array.from(merged.values()), tracking };
    }

    const parsed = parseItemsFromText(document.body?.textContent || "", guide);
    return parsed;
  };

  const isLoginPage = () => {
    const url = toLower(location.href);
    if (url.includes("signin") || url.includes("login")) {
      return true;
    }
    return Boolean(document.querySelector('input[type="password"]'));
  };

  const searchGuide = async (guide) => {
    if (isLoginPage()) {
      return { guide, status: "error", note: "Sesion no iniciada." };
    }
    const input = findSearchInput();
    if (!input) {
      return { guide, status: "error", note: "No se encontro el buscador." };
    }
    const searchType = guide.includes("-") ? "order" : "tracking";
    const filterOk = await selectSearchFilter(searchType);
    if (!filterOk) {
      return {
        guide,
        status: "error",
        note: "No se pudo seleccionar el filtro."
      };
    }
    await sleep(250);
    triggerSearch(input, guide);
    const status = await waitForResults(guide);
    if (status === "timeout") {
      return { guide, status: "error", note: "Tiempo de espera agotado." };
    }
    if (status === "empty") {
      return { guide, status: "not_found", items: [] };
    }
    const extracted = extractItemsFromPage(guide);
    if (!extracted.items || !extracted.items.length) {
      return { guide, status: "not_found", items: [] };
    }
    return {
      guide,
      status: "found",
      items: extracted.items,
      tracking: extracted.tracking || guide
    };
  };

  const sendProgress = (payload) => {
    chrome.runtime.sendMessage({ type: "LOOKUP_PROGRESS", payload });
  };

  const lookupGuides = async (guides) => {
    const results = [];
    cancelRequested = false;
    for (let i = 0; i < guides.length; i += 1) {
      const guide = guides[i];
      if (cancelRequested) {
        break;
      }
      sendProgress({ guide, index: i + 1, total: guides.length });
      const result = await searchGuide(guide);
      results.push(result);
      await sleep(250);
    }
    return { results, canceled: cancelRequested, total: guides.length };
  };

  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (!message || typeof message !== "object") {
      return;
    }
    if (message.type === "PING") {
      sendResponse({ ok: true });
      return;
    }
    if (message.type === "CANCEL_LOOKUP") {
      cancelRequested = true;
      sendResponse({ ok: true });
      return;
    }
    if (message.type === "START_LOOKUP") {
      const guides = Array.isArray(message.guides) ? message.guides : [];
      lookupGuides(guides).then((payload) => {
        chrome.runtime.sendMessage({
          type: "LOOKUP_COMPLETE",
          payload
        });
        sendResponse({ ok: true });
      });
      return true;
    }
  });
})();
