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

  const ORDER_ID_PATTERN = /^\d{3}-\d{7}-\d{7}$/;
  const TRACKING_ID_PATTERN = /^\d{8,14}$/;

  const BACK_TO_LIST_LABELS = [
    "go back to order list",
    "back to order list",
    "volver a la lista de pedidos",
    "regresar a la lista de pedidos",
    "ir a la lista de pedidos",
    "volver al listado de pedidos",
    "regresar al listado de pedidos"
  ];

  const ORDER_CONTENTS_LABELS = [
    "order contents",
    "contenido del pedido",
    "contenido de la orden",
    "articulos del pedido",
    "artículos del pedido"
  ];

  const ORDER_DETAILS_INDICATORS = [
    "order summary",
    "resumen del pedido",
    "order contents",
    "contenido del pedido",
    "contenido de la orden"
  ];

  const SEARCH_DELAY_MS = 1000;

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
  let lookupInProgress = false;

  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  const isVisible = (el) => {
    if (!el) {
      return false;
    }
    const style = window.getComputedStyle?.(el);
    if (style && (style.display === "none" || style.visibility === "hidden")) {
      return false;
    }
    const rect = el.getBoundingClientRect?.();
    if (!rect) {
      return false;
    }
    return rect.width > 0 && rect.height > 0;
  };

  const toLower = (value) => String(value || "").trim().toLowerCase();

  const isOrderId = (value) => ORDER_ID_PATTERN.test(String(value || "").trim());

  const getSearchType = (guide) => {
    const trimmed = String(guide || "").trim();
    if (isOrderId(trimmed)) {
      return "order";
    }
    const digits = trimmed.replace(/[^\d]/g, "");
    if (TRACKING_ID_PATTERN.test(digits)) {
      return "tracking";
    }
    return trimmed.includes("-") ? "order" : "tracking";
  };

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
    const inputs = deepQueryAll(SEARCH_INPUT_SELECTORS).filter(
      (el) => el && !isExcludedInput(el)
    );
    if (!inputs.length) {
      return null;
    }

    const visibleInputs = inputs.filter((el) => isVisible(el));
    const candidates = visibleInputs.length ? visibleInputs : inputs;

    const filterSelect = findFilterSelect();
    if (filterSelect) {
      const scope = filterSelect.closest("form, div, section") || filterSelect.parentElement;
      if (scope) {
        const scopedInput = candidates.find((input) => scope.contains(input));
        if (scopedInput) {
          return scopedInput;
        }
      }
    }

    const scored = candidates.map((input) => {
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

  const findBackToOrdersButton = () => {
    const candidates = Array.from(
      document.querySelectorAll(
        'a, button, [role="button"], input[type="button"], input[type="submit"]'
      )
    );
    return (
      candidates.find((el) => {
        if (!isVisible(el)) {
          return false;
        }
        const text = toLower(el.textContent || el.value || "");
        return BACK_TO_LIST_LABELS.some((label) => text.includes(label));
      }) || null
    );
  };

  const waitForSearchInput = async (timeoutMs = 20000) => {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      const input = findSearchInput();
      if (input) {
        return input;
      }
      await sleep(300);
    }
    return null;
  };

  const ensureSearchInput = async () => {
    if (isOrderDetailsPage()) {
      const backButton = findBackToOrdersButton();
      if (backButton) {
        backButton.click();
        return waitForSearchInput();
      }
    }
    const input = findSearchInput();
    if (input) {
      return input;
    }
    const backButton = findBackToOrdersButton();
    if (!backButton) {
      return null;
    }
    backButton.click();
    return waitForSearchInput();
  };

  const isOrderDetailsPage = () => {
    const bodyText = toLower(document.body?.textContent || "");
    return ORDER_DETAILS_INDICATORS.some((label) => bodyText.includes(label));
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
    let lastRows = initialRows;
    let stableCount = 0;
    const guideLower = toLower(guide);
    while (Date.now() - start < timeoutMs) {
      if (hasNoResults()) {
        return "empty";
      }
      if (isOrderDetailsPage()) {
        return "rows";
      }
      const rows = getOrderRows();
      if (rows.length) {
        const currentRows = getRowsSnapshot();
        if (currentRows) {
          if (currentRows === lastRows) {
            stableCount += 1;
          } else {
            stableCount = 0;
            lastRows = currentRows;
          }
          if (stableCount >= 2) {
            if (currentRows !== initialRows) {
              return "rows";
            }
            const input = findSearchInput();
            const inputValue = toLower(input?.value || "");
            if (guideLower && inputValue.includes(guideLower)) {
              return "rows";
            }
            if (Date.now() - start > SEARCH_DELAY_MS * 2) {
              return "rows";
            }
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

  const sanitizeTrackingToken = (token) => {
    const raw = String(token || "").trim();
    if (!raw) {
      return "";
    }
    let trimmed = raw;
    const lower = trimmed.toLowerCase();
    const cutWords = ["delivery", "delivered", "entrega", "fecha", "date"];
    cutWords.forEach((word) => {
      const index = lower.indexOf(word);
      if (index > 0) {
        trimmed = trimmed.slice(0, index);
      }
    });
    const digitMatch = trimmed.match(/^(\d{8,14})/);
    if (digitMatch) {
      return digitMatch[1];
    }
    return trimmed.replace(/[^A-Z0-9-]/gi, "");
  };

  const extractTrackingFromText = (text) => {
    const match = String(text || "").match(
      /(?:Tracking\s*ID|ID\s*de\s*seguimiento|ID\s*de\s*rastreo|Guia|Guía)\s*[:#]?\s*([A-Z0-9-]{6,})/i
    );
    return match ? sanitizeTrackingToken(match[1]) : "";
  };

  const extractTrackingFromDetails = () => {
    const labelSet = [
      "tracking id",
      "id de seguimiento",
      "id de rastreo"
    ];
    const candidates = Array.from(
      document.querySelectorAll("span, div, dt, dd, th, td, label, strong, b, a")
    );
    for (const el of candidates) {
      const text = toLower(el.textContent || "");
      if (!text || !labelSet.some((label) => text.includes(label))) {
        continue;
      }
      const combined = [
        el.textContent || "",
        el.nextElementSibling?.textContent || "",
        el.parentElement?.textContent || ""
      ].join(" ");
      const tracking = extractTrackingFromText(combined);
      if (tracking) {
        return tracking;
      }
    }
    return "";
  };

  const findSectionHeading = (labels) => {
    const labelSet = labels.map((label) => toLower(label));
    const headingCandidates = Array.from(
      document.querySelectorAll("h1, h2, h3, h4, h5, [role=\"heading\"]")
    );
    const match = headingCandidates.find((el) => {
      const text = toLower(el.textContent || "");
      return text && labelSet.some((label) => text.includes(label));
    });
    if (match) {
      return match;
    }
    const fallbackCandidates = Array.from(document.querySelectorAll("section, div, span"));
    return (
      fallbackCandidates.find((el) => {
        if (el.children.length > 0) {
          return false;
        }
        const text = toLower(el.textContent || "");
        if (!text || text.length > 60) {
          return false;
        }
        return text && labelSet.some((label) => text.includes(label));
      }) || null
    );
  };

  const findTableAfterNode = (node) => {
    let current = node;
    for (let depth = 0; depth < 4 && current; depth += 1) {
      let sibling = current.nextElementSibling;
      while (sibling) {
        if (sibling.matches?.("table")) {
          return sibling;
        }
        const nestedTable = sibling.querySelector?.("table");
        if (nestedTable) {
          return nestedTable;
        }
        sibling = sibling.nextElementSibling;
      }
      current = current.parentElement;
    }
    return null;
  };

  const tableHasLabelNearby = (table, labels) => {
    const labelSet = labels.map((label) => toLower(label));
    let node = table;
    for (let depth = 0; depth < 4 && node; depth += 1) {
      let sibling = node.previousElementSibling;
      while (sibling) {
        const text = toLower(sibling.textContent || "");
        if (labelSet.some((label) => text.includes(label))) {
          return true;
        }
        sibling = sibling.previousElementSibling;
      }
      node = node.parentElement;
    }
    return false;
  };

  const getHeaderInfo = (table) => {
    const theadRow = table.querySelector("thead tr");
    if (theadRow) {
      const cells = Array.from(theadRow.querySelectorAll("th, td"));
      if (cells.length) {
        return { cells, row: theadRow };
      }
    }
    const rows = Array.from(table.querySelectorAll("tr"));
    for (let i = 0; i < rows.length; i += 1) {
      const cells = Array.from(rows[i].querySelectorAll("th, td"));
      if (!cells.length) {
        continue;
      }
      const rowText = toLower(cells.map((cell) => cell.textContent || "").join(" "));
      const hasQuantity = /quantity|cantidad/.test(rowText);
      const hasProduct = /product|producto|product name|asin/.test(rowText);
      if (hasQuantity && hasProduct) {
        return { cells, row: rows[i] };
      }
    }
    return null;
  };

  const extractQuantityFromCell = (cell) => {
    const match = String(cell?.textContent || "").match(/\d+/);
    return match ? Number.parseInt(match[0], 10) : null;
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

  const mergeItemsByAsin = (items, tracking) => {
    const merged = new Map();
    items.forEach((item) => {
      if (!item.asin) {
        return;
      }
      if (!merged.has(item.asin)) {
        merged.set(item.asin, {
          ...item,
          tracking: item.tracking || tracking
        });
        return;
      }
      const existing = merged.get(item.asin);
      if (!existing) {
        return;
      }
      if (item.quantity && existing.quantity) {
        existing.quantity += item.quantity;
      }
    });
    return Array.from(merged.values());
  };

  const extractItemsFromOrderDetails = (guide) => {
    const heading = findSectionHeading(ORDER_CONTENTS_LABELS);
    const preferredTable = heading ? findTableAfterNode(heading) : null;
    const tables = Array.from(document.querySelectorAll("table"));
    const candidates = tables
      .map((table) => {
        const asins = extractAsinsFromText(table.textContent || "");
        if (!asins.length) {
          return null;
        }
        const headerInfo = getHeaderInfo(table);
        const headerText = toLower(
          headerInfo?.cells.map((cell) => cell.textContent || "").join(" ") || ""
        );
        const hasQuantity = /quantity|cantidad/.test(headerText);
        const hasProduct = /product|producto|product name|asin/.test(headerText);
        let score = asins.length;
        if (hasQuantity) {
          score += 3;
        }
        if (hasProduct) {
          score += 2;
        }
        if (headerInfo) {
          score += 2;
        }
        if (tableHasLabelNearby(table, ORDER_CONTENTS_LABELS)) {
          score += 5;
        }
        if (table === preferredTable) {
          score += 8;
        }
        return { table, headerInfo, score };
      })
      .filter(Boolean)
      .sort((a, b) => b.score - a.score);

    if (!candidates.length) {
      return { items: [], tracking: "" };
    }

    const target = candidates[0];
    const items = [];
    const headerLabels = target.headerInfo?.cells.map((cell) => toLower(cell.textContent || "")) || [];
    const qtyIndex = headerLabels.findIndex(
      (label) => label.includes("quantity") || label.includes("cantidad")
    );
    const bodyRows = Array.from(target.table.querySelectorAll("tbody tr"));
    let rows =
      bodyRows.length > 0 ? bodyRows : Array.from(target.table.querySelectorAll("tr"));
    if (target.headerInfo?.row) {
      rows = rows.filter((row) => row !== target.headerInfo.row);
    }

    rows.forEach((row) => {
      const rowText = row.textContent || "";
      const asins = extractAsinsFromText(rowText);
      if (!asins.length) {
        return;
      }
      let quantity = null;
      const cells = Array.from(row.querySelectorAll("td, th"));
      if (qtyIndex >= 0 && cells[qtyIndex]) {
        quantity = extractQuantityFromCell(cells[qtyIndex]);
      }
      if (!quantity) {
        const qtyMatches = extractQuantityMatches(rowText);
        if (qtyMatches.length === 1) {
          quantity = qtyMatches[0];
        }
      }
      asins.forEach((asin) => {
        items.push({
          asin,
          quantity,
          tracking: ""
        });
      });
    });

    const tracking =
      extractTrackingFromDetails() ||
      extractTrackingFromText(document.body?.textContent || "") ||
      guide;
    return { items: mergeItemsByAsin(items, tracking), tracking };
  };

  const extractItemsFromPage = (guide, searchType) => {
    if (searchType === "order") {
      return extractItemsFromOrderDetails(guide);
    }

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
      return { items: mergeItemsByAsin(items, tracking), tracking };
    }

    return parseItemsFromText(document.body?.textContent || "", guide);
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
    const input = await ensureSearchInput();
    if (!input) {
      return { guide, status: "error", note: "No se encontro el buscador." };
    }
    const searchType = getSearchType(guide);
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
    await sleep(SEARCH_DELAY_MS);
    const status = await waitForResults(guide);
    if (status === "timeout") {
      return { guide, status: "error", note: "Tiempo de espera agotado." };
    }
    if (status === "empty") {
      return { guide, status: "not_found", items: [] };
    }
    await sleep(SEARCH_DELAY_MS);
    const extracted = extractItemsFromPage(guide, searchType);
    if (!extracted.items || !extracted.items.length) {
      return { guide, status: "not_found", items: [] };
    }
    const result = {
      guide,
      status: "found",
      items: extracted.items,
      tracking: extracted.tracking || guide
    };
    if (searchType === "order") {
      await sleep(200);
      const backInput = await ensureSearchInput();
      if (!backInput) {
        result.note = "No se pudo volver a la lista de pedidos.";
      }
    }
    return result;
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
      if (lookupInProgress) {
        sendResponse({
          ok: false,
          error: "Ya hay una busqueda en curso."
        });
        return;
      }
      const guides = Array.isArray(message.guides) ? message.guides : [];
      lookupInProgress = true;
      lookupGuides(guides)
        .then((payload) => {
          chrome.runtime.sendMessage({
            type: "LOOKUP_COMPLETE",
            payload
          });
          sendResponse({ ok: true });
        })
        .catch((error) => {
          chrome.runtime.sendMessage({
            type: "LOOKUP_ERROR",
            payload: { error: error?.message || "Error en la busqueda." }
          });
          sendResponse({ ok: false, error: "Error en la busqueda." });
        })
        .finally(() => {
          lookupInProgress = false;
        });
      return true;
    }
  });
})();
