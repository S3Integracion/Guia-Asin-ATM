(() => {
  const ready = () => {
    chrome.runtime.sendMessage({
      type: "LOG_EVENT",
      level: "info",
      source: "excel_online",
      message: "Excel Online content script listo."
    });
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", ready, { once: true });
  } else {
    ready();
  }
})();
