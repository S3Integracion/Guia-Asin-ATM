(() => {
  const ready = () => {
    chrome.runtime.sendMessage({
      type: "LOG_EVENT",
      level: "info",
      source: "amazon_seller",
      message: "Amazon Seller content script listo."
    });
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", ready, { once: true });
  } else {
    ready();
  }
})();
