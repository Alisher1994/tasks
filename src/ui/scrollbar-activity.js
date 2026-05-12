let scrollbarActivityTimer = null;

function markScrollbarActivity() {
  document.documentElement.classList.add("scrollbars-active");
  clearTimeout(scrollbarActivityTimer);
  scrollbarActivityTimer = setTimeout(() => {
    document.documentElement.classList.remove("scrollbars-active");
  }, 900);
}

export function initGlobalScrollbarActivityTracker() {
  const opts = { passive: true, capture: true };
  ["scroll", "wheel", "touchmove"].forEach((eventName) => {
    window.addEventListener(eventName, markScrollbarActivity, opts);
  });
}
