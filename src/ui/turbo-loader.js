const turboTopLoader = document.getElementById("turboTopLoader");
const turboTopLoaderBar = document.getElementById("turboTopLoaderBar");

let turboLoaderProgress = 0;
let turboLoaderProgressTimer = null;
let turboLoaderHideTimer = null;
let getActiveSectionId = () => "";

export function configureTurboLoader({ getActiveSection } = {}) {
  if (typeof getActiveSection === "function") {
    getActiveSectionId = getActiveSection;
  }
}

function setTurboLoaderProgress(progress) {
  if (!turboTopLoaderBar) return;
  const safe = Math.max(0, Math.min(100, Number(progress) || 0));
  turboTopLoaderBar.style.width = `${safe}%`;
}

export function startTurboLoader() {
  if (!turboTopLoader || !turboTopLoaderBar) return;
  if (document.body.classList.contains("app-booting")) return;
  clearTimeout(turboLoaderHideTimer);
  turboTopLoader.classList.add("is-active");
  if (turboLoaderProgress <= 0 || turboLoaderProgress >= 100) {
    turboLoaderProgress = 8;
    setTurboLoaderProgress(turboLoaderProgress);
  }
  clearInterval(turboLoaderProgressTimer);
  turboLoaderProgressTimer = setInterval(() => {
    turboLoaderProgress = Math.min(90, turboLoaderProgress + (turboLoaderProgress < 45 ? 12 : turboLoaderProgress < 75 ? 6 : 2));
    setTurboLoaderProgress(turboLoaderProgress);
    if (turboLoaderProgress >= 90) {
      clearInterval(turboLoaderProgressTimer);
      turboLoaderProgressTimer = null;
    }
  }, 90);
}

export function finishTurboLoader() {
  if (!turboTopLoader || !turboTopLoaderBar) return;
  clearInterval(turboLoaderProgressTimer);
  turboLoaderProgressTimer = null;
  turboLoaderProgress = 100;
  setTurboLoaderProgress(100);
  clearTimeout(turboLoaderHideTimer);
  turboLoaderHideTimer = setTimeout(() => {
    turboTopLoader.classList.remove("is-active");
    turboLoaderProgress = 0;
    setTurboLoaderProgress(0);
  }, 170);
}

export function shouldUseTurboLoader(sectionId = getActiveSectionId()) {
  return String(sectionId || "") !== "tasks";
}

export function pulseTurboLoader() {
  if (!shouldUseTurboLoader()) return;
  startTurboLoader();
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      finishTurboLoader();
    });
  });
}
