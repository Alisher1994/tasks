import { escapeHtmlText, escapeRegexLiteral } from "../utils/escape.js";

const TEMPLATE_EDITOR_TEXTAREA_SELECTOR = ".task-message-template-input, .reminder-text-input, #telegramCloseAcceptedInput";

const templateEditorBySource = new WeakMap();
const templateSourceByEditor = new WeakMap();
const templateSelectionByEditor = new WeakMap();
let activeTemplateEditorSource = null;
let draggingTemplateTokenChip = null;
let draggingTemplateTokenSource = null;
let templateTokenList = [];

export function configureTemplateTokens(tokens) {
  templateTokenList = Array.from(
    new Set((tokens || []).map((t) => String(t || "").trim()).filter(Boolean))
  ).sort((a, b) => b.length - a.length);
}

function getTemplateTokenRegex() {
  if (!templateTokenList.length) return null;
  return new RegExp(`(${templateTokenList.map((token) => escapeRegexLiteral(token)).join("|")})`, "g");
}

function splitTemplateTextByTokens(text) {
  const src = String(text || "");
  const re = getTemplateTokenRegex();
  if (!re) return [{ type: "text", value: src }];
  const out = [];
  let last = 0;
  let m;
  while ((m = re.exec(src)) !== null) {
    if (m.index > last) {
      out.push({ type: "text", value: src.slice(last, m.index) });
    }
    out.push({ type: "token", value: m[0] });
    last = m.index + m[0].length;
  }
  if (last < src.length) {
    out.push({ type: "text", value: src.slice(last) });
  }
  if (!out.length) out.push({ type: "text", value: "" });
  return out;
}

export function formatTemplateTokenDisplay(token) {
  return String(token || "")
    .replace(/^\[/, "")
    .replace(/\]$/, "")
    .replace(/^\{/, "")
    .replace(/\}$/, "");
}

function createTemplateTokenChip(token) {
  const displayToken = formatTemplateTokenDisplay(token);
  const chip = document.createElement("span");
  chip.className = "template-token-chip";
  chip.setAttribute("contenteditable", "false");
  chip.setAttribute("draggable", "true");
  chip.dataset.token = String(token || "");
  chip.innerHTML = `
    <span class="template-token-chip__label">${escapeHtmlText(displayToken)}</span>
    <button type="button" class="template-token-chip__remove" aria-label="Удалить тег" title="Удалить тег">×</button>
  `;
  return chip;
}

function extractTemplateEditorValue(editor) {
  if (!editor) return "";
  let out = "";
  const walk = (node) => {
    if (!node) return;
    if (node.nodeType === Node.TEXT_NODE) {
      out += node.nodeValue || "";
      return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return;
    const el = /** @type {HTMLElement} */ (node);
    if (el.classList?.contains("template-token-chip")) {
      out += String(el.dataset.token || "");
      return;
    }
    if (el.tagName === "BR") {
      out += "\n";
      return;
    }
    const tag = el.tagName;
    const isBlock = tag === "DIV" || tag === "P";
    const children = Array.from(el.childNodes);
    children.forEach((child) => walk(child));
    if (isBlock && !out.endsWith("\n")) {
      out += "\n";
    }
  };
  Array.from(editor.childNodes).forEach((node) => walk(node));
  return out.replace(/\u200B/g, "");
}

function renderTemplateEditorFromValue(editor, value) {
  if (!editor) return;
  const prevScroll = editor.scrollTop;
  editor.innerHTML = "";
  const parts = splitTemplateTextByTokens(value);
  parts.forEach((part) => {
    if (part.type === "token") {
      editor.appendChild(createTemplateTokenChip(part.value));
      return;
    }
    editor.appendChild(document.createTextNode(part.value));
  });
  editor.scrollTop = prevScroll;
}

function syncTemplateEditorToSource(editor, source, { normalize = false } = {}) {
  if (!editor || !source) return;
  const value = extractTemplateEditorValue(editor);
  if (normalize) {
    renderTemplateEditorFromValue(editor, value);
  }
  source.value = value;
  source.dispatchEvent(new Event("input", { bubbles: true }));
}

function placeCaretAfterNode(node) {
  const sel = window.getSelection?.();
  if (!sel || !node) return;
  const range = document.createRange();
  range.setStartAfter(node);
  range.collapse(true);
  sel.removeAllRanges();
  sel.addRange(range);
}

function captureTemplateEditorSelection(editor) {
  const sel = window.getSelection?.();
  if (!sel || sel.rangeCount === 0 || !editor) return;
  const range = sel.getRangeAt(0);
  if (!editor.contains(range.commonAncestorContainer)) return;
  templateSelectionByEditor.set(editor, range.cloneRange());
}

function restoreTemplateEditorSelection(editor) {
  if (!editor) return false;
  const sel = window.getSelection?.();
  if (!sel) return false;
  const saved = templateSelectionByEditor.get(editor);
  if (saved && editor.contains(saved.commonAncestorContainer)) {
    sel.removeAllRanges();
    sel.addRange(saved);
    return true;
  }
  const range = document.createRange();
  range.selectNodeContents(editor);
  range.collapse(false);
  sel.removeAllRanges();
  sel.addRange(range);
  templateSelectionByEditor.set(editor, range.cloneRange());
  return true;
}

function getCaretRangeFromPoint(clientX, clientY, editor) {
  if (!editor) return null;
  let range = null;
  if (document.caretRangeFromPoint) {
    range = document.caretRangeFromPoint(clientX, clientY);
  } else if (document.caretPositionFromPoint) {
    const pos = document.caretPositionFromPoint(clientX, clientY);
    if (pos) {
      range = document.createRange();
      range.setStart(pos.offsetNode, pos.offset);
      range.collapse(true);
    }
  }
  if (!range) return null;
  if (!editor.contains(range.commonAncestorContainer)) return null;
  const chipAncestor =
    range.startContainer instanceof Element
      ? range.startContainer.closest(".template-token-chip")
      : range.startContainer.parentElement?.closest?.(".template-token-chip");
  if (chipAncestor instanceof HTMLElement && editor.contains(chipAncestor)) {
    const rect = chipAncestor.getBoundingClientRect();
    const placeBefore = clientX < rect.left + rect.width / 2;
    range = document.createRange();
    if (placeBefore) range.setStartBefore(chipAncestor);
    else range.setStartAfter(chipAncestor);
    range.collapse(true);
  }
  return range;
}

function insertTokenChipAtCaret(editor, token) {
  if (!editor || !token) return false;
  editor.focus();
  const sel = window.getSelection?.();
  if (!sel) return false;
  if (!sel.rangeCount || !editor.contains(sel.getRangeAt(0).commonAncestorContainer)) {
    restoreTemplateEditorSelection(editor);
  }
  if (!sel.rangeCount) return false;
  const range = sel.getRangeAt(0);
  if (!editor.contains(range.commonAncestorContainer)) return false;
  range.deleteContents();
  const chip = createTemplateTokenChip(token);
  range.insertNode(chip);
  placeCaretAfterNode(chip);
  captureTemplateEditorSelection(editor);
  return true;
}

export function insertTokenIntoTemplateSource(source, token) {
  if (!source || !token) return false;
  const editor = templateEditorBySource.get(source);
  if (editor && insertTokenChipAtCaret(editor, token)) {
    syncTemplateEditorToSource(editor, source);
    return true;
  }
  const start = typeof source.selectionStart === "number" ? source.selectionStart : source.value.length;
  const end = typeof source.selectionEnd === "number" ? source.selectionEnd : source.value.length;
  const v = source.value;
  source.value = `${v.slice(0, start)}${token}${v.slice(end)}`;
  const pos = start + token.length;
  source.focus();
  if (typeof source.setSelectionRange === "function") source.setSelectionRange(pos, pos);
  source.dispatchEvent(new Event("input", { bubbles: true }));
  return true;
}

export function resolveTemplateEditorSource(selector) {
  const matchesSelector = (el) => el instanceof HTMLTextAreaElement && (!selector || el.matches(selector));

  if (matchesSelector(activeTemplateEditorSource)) {
    return activeTemplateEditorSource;
  }

  const sel = window.getSelection?.();
  const anchorNode = sel && sel.rangeCount > 0 ? sel.anchorNode : null;
  const anchorEl =
    anchorNode instanceof Element
      ? anchorNode
      : anchorNode && anchorNode.parentElement instanceof Element
        ? anchorNode.parentElement
        : null;
  const editorFromSelection = anchorEl ? anchorEl.closest(".template-token-editor") : null;
  if (editorFromSelection instanceof HTMLElement) {
    const source = templateSourceByEditor.get(editorFromSelection);
    if (matchesSelector(source)) {
      activeTemplateEditorSource = source;
      return source;
    }
  }

  const focused = document.activeElement;
  if (focused instanceof HTMLElement) {
    const focusedEditor = focused.closest(".template-token-editor");
    if (focusedEditor instanceof HTMLElement) {
      const source = templateSourceByEditor.get(focusedEditor);
      if (matchesSelector(source)) {
        activeTemplateEditorSource = source;
        return source;
      }
    }
  }

  const fallback = document.querySelector(selector || TEMPLATE_EDITOR_TEXTAREA_SELECTOR);
  return fallback instanceof HTMLTextAreaElement ? fallback : null;
}

export function initTemplateTokenEditors() {
  const sources = Array.from(document.querySelectorAll(TEMPLATE_EDITOR_TEXTAREA_SELECTOR));
  sources.forEach((source) => {
    if (!(source instanceof HTMLTextAreaElement)) return;
    if (templateEditorBySource.has(source)) return;
    const editor = document.createElement("div");
    editor.className = "template-token-editor";
    editor.setAttribute("contenteditable", "true");
    editor.setAttribute("spellcheck", "false");
    editor.dataset.sourceFor = source.id || "";
    renderTemplateEditorFromValue(editor, source.value);
    source.classList.add("template-token-source");
    source.setAttribute("aria-hidden", "true");
    source.tabIndex = -1;
    source.style.display = "none";
    source.insertAdjacentElement("afterend", editor);
    templateEditorBySource.set(source, editor);
    templateSourceByEditor.set(editor, source);

    editor.addEventListener("focus", () => {
      activeTemplateEditorSource = source;
      captureTemplateEditorSelection(editor);
    });
    editor.addEventListener("click", (event) => {
      const removeBtn = event.target instanceof HTMLElement ? event.target.closest(".template-token-chip__remove") : null;
      if (removeBtn) {
        const chip = removeBtn.closest(".template-token-chip");
        chip?.remove();
        syncTemplateEditorToSource(editor, source);
        captureTemplateEditorSelection(editor);
        return;
      }
      activeTemplateEditorSource = source;
      captureTemplateEditorSelection(editor);
    });
    editor.addEventListener("keyup", () => captureTemplateEditorSelection(editor));
    editor.addEventListener("mouseup", () => captureTemplateEditorSelection(editor));
    editor.addEventListener("beforeinput", (event) => {
      if (event.inputType === "insertParagraph") {
        event.preventDefault();
        document.execCommand("insertText", false, "\n");
      }
    });
    editor.addEventListener("dragstart", (event) => {
      const target = event.target instanceof HTMLElement ? event.target.closest(".template-token-chip") : null;
      if (!(target instanceof HTMLElement)) return;
      draggingTemplateTokenChip = target;
      draggingTemplateTokenSource = source;
      target.classList.add("is-dragging");
      if (event.dataTransfer) {
        event.dataTransfer.effectAllowed = "move";
        event.dataTransfer.setData("text/plain", String(target.dataset.token || ""));
      }
    });
    editor.addEventListener("dragend", () => {
      if (draggingTemplateTokenChip instanceof HTMLElement) {
        draggingTemplateTokenChip.classList.remove("is-dragging");
      }
      editor.querySelectorAll(".template-token-chip.drop-before, .template-token-chip.drop-after").forEach((el) => {
        el.classList.remove("drop-before", "drop-after");
      });
      draggingTemplateTokenChip = null;
      draggingTemplateTokenSource = null;
    });
    editor.addEventListener("dragover", (event) => {
      if (!(draggingTemplateTokenChip instanceof HTMLElement)) return;
      if (draggingTemplateTokenSource !== source) return;
      event.preventDefault();
      if (event.dataTransfer) event.dataTransfer.dropEffect = "move";
    });
    editor.addEventListener("drop", (event) => {
      if (!(draggingTemplateTokenChip instanceof HTMLElement)) return;
      if (draggingTemplateTokenSource !== source) return;
      event.preventDefault();
      editor.querySelectorAll(".template-token-chip.drop-before, .template-token-chip.drop-after").forEach((el) => {
        el.classList.remove("drop-before", "drop-after");
      });
      let dropRange = getCaretRangeFromPoint(event.clientX, event.clientY, editor);
      if (!dropRange) {
        restoreTemplateEditorSelection(editor);
        const sel = window.getSelection?.();
        dropRange = sel && sel.rangeCount ? sel.getRangeAt(0).cloneRange() : null;
      }
      if (!dropRange || !editor.contains(dropRange.commonAncestorContainer)) {
        dropRange = document.createRange();
        dropRange.selectNodeContents(editor);
        dropRange.collapse(false);
      }
      draggingTemplateTokenChip.remove();
      dropRange.insertNode(draggingTemplateTokenChip);
      placeCaretAfterNode(draggingTemplateTokenChip);
      captureTemplateEditorSelection(editor);
      syncTemplateEditorToSource(editor, source);
    });
    editor.addEventListener("input", () => {
      activeTemplateEditorSource = source;
      captureTemplateEditorSelection(editor);
      syncTemplateEditorToSource(editor, source);
    });
    editor.addEventListener("blur", () => {
      syncTemplateEditorToSource(editor, source, { normalize: true });
      captureTemplateEditorSelection(editor);
    });
  });
}
