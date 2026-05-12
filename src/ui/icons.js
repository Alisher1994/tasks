const ATTRS = 'viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="lucide-icon" aria-hidden="true"';

const ICONS = {
  "layout-dashboard": `<svg ${ATTRS}><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/></svg>`,
  user: `<svg ${ATTRS}><circle cx="12" cy="8" r="4"/><path d="M6 20c1.6-2.7 4-4 6-4s4.4 1.3 6 4"/></svg>`,
  panelLeft: `<svg ${ATTRS}><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M9 3v18"/></svg>`,
  listChecks: `<svg ${ATTRS}><path d="M9 6h11"/><path d="M9 12h11"/><path d="M9 18h11"/><path d="m3 6 1 1 2-2"/><path d="m3 12 1 1 2-2"/><path d="m3 18 1 1 2-2"/></svg>`,
  users: `<svg ${ATTRS}><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>`,
  userCheck: `<svg ${ATTRS}><path d="m16 11 2 2 4-4"/><path d="M8 7a4 4 0 1 1 8 0 4 4 0 0 1-8 0"/><path d="M6 21v-2a6 6 0 0 1 9-5"/></svg>`,
  "check-circle-2": `<svg ${ATTRS}><circle cx="12" cy="12" r="10"/><path d="m9 12 2 2 4-4"/></svg>`,
  database: `<svg ${ATTRS}><ellipse cx="12" cy="5" rx="9" ry="3"/><path d="M3 5v6c0 1.7 4 3 9 3s9-1.3 9-3V5"/><path d="M3 11v6c0 1.7 4 3 9 3s9-1.3 9-3v-6"/></svg>`,
  "rotate-ccw": `<svg ${ATTRS}><path d="M3 2v6h6"/><path d="M3 8a9 9 0 1 0 2.6-4.6L3 6"/></svg>`,
  history: `<svg ${ATTRS}><path d="M3 12a9 9 0 1 0 3-6.7"/><path d="M3 3v6h6"/><path d="M12 7v5l3 2"/></svg>`,
  pencil: `<svg ${ATTRS}><path d="M12 20h9"/><path d="m16.5 3.5 4 4L7 21l-4 1 1-4Z"/></svg>`,
  "trash-2": `<svg ${ATTRS}><path d="M3 6h18"/><path d="M8 6V4h8v2"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6"/><path d="M10 11v6"/><path d="M14 11v6"/></svg>`,
  building2: `<svg ${ATTRS}><path d="M6 22V4a1 1 0 0 1 1-1h10a1 1 0 0 1 1 1v18"/><path d="M2 22h20"/><path d="M10 6h4"/><path d="M10 10h4"/><path d="M10 14h4"/><path d="M10 18h4"/></svg>`,
  settings: `<svg ${ATTRS}><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.6 1.6 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06A1.6 1.6 0 0 0 15 19.4a1.6 1.6 0 0 0-1 .6 1.6 1.6 0 0 0-.4 1V21a2 2 0 1 1-4 0v-.1a1.6 1.6 0 0 0-.4-1 1.6 1.6 0 0 0-1-.6 1.6 1.6 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06A1.6 1.6 0 0 0 4.6 15a1.6 1.6 0 0 0-.6-1 1.6 1.6 0 0 0-1-.4H3a2 2 0 1 1 0-4h.1a1.6 1.6 0 0 0 1-.4 1.6 1.6 0 0 0 .6-1 1.6 1.6 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06A1.6 1.6 0 0 0 9 4.6a1.6 1.6 0 0 0 1-.6 1.6 1.6 0 0 0 .4-1V3a2 2 0 1 1 4 0v.1a1.6 1.6 0 0 0 .4 1 1.6 1.6 0 0 0 1 .6 1.6 1.6 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06A1.6 1.6 0 0 0 19.4 9c.28.3.47.67.54 1.08.07.41.02.83-.14 1.22.16.39.21.81.14 1.22A1.6 1.6 0 0 0 19.4 15Z"/></svg>`,
  briefcase: `<svg ${ATTRS}><rect x="2" y="7" width="20" height="14" rx="2"/><path d="M16 7V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v2"/><path d="M2 13h20"/></svg>`,
  barChart: `<svg ${ATTRS}><path d="M3 3v18h18"/><path d="M7 16V8"/><path d="M12 16v-5"/><path d="M17 16V5"/></svg>`
};

export function iconSvg(name) {
  return ICONS[name] || "";
}

export function withIcon(iconName, text) {
  return `<span class="icon-label">${iconSvg(iconName)}<span>${text}</span></span>`;
}

export function withLucideIcon(iconName, text) {
  return `<span class="icon-label"><i data-lucide="${iconName}" class="lucide-icon" aria-hidden="true"></i><span>${text}</span></span>`;
}

export function initLucideIcons() {
  if (window.lucide && typeof window.lucide.createIcons === "function") {
    window.lucide.createIcons();
  }
}
