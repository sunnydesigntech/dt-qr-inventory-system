// Minimal inline SVG icon set — copy-pasteable, no dependencies
// Per design system: minimal — only status dots + QR/scan glyph + a small functional set
window.Icon = ({ name, size = 18, className = '', stroke = 1.75 }) => {
  const s = { width: size, height: size, fill: 'none', stroke: 'currentColor', strokeWidth: stroke, strokeLinecap: 'round', strokeLinejoin: 'round', className };
  switch (name) {
    case 'qr': return <svg {...s} viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1.5"/><rect x="14" y="3" width="7" height="7" rx="1.5"/><rect x="3" y="14" width="7" height="7" rx="1.5"/><path d="M14 14h3v3h-3zM20 14v3M14 20h3M20 20v.01"/></svg>;
    case 'scan': return <svg {...s} viewBox="0 0 24 24"><path d="M3 7V5a2 2 0 0 1 2-2h2M17 3h2a2 2 0 0 1 2 2v2M21 17v2a2 2 0 0 1-2 2h-2M7 21H5a2 2 0 0 1-2-2v-2M7 12h10"/></svg>;
    case 'search': return <svg {...s} viewBox="0 0 24 24"><circle cx="11" cy="11" r="7"/><path d="m21 21-4.3-4.3"/></svg>;
    case 'filter': return <svg {...s} viewBox="0 0 24 24"><path d="M3 5h18M6 12h12M10 19h4"/></svg>;
    case 'edit': return <svg {...s} viewBox="0 0 24 24"><path d="M12 20h9M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4Z"/></svg>;
    case 'check': return <svg {...s} viewBox="0 0 24 24"><path d="m5 12 5 5L20 7"/></svg>;
    case 'x': return <svg {...s} viewBox="0 0 24 24"><path d="M18 6 6 18M6 6l12 12"/></svg>;
    case 'warn': return <svg {...s} viewBox="0 0 24 24"><path d="M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0ZM12 9v4M12 17h.01"/></svg>;
    case 'hazard': return <svg {...s} viewBox="0 0 24 24"><path d="M12 3 22 20H2Z"/><path d="M12 10v4M12 18h.01"/></svg>;
    case 'chevron': return <svg {...s} viewBox="0 0 24 24"><path d="m9 18 6-6-6-6"/></svg>;
    case 'chevron-down': return <svg {...s} viewBox="0 0 24 24"><path d="m6 9 6 6 6-6"/></svg>;
    case 'arrow-left': return <svg {...s} viewBox="0 0 24 24"><path d="M19 12H5M12 19l-7-7 7-7"/></svg>;
    case 'plus': return <svg {...s} viewBox="0 0 24 24"><path d="M12 5v14M5 12h14"/></svg>;
    case 'minus': return <svg {...s} viewBox="0 0 24 24"><path d="M5 12h14"/></svg>;
    case 'box': return <svg {...s} viewBox="0 0 24 24"><path d="M21 8 12 3 3 8v8l9 5 9-5Z"/><path d="M3.3 7 12 12l8.7-5M12 22V12"/></svg>;
    case 'flask': return <svg {...s} viewBox="0 0 24 24"><path d="M9 3h6M10 3v6.5L4 19a2 2 0 0 0 1.7 3h12.6A2 2 0 0 0 20 19l-6-9.5V3"/></svg>;
    case 'wrench': return <svg {...s} viewBox="0 0 24 24"><path d="M14.7 6.3a4 4 0 1 0 4.99 4.99l-2.4-2.4 2.4-2.4-2.6-2.6-2.4 2.4ZM4 20l9-9"/></svg>;
    case 'cpu': return <svg {...s} viewBox="0 0 24 24"><rect x="4" y="4" width="16" height="16" rx="2"/><rect x="9" y="9" width="6" height="6"/><path d="M9 1v3M15 1v3M9 20v3M15 20v3M20 9h3M20 14h3M1 9h3M1 14h3"/></svg>;
    case 'printer': return <svg {...s} viewBox="0 0 24 24"><path d="M6 9V2h12v7M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/></svg>;
    case 'shield': return <svg {...s} viewBox="0 0 24 24"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10Z"/></svg>;
    case 'external': return <svg {...s} viewBox="0 0 24 24"><path d="M15 3h6v6M10 14 21 3M21 14v5a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5"/></svg>;
    case 'building': return <svg {...s} viewBox="0 0 24 24"><rect x="4" y="2" width="16" height="20" rx="1"/><path d="M9 22v-4h6v4M9 6h.01M15 6h.01M9 10h.01M15 10h.01M9 14h.01M15 14h.01"/></svg>;
    case 'list': return <svg {...s} viewBox="0 0 24 24"><path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01"/></svg>;
    case 'grid': return <svg {...s} viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/></svg>;
    case 'mobile': return <svg {...s} viewBox="0 0 24 24"><rect x="6" y="2" width="12" height="20" rx="2"/><path d="M12 18h.01"/></svg>;
    case 'desktop': return <svg {...s} viewBox="0 0 24 24"><rect x="2" y="3" width="20" height="14" rx="2"/><path d="M8 21h8M12 17v4"/></svg>;
    case 'refresh': return <svg {...s} viewBox="0 0 24 24"><path d="M3 12a9 9 0 0 1 15-6.7L21 8M21 3v5h-5M21 12a9 9 0 0 1-15 6.7L3 16M3 21v-5h5"/></svg>;
    case 'save': return <svg {...s} viewBox="0 0 24 24"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2Z"/><path d="M17 21v-8H7v8M7 3v5h8"/></svg>;
    case 'spark': return <svg {...s} viewBox="0 0 24 24"><path d="M12 3v3M12 18v3M3 12h3M18 12h3M5.6 5.6l2.1 2.1M16.3 16.3l2.1 2.1M5.6 18.4l2.1-2.1M16.3 7.7l2.1-2.1"/></svg>;
    case 'info': return <svg {...s} viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="M12 8h.01M11 12h1v5h1"/></svg>;
    case 'arrow-right': return <svg {...s} viewBox="0 0 24 24"><path d="M5 12h14M12 5l7 7-7 7"/></svg>;
    default: return null;
  }
};
