import type { ReactElement } from 'react';
import type { DocKind } from '../types';

/**
 * Stroke-based SVG icon set (Feather-style geometry, 24x24 grid).
 * Renders crisp at any size and inherits color via currentColor.
 */
const PATHS: Record<string, ReactElement> = {
  // --- text formatting ---
  bold: (
    <>
      <path d="M6 4h8a4 4 0 0 1 0 8H6z" />
      <path d="M6 12h9a4 4 0 0 1 0 8H6z" />
    </>
  ),
  italic: (
    <>
      <line x1="19" y1="4" x2="10" y2="4" />
      <line x1="14" y1="20" x2="5" y2="20" />
      <line x1="15" y1="4" x2="9" y2="20" />
    </>
  ),
  underline: (
    <>
      <path d="M6 3v7a6 6 0 0 0 6 6 6 6 0 0 0 6-6V3" />
      <line x1="4" y1="21" x2="20" y2="21" />
    </>
  ),
  strike: (
    <>
      <path d="M16 4H9a3 3 0 0 0-2.83 4" />
      <path d="M14 12a4 4 0 0 1 0 8H6" />
      <line x1="4" y1="12" x2="20" y2="12" />
    </>
  ),
  h1: (
    <text x="12" y="16.5" textAnchor="middle" fontSize="13" fontWeight="700" fill="currentColor" stroke="none" fontFamily="inherit">
      H1
    </text>
  ),
  h2: (
    <text x="12" y="16.5" textAnchor="middle" fontSize="13" fontWeight="700" fill="currentColor" stroke="none" fontFamily="inherit">
      H2
    </text>
  ),
  listBullet: (
    <>
      <line x1="8" y1="6" x2="21" y2="6" />
      <line x1="8" y1="12" x2="21" y2="12" />
      <line x1="8" y1="18" x2="21" y2="18" />
      <circle cx="3.5" cy="6" r="1.2" fill="currentColor" stroke="none" />
      <circle cx="3.5" cy="12" r="1.2" fill="currentColor" stroke="none" />
      <circle cx="3.5" cy="18" r="1.2" fill="currentColor" stroke="none" />
    </>
  ),
  listOrdered: (
    <>
      <line x1="10" y1="6" x2="21" y2="6" />
      <line x1="10" y1="12" x2="21" y2="12" />
      <line x1="10" y1="18" x2="21" y2="18" />
      <text x="3" y="8.5" fontSize="8" fontWeight="700" fill="currentColor" stroke="none">1</text>
      <text x="3" y="14.5" fontSize="8" fontWeight="700" fill="currentColor" stroke="none">2</text>
      <text x="3" y="20.5" fontSize="8" fontWeight="700" fill="currentColor" stroke="none">3</text>
    </>
  ),
  indent: (
    <>
      <line x1="21" y1="5" x2="3" y2="5" />
      <line x1="21" y1="12" x2="11" y2="12" />
      <line x1="21" y1="19" x2="3" y2="19" />
      <polyline points="7 9 10 12 7 15" />
    </>
  ),
  outdent: (
    <>
      <line x1="21" y1="5" x2="3" y2="5" />
      <line x1="21" y1="12" x2="11" y2="12" />
      <line x1="21" y1="19" x2="3" y2="19" />
      <polyline points="10 9 7 12 10 15" />
    </>
  ),
  alignLeft: (
    <>
      <line x1="21" y1="6" x2="3" y2="6" />
      <line x1="15" y1="12" x2="3" y2="12" />
      <line x1="17" y1="18" x2="3" y2="18" />
    </>
  ),
  alignCenter: (
    <>
      <line x1="21" y1="6" x2="3" y2="6" />
      <line x1="18" y1="12" x2="6" y2="12" />
      <line x1="19" y1="18" x2="5" y2="18" />
    </>
  ),
  alignRight: (
    <>
      <line x1="21" y1="6" x2="3" y2="6" />
      <line x1="21" y1="12" x2="9" y2="12" />
      <line x1="21" y1="18" x2="7" y2="18" />
    </>
  ),
  alignJustify: (
    <>
      <line x1="21" y1="6" x2="3" y2="6" />
      <line x1="21" y1="12" x2="3" y2="12" />
      <line x1="21" y1="18" x2="3" y2="18" />
    </>
  ),
  fontColor: (
    <>
      <path d="M12 3l5.2 11.2" />
      <path d="M6.8 14.2L12 3" />
      <path d="M8.5 10.5h7" />
      <path d="M5 18.5h14" />
      <path d="M7 21.5h10" opacity="0.45" />
    </>
  ),
  highlight: (
    <>
      <path d="M12 20h9" />
      <path d="M16.5 3.5a2.12 2.12 0 0 1 3 3L8 18l-4 1 1-4z" />
      <line x1="4" y1="3" x2="9" y2="3" opacity="0.5" />
    </>
  ),
  clearFormat: (
    <>
      <path d="M7 5l10 14" />
      <path d="M17 5L7 19" />
      <path d="M5.5 8.5h6" opacity="0.6" />
    </>
  ),

  // --- actions ---
  undo: (
    <>
      <polyline points="9 14 4 9 9 4" />
      <path d="M20 20v-7a4 4 0 0 0-4-4H4" />
    </>
  ),
  redo: (
    <>
      <polyline points="15 14 20 9 15 4" />
      <path d="M4 20v-7a4 4 0 0 1 4-4h12" />
    </>
  ),
  save: (
    <>
      <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z" />
      <polyline points="17 21 17 13 7 13 7 21" />
      <polyline points="7 3 7 8 15 8" />
    </>
  ),
  share: (
    <>
      <circle cx="18" cy="5" r="3" />
      <circle cx="6" cy="12" r="3" />
      <circle cx="18" cy="19" r="3" />
      <line x1="8.6" y1="13.5" x2="15.4" y2="17.5" />
      <line x1="15.4" y1="6.5" x2="8.6" y2="10.5" />
    </>
  ),
  download: (
    <>
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
      <polyline points="7 10 12 15 17 10" />
      <line x1="12" y1="15" x2="12" y2="3" />
    </>
  ),
  upload: (
    <>
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
      <polyline points="7 8 12 3 17 8" />
      <line x1="12" y1="3" x2="12" y2="15" />
    </>
  ),
  folder: (
    <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z" />
  ),
  plus: (
    <>
      <line x1="12" y1="5" x2="12" y2="19" />
      <line x1="5" y1="12" x2="19" y2="12" />
    </>
  ),
  close: (
    <>
      <line x1="18" y1="6" x2="6" y2="18" />
      <line x1="6" y1="6" x2="18" y2="18" />
    </>
  ),
  check: <polyline points="20 6 9 17 4 12" />,
  chevronLeft: <polyline points="15 18 9 12 15 6" />,
  chevronRight: <polyline points="9 18 15 12 9 6" />,
  chevronDown: <polyline points="6 9 12 15 18 9" />,
  chevronUp: <polyline points="18 15 12 9 6 15" />,
  arrowLeft: (
    <>
      <line x1="19" y1="12" x2="5" y2="12" />
      <polyline points="12 19 5 12 12 5" />
    </>
  ),
  search: (
    <>
      <circle cx="11" cy="11" r="8" />
      <line x1="21" y1="21" x2="16.65" y2="16.65" />
    </>
  ),
  more: (
    <>
      <circle cx="12" cy="5" r="1.7" fill="currentColor" stroke="none" />
      <circle cx="12" cy="12" r="1.7" fill="currentColor" stroke="none" />
      <circle cx="12" cy="19" r="1.7" fill="currentColor" stroke="none" />
    </>
  ),
  play: <polygon points="6 3 20 12 6 21" fill="currentColor" stroke="none" />,
  stop: (
    <>
      <circle cx="12" cy="12" r="9.5" />
      <rect x="8.5" y="8.5" width="7" height="7" rx="1" fill="currentColor" stroke="none" />
    </>
  ),
  send: (
    <>
      <line x1="22" y1="2" x2="11" y2="13" />
      <polygon points="22 2 15 22 11 13 2 9 22 2" />
    </>
  ),
  copy: (
    <>
      <rect x="9" y="9" width="13" height="13" rx="2" />
      <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1" />
    </>
  ),
  trash: (
    <>
      <polyline points="3 6 5 6 21 6" />
      <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2" />
    </>
  ),
  edit: (
    <>
      <path d="M12 20h9" />
      <path d="M16.5 3.5a2.12 2.12 0 0 1 3 3L7 19l-4 1 1-4z" />
    </>
  ),

  // --- insert / objects ---
  table: (
    <>
      <rect x="3" y="3" width="18" height="18" rx="2" />
      <line x1="3" y1="9" x2="21" y2="9" />
      <line x1="3" y1="15" x2="21" y2="15" />
      <line x1="12" y1="3" x2="12" y2="21" />
    </>
  ),
  hr: (
    <>
      <line x1="3" y1="12" x2="21" y2="12" strokeWidth="2.6" />
      <line x1="7" y1="6" x2="17" y2="6" opacity="0.4" />
      <line x1="7" y1="18" x2="17" y2="18" opacity="0.4" />
    </>
  ),
  link: (
    <>
      <path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71" />
      <path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71" />
    </>
  ),
  calendar: (
    <>
      <rect x="3" y="4" width="18" height="18" rx="2" />
      <line x1="16" y1="2" x2="16" y2="6" />
      <line x1="8" y1="2" x2="8" y2="6" />
      <line x1="3" y1="10" x2="21" y2="10" />
    </>
  ),
  image: (
    <>
      <rect x="3" y="3" width="18" height="18" rx="2" />
      <circle cx="8.5" cy="8.5" r="1.5" />
      <polyline points="21 15 16 10 5 21" />
    </>
  ),
  fx: (
    <text x="12" y="16.5" textAnchor="middle" fontSize="14" fontWeight="600" fontStyle="italic" fill="currentColor" stroke="none" fontFamily="Georgia, serif">
      fx
    </text>
  ),

  // --- navigation / app ---
  home: (
    <>
      <path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z" />
      <polyline points="9 22 9 12 15 12 15 22" />
    </>
  ),
  fileText: (
    <>
      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
      <polyline points="14 2 14 8 20 8" />
      <line x1="16" y1="13" x2="8" y2="13" />
      <line x1="16" y1="17" x2="8" y2="17" />
    </>
  ),
  grid: (
    <>
      <rect x="3" y="3" width="7" height="7" rx="1" />
      <rect x="14" y="3" width="7" height="7" rx="1" />
      <rect x="14" y="14" width="7" height="7" rx="1" />
      <rect x="3" y="14" width="7" height="7" rx="1" />
    </>
  ),
  monitor: (
    <>
      <rect x="2" y="3" width="20" height="14" rx="2" />
      <line x1="8" y1="21" x2="16" y2="21" />
      <line x1="12" y1="17" x2="12" y2="21" />
    </>
  ),
  chat: (
    <path d="M21 11.5a8.38 8.38 0 0 1-.9 3.8 8.5 8.5 0 0 1-7.6 4.7 8.38 8.38 0 0 1-3.8-.9L3 21l1.9-5.7a8.38 8.38 0 0 1-.9-3.8 8.5 8.5 0 0 1 4.7-7.6 8.38 8.38 0 0 1 3.8-.9h.5a8.48 8.48 0 0 1 8 8z" />
  ),
  settings: (
    <>
      <line x1="4" y1="21" x2="4" y2="14" />
      <line x1="4" y1="10" x2="4" y2="3" />
      <line x1="12" y1="21" x2="12" y2="12" />
      <line x1="12" y1="8" x2="12" y2="3" />
      <line x1="20" y1="21" x2="20" y2="16" />
      <line x1="20" y1="12" x2="20" y2="3" />
      <line x1="1" y1="14" x2="7" y2="14" />
      <line x1="9" y1="8" x2="15" y2="8" />
      <line x1="17" y1="16" x2="23" y2="16" />
    </>
  ),
  clock: (
    <>
      <circle cx="12" cy="12" r="9.5" />
      <polyline points="12 6 12 12 16 14" />
    </>
  ),
  sparkle: (
    <>
      <path d="M11 3l1.7 4.8L17.5 9.5l-4.8 1.7L11 16l-1.7-4.8L4.5 9.5l4.8-1.7z" fill="currentColor" stroke="none" />
      <path d="M18.5 14l.9 2.4 2.4.9-2.4.9-.9 2.4-.9-2.4-2.4-.9 2.4-.9z" fill="currentColor" stroke="none" />
    </>
  ),
  key: (
    <path d="M21 2l-2 2m-7.61 7.61a5.5 5.5 0 1 1-7.778 7.778 5.5 5.5 0 0 1 7.777-7.777zm0 0L15.5 7.5m0 0l3 3L22 7l-3-3m-3.5 3.5L19 4" />
  ),
  cpu: (
    <>
      <rect x="4" y="4" width="16" height="16" rx="2" />
      <rect x="9" y="9" width="6" height="6" />
      <line x1="9" y1="1" x2="9" y2="4" />
      <line x1="15" y1="1" x2="15" y2="4" />
      <line x1="9" y1="20" x2="9" y2="23" />
      <line x1="15" y1="20" x2="15" y2="23" />
      <line x1="20" y1="9" x2="23" y2="9" />
      <line x1="20" y1="14" x2="23" y2="14" />
      <line x1="1" y1="9" x2="4" y2="9" />
      <line x1="1" y1="14" x2="4" y2="14" />
    </>
  ),
  database: (
    <>
      <ellipse cx="12" cy="5" rx="9" ry="3" />
      <path d="M21 12c0 1.66-4 3-9 3s-9-1.34-9-3" />
      <path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5" />
    </>
  ),
  info: (
    <>
      <circle cx="12" cy="12" r="9.5" />
      <line x1="12" y1="16" x2="12" y2="12" />
      <circle cx="12" cy="8" r="0.6" fill="currentColor" stroke="none" />
    </>
  ),
  eye: (
    <>
      <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z" />
      <circle cx="12" cy="12" r="3" />
    </>
  ),

  // --- full office command set ---
  fontSize: (
    <>
      <text x="7" y="17" textAnchor="middle" fontSize="14" fontWeight="700" fill="currentColor" stroke="none" fontFamily="inherit">A</text>
      <text x="16.5" y="17" textAnchor="middle" fontSize="9" fontWeight="600" fill="currentColor" stroke="none" fontFamily="inherit">a</text>
      <line x1="4" y1="20" x2="20" y2="20" opacity="0.5" />
    </>
  ),
  fontName: (
    <>
      <text x="12" y="16.5" textAnchor="middle" fontSize="13" fontWeight="700" fill="currentColor" stroke="none" fontFamily="Georgia, serif">Aa</text>
    </>
  ),
  fill: (
    <>
      <path d="M9.5 3l8 8-7 7a2.1 2.1 0 0 1-3 0l-4.5-4.5a2.1 2.1 0 0 1 0-3z" />
      <path d="M4.5 10.5h11" opacity="0.55" />
      <path d="M20.5 15.5c1 1.4 1.5 2.4 1.5 3.2a1.9 1.9 0 0 1-3.8 0c0-.8.8-1.8 2.3-3.2z" fill="currentColor" stroke="none" />
    </>
  ),
  sigma: (
    <text x="12" y="17" textAnchor="middle" fontSize="15" fontWeight="700" fill="currentColor" stroke="none" fontFamily="Georgia, serif">Σ</text>
  ),
  sortAZ: (
    <>
      <text x="6.5" y="9" textAnchor="middle" fontSize="9" fontWeight="700" fill="currentColor" stroke="none" fontFamily="inherit">A</text>
      <text x="6.5" y="19.5" textAnchor="middle" fontSize="9" fontWeight="700" fill="currentColor" stroke="none" fontFamily="inherit">Z</text>
      <line x1="15" y1="5" x2="15" y2="18" opacity="0.6" />
      <polyline points="11.5 14.5 15 18.5 18.5 14.5" />
    </>
  ),
  sortZA: (
    <>
      <text x="6.5" y="9" textAnchor="middle" fontSize="9" fontWeight="700" fill="currentColor" stroke="none" fontFamily="inherit">Z</text>
      <text x="6.5" y="19.5" textAnchor="middle" fontSize="9" fontWeight="700" fill="currentColor" stroke="none" fontFamily="inherit">A</text>
      <line x1="15" y1="19" x2="15" y2="6" opacity="0.6" />
      <polyline points="11.5 9.5 15 5.5 18.5 9.5" />
    </>
  ),
  rowAbove: (
    <>
      <rect x="3" y="10" width="18" height="6" rx="1" />
      <line x1="3" y1="20" x2="21" y2="20" opacity="0.4" />
      <line x1="12" y1="7" x2="12" y2="2.5" />
      <polyline points="9.5 4.5 12 2 14.5 4.5" />
    </>
  ),
  rowBelow: (
    <>
      <rect x="3" y="8" width="18" height="6" rx="1" />
      <line x1="3" y1="4" x2="21" y2="4" opacity="0.4" />
      <line x1="12" y1="17" x2="12" y2="21.5" />
      <polyline points="9.5 19.5 12 22 14.5 19.5" />
    </>
  ),
  rowDelete: (
    <>
      <rect x="3" y="9" width="18" height="6" rx="1" />
      <line x1="9" y1="12" x2="15" y2="12" />
      <line x1="4" y1="20" x2="8" y2="20" opacity="0.4" />
      <line x1="16" y1="20" x2="20" y2="20" opacity="0.4" />
    </>
  ),
  colLeft: (
    <>
      <rect x="9" y="3" width="6" height="18" rx="1" />
      <line x1="2" y1="12" x2="6.5" y2="12" opacity="0.6" />
      <line x1="20" y1="3" x2="20" y2="21" opacity="0.4" />
      <line x1="3.5" y1="9.5" x2="3.5" y2="14.5" opacity="0.6" />
    </>
  ),
  colRight: (
    <>
      <rect x="9" y="3" width="6" height="18" rx="1" />
      <line x1="17.5" y1="12" x2="22" y2="12" opacity="0.6" />
      <line x1="4" y1="3" x2="4" y2="21" opacity="0.4" />
      <line x1="20.5" y1="9.5" x2="20.5" y2="14.5" opacity="0.6" />
    </>
  ),
  colDelete: (
    <>
      <rect x="9" y="3" width="6" height="18" rx="1" />
      <line x1="12" y1="9" x2="12" y2="15" />
      <line x1="2.5" y1="6" x2="5" y2="6" opacity="0.4" />
      <line x1="2.5" y1="18" x2="5" y2="18" opacity="0.4" />
      <line x1="19" y1="6" x2="21.5" y2="6" opacity="0.4" />
      <line x1="19" y1="18" x2="21.5" y2="18" opacity="0.4" />
    </>
  ),
  paste: (
    <>
      <rect x="6" y="4" width="12" height="17" rx="2" />
      <path d="M9 4a3 3 0 0 1 6 0" />
      <rect x="9" y="1.8" width="6" height="4" rx="1.2" fill="currentColor" stroke="none" opacity="0.85" />
      <line x1="9" y1="11" x2="15" y2="11" opacity="0.55" />
      <line x1="9" y1="15" x2="13" y2="15" opacity="0.55" />
    </>
  ),
  replace: (
    <>
      <path d="M4 8h12" />
      <polyline points="13 4.5 16.5 8 13 11.5" />
      <path d="M20 16H8" />
      <polyline points="11 12.5 7.5 16 11 19.5" />
    </>
  ),
  layoutTitle: (
    <>
      <rect x="3" y="4" width="18" height="16" rx="1.5" />
      <line x1="8" y1="10" x2="16" y2="10" />
      <line x1="10" y1="14" x2="14" y2="14" opacity="0.55" />
    </>
  ),
  layoutContent: (
    <>
      <rect x="3" y="4" width="18" height="16" rx="1.5" />
      <line x1="6" y1="9" x2="18" y2="9" />
      <line x1="6" y1="12.5" x2="18" y2="12.5" opacity="0.55" />
      <line x1="6" y1="16" x2="14" y2="16" opacity="0.55" />
    </>
  ),
  layoutSection: (
    <>
      <rect x="3" y="4" width="18" height="16" rx="1.5" />
      <rect x="3" y="9.5" width="18" height="5" fill="currentColor" stroke="none" opacity="0.75" rx="1" />
    </>
  ),
  layoutBlank: (
    <>
      <rect x="3" y="4" width="18" height="16" rx="1.5" />
    </>
  ),
  theme: (
    <>
      <circle cx="12" cy="12" r="8.5" />
      <path d="M12 3.5a8.5 8.5 0 0 1 0 17z" fill="currentColor" stroke="none" opacity="0.5" />
    </>
  ),
  quote: (
    <>
      <path d="M5 11h4v4a2 2 0 0 1-2 2H5a4 4 0 0 1 0-6z" opacity="0.85" />
      <path d="M14 11h4v4a2 2 0 0 1-2 2h-2a4 4 0 0 1 0-6z" opacity="0.85" />
      <line x1="7" y1="7.5" x2="17" y2="7.5" opacity="0.4" />
    </>
  ),
  minus: <line x1="5" y1="12" x2="19" y2="12" />,
  selectRange: (
    <>
      <rect x="3" y="3" width="18" height="18" rx="2" strokeDasharray="3.5 2.6" />
      <rect x="10" y="10" width="8" height="8" rx="1" fill="currentColor" stroke="none" opacity="0.55" />
    </>
  ),
};

export type IconName = keyof typeof PATHS;

export function Icon({
  name,
  size = 20,
  strokeWidth = 2,
  className,
}: {
  name: string;
  size?: number;
  strokeWidth?: number;
  className?: string;
}) {
  return (
    <svg
      className={className}
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth={strokeWidth}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      {PATHS[name] ?? <circle cx="12" cy="12" r="9" />}
    </svg>
  );
}

/** Microsoft Office-style file type badge: colored rounded square + white letter + fold. */
export function FileTypeIcon({ kind, size = 36 }: { kind: DocKind; size?: number }) {
  const conf: Record<DocKind, { bg: string; fold: string; letter: string }> = {
    doc: { bg: '#185ABD', fold: '#0F3F8C', letter: 'W' },
    sheet: { bg: '#107C41', fold: '#0A532B', letter: 'X' },
    deck: { bg: '#C43E1C', fold: '#8A2B14', letter: 'P' },
    chat: { bg: '#5B5FC7', fold: '#40429A', letter: 'A' },
  };
  const c = conf[kind];
  return (
    <svg width={size} height={size} viewBox="0 0 48 48" aria-hidden="true">
      <rect x="2" y="2" width="44" height="44" rx="7" fill={c.bg} />
      <path d="M32 2h7a7 7 0 0 1 7 7z" fill={c.fold} opacity="0.55" />
      <text
        x="24"
        y="25"
        textAnchor="middle"
        dominantBaseline="central"
        fontSize="24"
        fontWeight="700"
        fontFamily="'Segoe UI', sans-serif"
        fill="#ffffff"
      >
        {c.letter}
      </text>
    </svg>
  );
}
