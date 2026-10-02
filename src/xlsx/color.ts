/**
 * Color conversion between ExcelJS colors and Univer's `#rrggbb` strings.
 *
 * ExcelJS hands back three different color shapes depending on how the source
 * file encoded them, and only `argb` is directly usable:
 *   { argb: 'FFFF0000' }            explicit color
 *   { theme: 4, tint: -0.25 }       theme color + luminance tint
 *   { indexed: 10 }                 legacy 56-color palette (pre-2007 files)
 */

/**
 * Theme color palette for the default Office theme, in the order that
 * SpreadsheetML's `color@theme` index uses. Note this is NOT the order the
 * colors appear in `xl/theme/theme1.xml`: Excel swaps the light/dark pairs, so
 * index 0 is lt1 and index 1 is dk1.
 */
const THEME_COLORS = [
  '#ffffff', // 0 lt1  (background 1)
  '#000000', // 1 dk1  (text 1)
  '#e7e6e6', // 2 lt2  (background 2)
  '#44546a', // 3 dk2  (text 2)
  '#4472c4', // 4 accent1
  '#ed7d31', // 5 accent2
  '#a5a5a5', // 6 accent3
  '#ffc000', // 7 accent4
  '#5b9bd5', // 8 accent5
  '#70ad47', // 9 accent6
  '#0563c1', // 10 hlink
  '#954f72', // 11 folHlink
];

/** The fixed legacy palette referenced by `color@indexed`. */
const INDEXED_COLORS = [
  '#000000', '#ffffff', '#ff0000', '#00ff00', '#0000ff', '#ffff00', '#ff00ff', '#00ffff',
  '#000000', '#ffffff', '#ff0000', '#00ff00', '#0000ff', '#ffff00', '#ff00ff', '#00ffff',
  '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#c0c0c0', '#808080',
  '#9999ff', '#993366', '#ffffcc', '#ccffff', '#660066', '#ff8080', '#0066cc', '#ccccff',
  '#000080', '#ff00ff', '#ffff00', '#00ffff', '#800080', '#800000', '#008080', '#0000ff',
  '#00ccff', '#ccffff', '#ccffcc', '#ffff99', '#99ccff', '#ff99cc', '#cc99ff', '#ffcc99',
  '#3366ff', '#33cccc', '#99cc00', '#ffcc00', '#ff9900', '#ff6600', '#666699', '#969696',
  '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333',
  '#000000', '#ffffff', // 64/65: system foreground / background
];

export interface ExcelColor {
  argb?: string;
  theme?: number;
  tint?: number;
  indexed?: number;
}

function clamp255(n: number): number {
  return Math.max(0, Math.min(255, Math.round(n)));
}

function hex2(n: number): string {
  return clamp255(n).toString(16).padStart(2, '0');
}

function parseHex(hex: string): [number, number, number] {
  const h = hex.replace('#', '');
  return [
    Number.parseInt(h.slice(0, 2), 16),
    Number.parseInt(h.slice(2, 4), 16),
    Number.parseInt(h.slice(4, 6), 16),
  ];
}

function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
  const rn = r / 255;
  const gn = g / 255;
  const bn = b / 255;
  const max = Math.max(rn, gn, bn);
  const min = Math.min(rn, gn, bn);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];

  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === rn) h = ((gn - bn) / d + (gn < bn ? 6 : 0)) / 6;
  else if (max === gn) h = ((bn - rn) / d + 2) / 6;
  else h = ((rn - gn) / d + 4) / 6;
  return [h, s, l];
}

function hueToRgb(p: number, q: number, t: number): number {
  let tt = t;
  if (tt < 0) tt += 1;
  if (tt > 1) tt -= 1;
  if (tt < 1 / 6) return p + (q - p) * 6 * tt;
  if (tt < 1 / 2) return q;
  if (tt < 2 / 3) return p + (q - p) * (2 / 3 - tt) * 6;
  return p;
}

function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  if (s === 0) return [l * 255, l * 255, l * 255];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return [
    hueToRgb(p, q, h + 1 / 3) * 255,
    hueToRgb(p, q, h) * 255,
    hueToRgb(p, q, h - 1 / 3) * 255,
  ];
}

/**
 * Apply SpreadsheetML's `tint` to a base color. The spec defines the shift on
 * HSL luminance, not on the raw RGB channels — doing it in RGB visibly skews
 * saturated accent colors, so convert properly.
 */
function applyTint(hex: string, tint: number): string {
  if (!tint) return hex;
  const [r, g, b] = parseHex(hex);
  const [h, s, l] = rgbToHsl(r, g, b);
  const lum = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  const [nr, ng, nb] = hslToRgb(h, s, Math.max(0, Math.min(1, lum)));
  return `#${hex2(nr)}${hex2(ng)}${hex2(nb)}`;
}

/** ExcelJS color -> `#rrggbb`, or undefined when the color is absent/unusable. */
export function excelColorToHex(color: ExcelColor | undefined | null): string | undefined {
  if (!color) return undefined;

  if (typeof color.argb === 'string' && /^[0-9a-f]{6,8}$/i.test(color.argb)) {
    // argb is normally 8 chars (alpha first); 6-char values show up in the wild.
    const rgb = color.argb.length === 8 ? color.argb.slice(2) : color.argb;
    return `#${rgb.toLowerCase()}`;
  }

  if (typeof color.theme === 'number') {
    const base = THEME_COLORS[color.theme];
    if (base) return applyTint(base, color.tint ?? 0);
  }

  if (typeof color.indexed === 'number') {
    const base = INDEXED_COLORS[color.indexed];
    if (base) return applyTint(base, color.tint ?? 0);
  }

  return undefined;
}

/** `#rrggbb` (or any 3/6-digit hex) -> ExcelJS `{ argb: 'FFrrggbb' }`. */
export function hexToExcelColor(hex: string | undefined | null): { argb: string } | undefined {
  if (!hex) return undefined;
  let h = hex.trim().replace('#', '');
  if (h.length === 3) h = h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
  if (!/^[0-9a-f]{6}$/i.test(h)) return undefined;
  return { argb: `FF${h.toUpperCase()}` };
}
