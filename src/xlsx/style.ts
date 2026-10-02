/**
 * Bidirectional mapping between an ExcelJS `Style` and Univer's `IStyleData`.
 *
 * Both sides are kept in one file on purpose: every field added to one
 * direction needs its mirror, and splitting them is how round-trip bugs creep
 * in. Anything not represented here is dropped on save — see `inspect.ts` for
 * the warning surfaced to the user.
 */
import {
  BaselineOffset,
  BooleanNumber,
  BorderStyleTypes,
  HorizontalAlign,
  TextDecoration,
  VerticalAlign,
  WrapStrategy,
} from '@univerjs/core';
import type { IBorderStyleData, IStyleData, Nullable } from '@univerjs/core';
import type { Alignment, Border, Borders, Fill, Font, Style } from 'exceljs';
import { excelColorToHex, hexToExcelColor } from './color';
import type { ExcelColor } from './color';

type ExcelStyle = Partial<Style>;

/* ------------------------------------------------------------------ borders */

const BORDER_STYLE_TO_UNIVER: Record<string, BorderStyleTypes> = {
  thin: BorderStyleTypes.THIN,
  hair: BorderStyleTypes.HAIR,
  dotted: BorderStyleTypes.DOTTED,
  dashed: BorderStyleTypes.DASHED,
  dashDot: BorderStyleTypes.DASH_DOT,
  dashDotDot: BorderStyleTypes.DASH_DOT_DOT,
  double: BorderStyleTypes.DOUBLE,
  medium: BorderStyleTypes.MEDIUM,
  mediumDashed: BorderStyleTypes.MEDIUM_DASHED,
  mediumDashDot: BorderStyleTypes.MEDIUM_DASH_DOT,
  mediumDashDotDot: BorderStyleTypes.MEDIUM_DASH_DOT_DOT,
  slantDashDot: BorderStyleTypes.SLANT_DASH_DOT,
  thick: BorderStyleTypes.THICK,
};

const BORDER_STYLE_TO_EXCEL = new Map<BorderStyleTypes, Border['style']>(
  Object.entries(BORDER_STYLE_TO_UNIVER).map(
    ([name, value]) => [value, name as Border['style']],
  ),
);

/** Univer border slot -> ExcelJS `Borders` key. Diagonals collapse onto one slot. */
const BORDER_SLOTS: Array<[keyof Borders, 't' | 'r' | 'b' | 'l']> = [
  ['top', 't'],
  ['right', 'r'],
  ['bottom', 'b'],
  ['left', 'l'],
];

function toUniverBorder(border: Partial<Border> | undefined): IBorderStyleData | undefined {
  if (!border?.style) return undefined;
  const s = BORDER_STYLE_TO_UNIVER[border.style];
  if (s === undefined) return undefined;
  return { s, cl: { rgb: excelColorToHex(border.color as ExcelColor) ?? '#000000' } };
}

function toExcelBorder(border: Nullable<IBorderStyleData>): Partial<Border> | undefined {
  if (!border || border.s === BorderStyleTypes.NONE) return undefined;
  const style = BORDER_STYLE_TO_EXCEL.get(border.s);
  if (!style) return undefined;
  const color = hexToExcelColor(border.cl?.rgb ?? undefined);
  return color ? { style, color } : { style };
}

/* -------------------------------------------------------------- alignment */

const H_ALIGN_TO_UNIVER: Record<string, HorizontalAlign> = {
  left: HorizontalAlign.LEFT,
  center: HorizontalAlign.CENTER,
  right: HorizontalAlign.RIGHT,
  justify: HorizontalAlign.JUSTIFIED,
  distributed: HorizontalAlign.DISTRIBUTED,
  // `fill` and `centerContinuous` have no Univer equivalent; approximate.
  fill: HorizontalAlign.LEFT,
  centerContinuous: HorizontalAlign.CENTER,
};

const H_ALIGN_TO_EXCEL: Record<number, Alignment['horizontal']> = {
  [HorizontalAlign.LEFT]: 'left',
  [HorizontalAlign.CENTER]: 'center',
  [HorizontalAlign.RIGHT]: 'right',
  [HorizontalAlign.JUSTIFIED]: 'justify',
  [HorizontalAlign.BOTH]: 'justify',
  [HorizontalAlign.DISTRIBUTED]: 'distributed',
};

const V_ALIGN_TO_UNIVER: Record<string, VerticalAlign> = {
  top: VerticalAlign.TOP,
  middle: VerticalAlign.MIDDLE,
  bottom: VerticalAlign.BOTTOM,
  justify: VerticalAlign.MIDDLE,
  distributed: VerticalAlign.MIDDLE,
};

const V_ALIGN_TO_EXCEL: Record<number, Alignment['vertical']> = {
  [VerticalAlign.TOP]: 'top',
  [VerticalAlign.MIDDLE]: 'middle',
  [VerticalAlign.BOTTOM]: 'bottom',
};

/* ------------------------------------------------------------------- fill */

/** Solid pattern fills map cleanly; gradients degrade to their first stop. */
function fillToHex(fill: Fill | undefined): string | undefined {
  if (!fill) return undefined;
  if (fill.type === 'pattern') {
    if (fill.pattern === 'none') return undefined;
    // For non-solid patterns the foreground is the closest single-color stand-in.
    return excelColorToHex((fill.fgColor ?? fill.bgColor) as ExcelColor);
  }
  if (fill.type === 'gradient') {
    const stop = fill.stops?.[0];
    return excelColorToHex(stop?.color as ExcelColor);
  }
  return undefined;
}

/* ------------------------------------------------- ExcelJS -> Univer style */

export function toUniverStyle(style: ExcelStyle | undefined): IStyleData | undefined {
  if (!style) return undefined;

  const out: IStyleData = {};
  const font = style.font as Partial<Font> | undefined;

  if (font) {
    if (font.name) out.ff = font.name;
    if (typeof font.size === 'number') out.fs = font.size;
    if (font.bold !== undefined) out.bl = font.bold ? BooleanNumber.TRUE : BooleanNumber.FALSE;
    if (font.italic !== undefined) out.it = font.italic ? BooleanNumber.TRUE : BooleanNumber.FALSE;

    if (font.underline) {
      // ExcelJS reports `true` or one of single/double/singleAccounting/doubleAccounting.
      const doubled = font.underline === 'double' || font.underline === 'doubleAccounting';
      out.ul = {
        s: BooleanNumber.TRUE,
        t: doubled ? TextDecoration.DOUBLE : TextDecoration.SINGLE,
      };
    }
    if (font.strike) out.st = { s: BooleanNumber.TRUE };

    const fontColor = excelColorToHex(font.color as ExcelColor);
    if (fontColor) out.cl = { rgb: fontColor };

    if (font.vertAlign === 'superscript') out.va = BaselineOffset.SUPERSCRIPT;
    else if (font.vertAlign === 'subscript') out.va = BaselineOffset.SUBSCRIPT;
  }

  const bg = fillToHex(style.fill);
  if (bg) out.bg = { rgb: bg };

  if (style.numFmt) out.n = { pattern: style.numFmt };

  const align = style.alignment as Partial<Alignment> | undefined;
  if (align) {
    if (align.horizontal && H_ALIGN_TO_UNIVER[align.horizontal] !== undefined) {
      out.ht = H_ALIGN_TO_UNIVER[align.horizontal];
    }
    if (align.vertical && V_ALIGN_TO_UNIVER[align.vertical] !== undefined) {
      out.vt = V_ALIGN_TO_UNIVER[align.vertical];
    }
    // Univer has no separate "no wrap" default, so only set when meaningful.
    if (align.wrapText) out.tb = WrapStrategy.WRAP;
    if (align.shrinkToFit) out.stf = BooleanNumber.TRUE;

    if (align.textRotation === 'vertical') {
      out.tr = { a: 0, v: BooleanNumber.TRUE };
    } else if (typeof align.textRotation === 'number' && align.textRotation !== 0) {
      out.tr = { a: align.textRotation, v: BooleanNumber.FALSE };
    }
    if (align.indent) out.pd = { l: align.indent * 8, t: 0, b: 0, r: 0 };
  }

  const border = style.border as Partial<Borders> | undefined;
  if (border) {
    const bd: NonNullable<IStyleData['bd']> = {};
    for (const [excelKey, univerKey] of BORDER_SLOTS) {
      const mapped = toUniverBorder(border[excelKey] as Partial<Border> | undefined);
      if (mapped) bd[univerKey] = mapped;
    }
    const diagonal = border.diagonal;
    if (diagonal?.style) {
      const mapped = toUniverBorder(diagonal);
      if (mapped) {
        if (diagonal.down) bd.tl_br = mapped;
        if (diagonal.up) bd.bl_tr = mapped;
      }
    }
    if (Object.keys(bd).length > 0) out.bd = bd;
  }

  return Object.keys(out).length > 0 ? out : undefined;
}

/* ------------------------------------------------- Univer -> ExcelJS style */

export function toExcelStyle(style: IStyleData | undefined | null): ExcelStyle | undefined {
  if (!style) return undefined;

  const out: ExcelStyle = {};
  const font: Partial<Font> = {};

  if (style.ff) font.name = style.ff;
  if (typeof style.fs === 'number') font.size = style.fs;
  if (style.bl !== undefined) font.bold = style.bl === BooleanNumber.TRUE;
  if (style.it !== undefined) font.italic = style.it === BooleanNumber.TRUE;

  if (style.ul?.s === BooleanNumber.TRUE) {
    font.underline = style.ul.t === TextDecoration.DOUBLE ? 'double' : 'single';
  }
  if (style.st?.s === BooleanNumber.TRUE) font.strike = true;

  const fontColor = hexToExcelColor(style.cl?.rgb ?? undefined);
  if (fontColor) font.color = fontColor;

  if (style.va === BaselineOffset.SUPERSCRIPT) font.vertAlign = 'superscript';
  else if (style.va === BaselineOffset.SUBSCRIPT) font.vertAlign = 'subscript';

  if (Object.keys(font).length > 0) out.font = font;

  const bgColor = hexToExcelColor(style.bg?.rgb ?? undefined);
  if (bgColor) {
    out.fill = { type: 'pattern', pattern: 'solid', fgColor: bgColor, bgColor };
  }

  if (style.n?.pattern) out.numFmt = style.n.pattern;

  const alignment: Partial<Alignment> = {};
  if (style.ht != null && H_ALIGN_TO_EXCEL[style.ht]) alignment.horizontal = H_ALIGN_TO_EXCEL[style.ht];
  if (style.vt != null && V_ALIGN_TO_EXCEL[style.vt]) alignment.vertical = V_ALIGN_TO_EXCEL[style.vt];
  if (style.tb === WrapStrategy.WRAP) alignment.wrapText = true;
  if (style.stf === BooleanNumber.TRUE) alignment.shrinkToFit = true;
  if (style.tr) {
    alignment.textRotation = style.tr.v === BooleanNumber.TRUE ? 'vertical' : style.tr.a;
  }
  if (style.pd?.l) alignment.indent = Math.round(style.pd.l / 8);
  if (Object.keys(alignment).length > 0) out.alignment = alignment;

  if (style.bd) {
    const border: Partial<Borders> = {};
    for (const [excelKey, univerKey] of BORDER_SLOTS) {
      const mapped = toExcelBorder(style.bd[univerKey]);
      if (mapped) border[excelKey] = mapped as never;
    }
    const down = toExcelBorder(style.bd.tl_br);
    const up = toExcelBorder(style.bd.bl_tr);
    if (down || up) {
      border.diagonal = { ...(down ?? up), up: Boolean(up), down: Boolean(down) };
    }
    if (Object.keys(border).length > 0) out.border = border;
  }

  return Object.keys(out).length > 0 ? out : undefined;
}

/* ------------------------------------------------------------ style table */

/**
 * Serialises a value with object keys in a stable order, so two styles that
 * differ only in key order hash the same.
 *
 * Note this cannot be `JSON.stringify(value, sortedKeys)`: an array replacer is
 * a property *allowlist* applied recursively, which would reduce every nested
 * object to `{}` and make unrelated styles compare equal.
 */
function stableKey(value: unknown): string {
  if (value === null || typeof value !== 'object') return JSON.stringify(value) ?? 'null';
  if (Array.isArray(value)) return `[${value.map(stableKey).join(',')}]`;

  const entries = Object.entries(value as Record<string, unknown>)
    .filter(([, v]) => v !== undefined)
    .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0));
  return `{${entries.map(([k, v]) => `${JSON.stringify(k)}:${stableKey(v)}`).join(',')}}`;
}

/**
 * Collects styles into Univer's shared `styles` table, deduplicating by value
 * so that a sheet with one uniform header row stores that style once rather
 * than per cell. Inline styles would also work but bloat the snapshot enough to
 * slow down load on real-world files.
 */
export class StyleTable {
  private readonly byKey = new Map<string, string>();
  readonly styles: Record<string, IStyleData> = {};

  /** Returns a style id, or undefined when the style carries no information. */
  add(style: IStyleData | undefined): string | undefined {
    if (!style || Object.keys(style).length === 0) return undefined;
    const key = stableKey(style);
    const existing = this.byKey.get(key);
    if (existing) return existing;
    const id = `s${this.byKey.size + 1}`;
    this.byKey.set(key, id);
    this.styles[id] = style;
    return id;
  }
}

/** Resolves a Univer style reference (id or inline object) to style data. */
export function resolveStyle(
  ref: Nullable<IStyleData | string>,
  styles: Record<string, Nullable<IStyleData>>,
): IStyleData | undefined {
  if (!ref) return undefined;
  if (typeof ref === 'string') return styles[ref] ?? undefined;
  return ref;
}
