/**
 * Renders print sheets as an HTML table.
 *
 * Returns markup only — the surrounding rules live in `App.css` under
 * `@media print`, so the screen and print stylesheets stay in one place. What
 * has to be inline is the per-cell formatting, which is data rather than design
 * and would otherwise need a generated stylesheet per document.
 *
 * Everything that reaches an attribute here came out of a spreadsheet file, so
 * colours and font names are validated rather than trusted: a `cl.rgb` of
 * `red" onload="…` must not be able to close the attribute it sits in.
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
import type { IBorderStyleData, Nullable } from '@univerjs/core';
import type { PrintCell, PrintSheet } from './area';
import { MAX_REPEAT_HEADER_ROWS } from './options';
import type { PrintOptions } from './options';

/* ------------------------------------------------------------------ escaping */

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Accepts only a literal hex colour. Univer's own converters produce `#rrggbb`
 * and anything else in the field is either a theme reference this app does not
 * resolve or a hostile file, both of which are better dropped than passed on.
 */
function safeColor(value: Nullable<string> | undefined): string | undefined {
  if (!value) return undefined;
  return /^#(?:[0-9a-f]{3}|[0-9a-f]{6}|[0-9a-f]{8})$/i.test(value) ? value : undefined;
}

/**
 * Keeps a font name to characters that cannot end the declaration early: no
 * quotes, semicolons, colons or brackets survive. The allowed range covers
 * non-Latin names, which are all well above the punctuation being removed.
 */
function safeFontFamily(value: string | undefined): string | undefined {
  if (!value) return undefined;
  const cleaned = value
    .replace(/[^\w\sÀ-￿-]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
  return cleaned ? `'${cleaned}'` : undefined;
}

/* ------------------------------------------------------------------- borders */

/**
 * Excel's thirteen line styles collapse onto what CSS can draw: a weight and
 * one of solid, dashed, dotted or double. The dash-dot family has no CSS
 * equivalent, so it prints dashed at the matching weight.
 */
const BORDER_CSS: Record<number, string> = {
  [BorderStyleTypes.THIN]: '1px solid',
  [BorderStyleTypes.HAIR]: '1px solid',
  [BorderStyleTypes.DOTTED]: '1px dotted',
  [BorderStyleTypes.DASHED]: '1px dashed',
  [BorderStyleTypes.DASH_DOT]: '1px dashed',
  [BorderStyleTypes.DASH_DOT_DOT]: '1px dashed',
  [BorderStyleTypes.DOUBLE]: '3px double',
  [BorderStyleTypes.MEDIUM]: '2px solid',
  [BorderStyleTypes.MEDIUM_DASHED]: '2px dashed',
  [BorderStyleTypes.MEDIUM_DASH_DOT]: '2px dashed',
  [BorderStyleTypes.MEDIUM_DASH_DOT_DOT]: '2px dashed',
  [BorderStyleTypes.SLANT_DASH_DOT]: '2px dashed',
  [BorderStyleTypes.THICK]: '3px solid',
};

function borderValue(border: Nullable<IBorderStyleData>): string | undefined {
  if (!border || border.s === BorderStyleTypes.NONE) return undefined;
  const css = BORDER_CSS[border.s];
  if (!css) return undefined;
  return `${css} ${safeColor(border.cl?.rgb) ?? '#000000'}`;
}

/* ----------------------------------------------------------------- alignment */

const H_ALIGN_CSS: Record<number, string> = {
  [HorizontalAlign.LEFT]: 'left',
  [HorizontalAlign.CENTER]: 'center',
  [HorizontalAlign.RIGHT]: 'right',
  [HorizontalAlign.JUSTIFIED]: 'justify',
  [HorizontalAlign.BOTH]: 'justify',
  [HorizontalAlign.DISTRIBUTED]: 'justify',
};

const V_ALIGN_CSS: Record<number, string> = {
  [VerticalAlign.TOP]: 'top',
  [VerticalAlign.MIDDLE]: 'middle',
  [VerticalAlign.BOTTOM]: 'bottom',
};

/* ---------------------------------------------------------------- cell style */

function cellCss(cell: PrintCell): string {
  const style = cell.style;
  const declarations: string[] = [];

  const family = safeFontFamily(style?.ff ?? undefined);
  if (family) declarations.push(`font-family:${family}`);
  if (typeof style?.fs === 'number' && style.fs > 0) declarations.push(`font-size:${style.fs}pt`);
  if (style?.bl === BooleanNumber.TRUE) declarations.push('font-weight:700');
  if (style?.it === BooleanNumber.TRUE) declarations.push('font-style:italic');

  const decorations: string[] = [];
  if (style?.ul?.s === BooleanNumber.TRUE) decorations.push('underline');
  if (style?.st?.s === BooleanNumber.TRUE) decorations.push('line-through');
  if (decorations.length > 0) {
    const doubled = style?.ul?.t === TextDecoration.DOUBLE ? ' double' : '';
    declarations.push(`text-decoration:${decorations.join(' ')}${doubled}`);
  }

  const color = safeColor(style?.cl?.rgb);
  if (color) declarations.push(`color:${color}`);
  const background = safeColor(style?.bg?.rgb);
  if (background) declarations.push(`background-color:${background}`);

  // Excel right-aligns numbers and left-aligns text when no alignment is set,
  // and a printed column of figures reads wrong without it.
  const horizontal = style?.ht != null ? H_ALIGN_CSS[style.ht] : undefined;
  if (horizontal) declarations.push(`text-align:${horizontal}`);
  else if (cell.numeric) declarations.push('text-align:right');

  const vertical = style?.vt != null ? V_ALIGN_CSS[style.vt] : undefined;
  if (vertical) declarations.push(`vertical-align:${vertical}`);

  // The stylesheet's default is `pre`, which keeps line breaks without wrapping;
  // a wrapped cell needs the wrapping variant of the same thing.
  if (style?.tb === WrapStrategy.WRAP) declarations.push('white-space:pre-wrap');

  if (style?.bd) {
    const top = borderValue(style.bd.t);
    const right = borderValue(style.bd.r);
    const bottom = borderValue(style.bd.b);
    const left = borderValue(style.bd.l);
    if (top) declarations.push(`border-top:${top}`);
    if (right) declarations.push(`border-right:${right}`);
    if (bottom) declarations.push(`border-bottom:${bottom}`);
    if (left) declarations.push(`border-left:${left}`);
  }

  if (typeof style?.pd?.l === 'number' && style.pd.l > 0) {
    declarations.push(`padding-left:${Math.round(style.pd.l)}px`);
  }

  return declarations.join(';');
}

/** Super- and subscript are a wrapping element, not a declaration. */
function cellText(cell: PrintCell): string {
  const text = escapeHtml(cell.text);
  if (!text) return '';
  if (cell.style?.va === BaselineOffset.SUPERSCRIPT) return `<sup>${text}</sup>`;
  if (cell.style?.va === BaselineOffset.SUBSCRIPT) return `<sub>${text}</sub>`;
  return text;
}

function renderCell(cell: PrintCell): string {
  const attributes: string[] = [];
  if (cell.colSpan > 1) attributes.push(` colspan="${cell.colSpan}"`);
  if (cell.rowSpan > 1) attributes.push(` rowspan="${cell.rowSpan}"`);
  const css = cellCss(cell);
  if (css) attributes.push(` style="${css}"`);
  return `<td${attributes.join('')}>${cellText(cell)}</td>`;
}

function renderRow(cells: PrintCell[], height: number): string {
  // `height` is a minimum for a table row, so a wrapped cell can still grow.
  return `<tr style="height:${Math.round(height)}px">${cells.map(renderCell).join('')}</tr>`;
}

/**
 * Column widths.
 *
 * "Fit to width" is not a scale factor: the table is told to fill the page and
 * each column is given its share as a percentage, which keeps the proportions
 * between columns while letting the printer decide the absolute size. At actual
 * size the pixel widths go through unchanged and a sheet wider than the paper
 * is clipped, which is the honest outcome — there is no horizontal pagination.
 */
function renderColgroup(widths: number[], fitToWidth: boolean): string {
  const total = widths.reduce((sum, width) => sum + width, 0);
  const columns = widths.map((width) => {
    if (!fitToWidth) return `<col style="width:${Math.round(width)}px">`;
    const share = total > 0 ? (width / total) * 100 : 100 / Math.max(1, widths.length);
    return `<col style="width:${share.toFixed(4)}%">`;
  });
  return `<colgroup>${columns.join('')}</colgroup>`;
}

function renderSheet(sheet: PrintSheet, options: PrintOptions, showName: boolean): string {
  const repeat = Math.min(
    Math.max(0, Math.trunc(options.repeatHeaderRows)),
    MAX_REPEAT_HEADER_ROWS,
    // A repeat covering every row would leave nothing to paginate.
    Math.max(0, sheet.rows.length - 1),
  );

  const header = sheet.rows
    .slice(0, repeat)
    .map((cells, index) => renderRow(cells, sheet.rowHeights[index]))
    .join('');
  const body = sheet.rows
    .slice(repeat)
    .map((cells, index) => renderRow(cells, sheet.rowHeights[repeat + index]))
    .join('');

  const gridlines = options.gridlines && sheet.showGridlines ? ' gridlines' : '';
  const width = options.fitToWidth ? 'width:100%' : 'width:auto';
  const caption = showName ? `<h2>${escapeHtml(sheet.name)}</h2>` : '';

  return (
    `<section class="print-sheet">${caption}` +
    `<table class="print-table${gridlines}" style="${width}">` +
    renderColgroup(sheet.columnWidths, options.fitToWidth) +
    (header ? `<thead>${header}</thead>` : '') +
    `<tbody>${body}</tbody>` +
    '</table></section>'
  );
}

/**
 * Builds the markup for the print container. Sheet names are only titled when
 * more than one prints, so the common single-sheet case gets no stray heading.
 */
export function buildPrintHtml(sheets: PrintSheet[], options: PrintOptions): string {
  const showNames = sheets.length > 1;
  return sheets.map((sheet) => renderSheet(sheet, options, showNames)).join('');
}
