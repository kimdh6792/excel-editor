/**
 * Flattens a region of a Univer worksheet into the shape a printed table needs.
 *
 * Printing cannot reuse the on-screen grid: Univer draws into a canvas, and a
 * canvas prints as whatever pixels happen to be on screen — one page, cropped
 * at the viewport. So the sheet is re-rendered as an HTML table instead, and
 * this module is the half of that which needs no DOM: it resolves styles,
 * widths, heights and merges from the snapshot so {@link ../print/html} has
 * nothing left to decide.
 *
 * Display text is not derived here. Univer already formats values against each
 * cell's number format, and reimplementing that would be a second, divergent
 * formatter — `FRange.getDisplayValues()` is the source, handed in by the
 * caller.
 */
import { BooleanNumber, CellValueType } from '@univerjs/core';
import type { ICellData, IRange, IStyleData, IWorksheetData, Nullable } from '@univerjs/core';
import { resolveStyle } from '../xlsx/style';
import {
  DEFAULT_COLUMN_WIDTH_CHARS,
  DEFAULT_ROW_HEIGHT_POINTS,
  charsToPixels,
  pointsToPixels,
} from '../xlsx/units';

export interface PrintCell {
  /** Display text as Univer formats it, number format already applied. */
  text: string;
  style?: IStyleData;
  /** True for numeric cells, which Excel right-aligns unless told otherwise. */
  numeric: boolean;
  rowSpan: number;
  colSpan: number;
}

export interface PrintSheet {
  name: string;
  /** One entry per visible row; cells swallowed by a merge are left out. */
  rows: PrintCell[][];
  /** Height in CSS pixels of each visible row, parallel to {@link rows}. */
  rowHeights: number[];
  /** Width in CSS pixels of each visible column. */
  columnWidths: number[];
  showGridlines: boolean;
}

/** The region to print, in absolute sheet coordinates, end indices inclusive. */
export interface Area {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
}

/**
 * A worksheet as it appears inside a snapshot, where every field is optional:
 * `IWorkbookData.sheets` holds `Partial<IWorksheetData>`, so nothing can be
 * assumed present.
 */
export type SheetData = Partial<IWorksheetData>;

export interface BuildPrintSheetInput {
  name: string;
  sheet: SheetData;
  /** The workbook's shared style table, for cells that reference it by id. */
  styles: Record<string, Nullable<IStyleData>>;
  /** Display text, indexed relative to `area` — `values[r - startRow][c - startColumn]`. */
  values: string[][];
  area: Area;
}

/**
 * Narrows an inclusive area to the sheet's own bounds, so a stale selection
 * cannot ask for rows that are not there. A snapshot without a stated row or
 * column count is left alone on that axis — there is nothing to clamp against.
 */
export function clampArea(area: Area, sheet: SheetData): Area {
  const lastRow =
    typeof sheet.rowCount === 'number' ? Math.max(0, sheet.rowCount - 1) : Math.max(0, area.endRow);
  const lastColumn =
    typeof sheet.columnCount === 'number'
      ? Math.max(0, sheet.columnCount - 1)
      : Math.max(0, area.endColumn);
  const startRow = Math.min(Math.max(0, area.startRow), lastRow);
  const startColumn = Math.min(Math.max(0, area.startColumn), lastColumn);
  return {
    startRow,
    startColumn,
    endRow: Math.min(Math.max(startRow, area.endRow), lastRow),
    endColumn: Math.min(Math.max(startColumn, area.endColumn), lastColumn),
  };
}

function isHidden(entry: { hd?: BooleanNumber } | undefined): boolean {
  return entry?.hd === BooleanNumber.TRUE;
}

function columnWidth(sheet: SheetData, column: number): number {
  const stated = sheet.columnData?.[column]?.w;
  if (typeof stated === 'number' && stated > 0) return stated;
  if (typeof sheet.defaultColumnWidth === 'number' && sheet.defaultColumnWidth > 0) {
    return sheet.defaultColumnWidth;
  }
  return charsToPixels(DEFAULT_COLUMN_WIDTH_CHARS);
}

function rowHeight(sheet: SheetData, row: number): number {
  const data = sheet.rowData?.[row];
  // `ah` is the height Univer measured for a self-adaptive row; it only wins
  // when no explicit height was stored, otherwise a user-resized row would
  // print at the measured height instead of the one they chose.
  const stated = data?.h ?? data?.ah;
  if (typeof stated === 'number' && stated > 0) return stated;
  if (typeof sheet.defaultRowHeight === 'number' && sheet.defaultRowHeight > 0) {
    return sheet.defaultRowHeight;
  }
  return pointsToPixels(DEFAULT_ROW_HEIGHT_POINTS);
}

/**
 * Cell style, falling back to the row's and then the column's.
 *
 * A file that formats a whole column carries the style on `columnData` alone,
 * so reading `cellData` only would print those cells unstyled.
 */
function cellStyle(
  cell: ICellData | undefined,
  sheet: SheetData,
  row: number,
  column: number,
  styles: Record<string, Nullable<IStyleData>>,
): IStyleData | undefined {
  return (
    resolveStyle(cell?.s, styles) ??
    resolveStyle(sheet.rowData?.[row]?.s, styles) ??
    resolveStyle(sheet.columnData?.[column]?.s, styles)
  );
}

function isNumeric(cell: ICellData | undefined): boolean {
  if (!cell) return false;
  // A formula result arrives as a plain number with no `t`, so the value's own
  // type has to count as well.
  if (cell.t === CellValueType.NUMBER) return true;
  if (cell.t != null) return false;
  return typeof cell.v === 'number';
}

/** Spans keyed by `row,column` of the cell that carries them. */
type SpanMap = Map<string, { rowSpan: number; colSpan: number }>;

const key = (row: number, column: number) => `${row},${column}`;

/**
 * Works out which cells a merge makes disappear and what the survivor spans.
 *
 * Counted in *visible* rows and columns inside the area, not absolute ones: a
 * merge spanning a hidden row prints one row shorter, and a merge reaching past
 * the printed area is clipped to it, with its top-left corner moving to the
 * first visible cell that remains.
 */
function planMerges(
  merges: IRange[] | undefined,
  area: Area,
  visibleRows: number[],
  visibleColumns: number[],
): { spans: SpanMap; covered: Set<string> } {
  const spans: SpanMap = new Map();
  const covered = new Set<string>();
  if (!merges) return { spans, covered };

  for (const merge of merges) {
    const top = Math.max(merge.startRow, area.startRow);
    const left = Math.max(merge.startColumn, area.startColumn);
    const bottom = Math.min(merge.endRow, area.endRow);
    const right = Math.min(merge.endColumn, area.endColumn);
    if (top > bottom || left > right) continue;

    const rows = visibleRows.filter((r) => r >= top && r <= bottom);
    const columns = visibleColumns.filter((c) => c >= left && c <= right);
    if (rows.length === 0 || columns.length === 0) continue;

    const anchorRow = rows[0];
    const anchorColumn = columns[0];
    spans.set(key(anchorRow, anchorColumn), {
      rowSpan: rows.length,
      colSpan: columns.length,
    });

    for (const r of rows) {
      for (const c of columns) {
        if (r === anchorRow && c === anchorColumn) continue;
        covered.add(key(r, c));
      }
    }
  }

  return { spans, covered };
}

export function buildPrintSheet({
  name,
  sheet,
  styles,
  values,
  area,
}: BuildPrintSheetInput): PrintSheet {
  const bounds = clampArea(area, sheet);

  const visibleRows: number[] = [];
  for (let r = bounds.startRow; r <= bounds.endRow; r += 1) {
    if (!isHidden(sheet.rowData?.[r])) visibleRows.push(r);
  }

  const visibleColumns: number[] = [];
  for (let c = bounds.startColumn; c <= bounds.endColumn; c += 1) {
    if (!isHidden(sheet.columnData?.[c])) visibleColumns.push(c);
  }

  const { spans, covered } = planMerges(sheet.mergeData, bounds, visibleRows, visibleColumns);

  const rows: PrintCell[][] = [];
  for (const r of visibleRows) {
    const out: PrintCell[] = [];
    for (const c of visibleColumns) {
      if (covered.has(key(r, c))) continue;

      const cell = sheet.cellData?.[r]?.[c] as ICellData | undefined;
      const span = spans.get(key(r, c));
      out.push({
        text: values[r - bounds.startRow]?.[c - bounds.startColumn] ?? '',
        style: cellStyle(cell, sheet, r, c, styles),
        numeric: isNumeric(cell),
        rowSpan: span?.rowSpan ?? 1,
        colSpan: span?.colSpan ?? 1,
      });
    }
    rows.push(out);
  }

  return {
    name,
    rows,
    rowHeights: visibleRows.map((r) => rowHeight(sheet, r)),
    columnWidths: visibleColumns.map((c) => columnWidth(sheet, c)),
    showGridlines: sheet.showGridlines !== 0,
  };
}

/** True when nothing in the area would put ink on the page. */
export function isBlank(sheet: PrintSheet): boolean {
  return sheet.rows.every((row) => row.every((cell) => cell.text.trim() === ''));
}
