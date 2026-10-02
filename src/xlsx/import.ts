/**
 * xlsx bytes -> Univer workbook snapshot.
 *
 * ExcelJS does the XML parsing; this module is purely the model translation.
 * Anything ExcelJS surfaces but Univer cannot represent natively (rich text
 * runs, hyperlinks) is stashed on the cell's `custom` field so `export.ts` can
 * put it back when the cell was not edited.
 */
import ExcelJS from 'exceljs';
import { CellValueType, LocaleType } from '@univerjs/core';
import type {
  ICellData,
  IRange,
  IStyleData,
  IWorkbookData,
  IWorksheetData,
} from '@univerjs/core';
import { buildFilterResource } from './filter';
import { normalizeXlsx } from './normalize';
import { parseRangeRef } from './range';
import { MIN_COLS, MIN_ROWS } from './snapshot';
import { StyleTable, toUniverStyle } from './style';
import {
  DEFAULT_COLUMN_WIDTH_CHARS,
  DEFAULT_ROW_HEIGHT_POINTS,
  charsToPixels,
  dateToSerial,
  pointsToPixels,
} from './units';


/** Data we carry through the round trip because Univer has no field for it. */
export interface PreservedCell {
  /** Original ExcelJS rich-text runs, plus the flattened text they produced. */
  richText?: { runs: unknown; text: string };
  /** Original hyperlink target, plus the display text it produced. */
  hyperlink?: { target: string; text: string };
}

/** Original row height in points, before conversion to pixels. */
export interface PreservedRow {
  points: number;
}

/** Original column width in character units, before conversion to pixels. */
export interface PreservedColumn {
  chars: number;
}

export interface ImportResult {
  snapshot: IWorkbookData;
  /** Sheet names in file order, for UI display. */
  sheetNames: string[];
}

/** Builds the Univer cell for one ExcelJS cell, or undefined when it holds nothing. */
function convertCell(cell: ExcelJS.Cell, styleId: string | undefined): ICellData | undefined {
  const out: ICellData = {};
  if (styleId) out.s = styleId;

  const preserved: PreservedCell = {};
  const {
    Formula,
    Date: DateType,
    Hyperlink,
    RichText,
    Error: ErrorType,
    Boolean: BoolType,
    Number: NumType,
    Merge: MergeType,
  } = ExcelJS.ValueType;

  switch (cell.type) {
    case MergeType: {
      // ExcelJS proxies a merged slave's `value` to its master. Univer expects
      // only the top-left cell of a merge to hold the value, so take the style
      // and nothing else — otherwise the text is duplicated across the range.
      break;
    }

    case Formula: {
      // `cell.formula` translates shared formulas into this cell's own A1 refs.
      if (cell.formula) out.f = `=${cell.formula}`;
      const result = cell.result as unknown;
      if (result instanceof Date) {
        out.v = dateToSerial(result);
        out.t = CellValueType.NUMBER;
      } else if (typeof result === 'number') {
        out.v = result;
        out.t = CellValueType.NUMBER;
      } else if (typeof result === 'boolean') {
        out.v = result;
        out.t = CellValueType.BOOLEAN;
      } else if (result && typeof result === 'object' && 'error' in result) {
        out.v = String((result as { error: string }).error);
        out.t = CellValueType.STRING;
      } else if (typeof result === 'string') {
        out.v = result;
        out.t = CellValueType.STRING;
      }
      break;
    }

    case DateType: {
      // Univer stores dates the way the file does: a serial plus a number format.
      out.v = dateToSerial(cell.value as Date);
      out.t = CellValueType.NUMBER;
      break;
    }

    case Hyperlink: {
      const value = cell.value as { hyperlink: string; text: string };
      out.v = value.text ?? value.hyperlink ?? '';
      out.t = CellValueType.STRING;
      preserved.hyperlink = { target: value.hyperlink, text: String(out.v) };
      break;
    }

    case RichText: {
      // Univer models rich text as a nested document; flattening to plain text
      // keeps the grid usable, and the original runs are restored on save as
      // long as the text was not edited.
      const value = cell.value as { richText: Array<{ text: string }> };
      const text = value.richText.map((run) => run.text).join('');
      out.v = text;
      out.t = CellValueType.STRING;
      preserved.richText = { runs: value.richText, text };
      break;
    }

    case ErrorType: {
      out.v = String((cell.value as { error: string }).error);
      out.t = CellValueType.STRING;
      break;
    }

    case BoolType: {
      out.v = cell.value as boolean;
      out.t = CellValueType.BOOLEAN;
      break;
    }

    case NumType: {
      out.v = cell.value as number;
      out.t = CellValueType.NUMBER;
      break;
    }

    default: {
      // String, SharedString, Null, Merge.
      if (cell.value !== null && cell.value !== undefined) {
        out.v = String(cell.value);
        out.t = CellValueType.STRING;
      }
    }
  }

  if (Object.keys(preserved).length > 0) out.custom = preserved;

  // Styles alone are worth keeping (empty but bordered/filled cells).
  return out.s !== undefined || out.v !== undefined || out.f !== undefined ? out : undefined;
}

function convertSheet(
  ws: ExcelJS.Worksheet,
  sheetId: string,
  styleTable: StyleTable,
): Partial<IWorksheetData> {
  const cellData: Record<number, Record<number, ICellData>> = {};
  let maxRow = 0;
  let maxCol = 0;

  ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const styleId = styleTable.add(toUniverStyle(cell.style));
      const converted = convertCell(cell, styleId);
      if (!converted) return;

      const r = rowNumber - 1;
      const c = colNumber - 1;
      (cellData[r] ??= {})[c] = converted;
      if (r > maxRow) maxRow = r;
      if (c > maxCol) maxCol = c;
    });
  });

  /* ------------------------------------------------------------ rows/cols */

  // Univer sizes everything in pixels while xlsx uses points and character
  // widths, and neither conversion is exact. The source value rides along in
  // `custom` so a row or column the user never resized is written back byte-for
  // -byte identical instead of drifting a fraction every save.
  const rowData: Record<number, { h?: number; hd?: 0 | 1; custom?: PreservedRow }> = {};
  for (let r = 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const entry: { h?: number; hd?: 0 | 1; custom?: PreservedRow } = {};
    if (typeof row.height === 'number') {
      entry.h = pointsToPixels(row.height);
      entry.custom = { points: row.height };
    }
    if (row.hidden) entry.hd = 1;
    if (Object.keys(entry).length > 0) rowData[r - 1] = entry;
  }

  const columnData: Record<
    number,
    { w?: number; hd?: 0 | 1; s?: string; custom?: PreservedColumn }
  > = {};
  for (let c = 1; c <= ws.columnCount; c++) {
    const col = ws.getColumn(c);
    const entry: { w?: number; hd?: 0 | 1; s?: string; custom?: PreservedColumn } = {};
    if (typeof col.width === 'number') {
      entry.w = charsToPixels(col.width);
      entry.custom = { chars: col.width };
    }
    if (col.hidden) entry.hd = 1;
    // A column-level style paints cells that hold no data of their own.
    const colStyleId = styleTable.add(toUniverStyle(col.style));
    if (colStyleId) entry.s = colStyleId;
    if (Object.keys(entry).length > 0) columnData[c - 1] = entry;
  }

  /* ----------------------------------------------------------- merges etc */

  const merges = (ws.model?.merges ?? []) as string[];
  const mergeData = merges
    .map(parseRangeRef)
    .filter((r): r is IRange => Boolean(r));

  const view = ws.views?.[0];
  const freeze =
    view && view.state === 'frozen'
      ? {
          xSplit: view.xSplit ?? 0,
          ySplit: view.ySplit ?? 0,
          startRow: view.ySplit ?? 0,
          startColumn: view.xSplit ?? 0,
        }
      : { xSplit: 0, ySplit: 0, startRow: -1, startColumn: -1 };

  const props = ws.properties;
  const defaultRowHeight = pointsToPixels(props?.defaultRowHeight || DEFAULT_ROW_HEIGHT_POINTS);
  const defaultColumnWidth = charsToPixels(props?.defaultColWidth || DEFAULT_COLUMN_WIDTH_CHARS);

  return {
    id: sheetId,
    name: ws.name,
    tabColor: (props?.tabColor as { argb?: string } | undefined)?.argb
      ? `#${(props!.tabColor as { argb: string }).argb.slice(-6).toLowerCase()}`
      : '',
    hidden: ws.state === 'visible' ? 0 : ws.state === 'veryHidden' ? 2 : 1,
    rowCount: Math.max(maxRow + 1, ws.rowCount, MIN_ROWS),
    columnCount: Math.max(maxCol + 1, ws.columnCount, MIN_COLS),
    zoomRatio: 1,
    freeze,
    scrollTop: 0,
    scrollLeft: 0,
    defaultColumnWidth,
    defaultRowHeight,
    mergeData,
    cellData,
    rowData,
    columnData,
    rowHeader: { width: 46 },
    columnHeader: { height: 20 },
    showGridlines: view?.showGridLines === false ? 0 : 1,
    rightToLeft: view?.rightToLeft ? 1 : 0,
  };
}

/** Reads xlsx bytes into a Univer snapshot. */
export async function importXlsx(bytes: ArrayBuffer | Uint8Array, name: string): Promise<ImportResult> {
  const raw = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  // Some writers omit the optional `r` attributes that ExcelJS insists on.
  const normalized = normalizeXlsx(raw);

  const workbook = new ExcelJS.Workbook();
  // ExcelJS wants a plain ArrayBuffer; a Uint8Array view would include its offset.
  await workbook.xlsx.load(
    normalized.buffer.slice(
      normalized.byteOffset,
      normalized.byteOffset + normalized.byteLength,
    ) as ArrayBuffer,
  );

  const styleTable = new StyleTable();
  const sheets: Record<string, Partial<IWorksheetData>> = {};
  const sheetOrder: string[] = [];
  const sheetNames: string[] = [];
  const filterRefs = new Map<string, string>();

  workbook.eachSheet((ws, index) => {
    const sheetId = `sheet-${index}`;
    sheets[sheetId] = convertSheet(ws, sheetId, styleTable);
    sheetOrder.push(sheetId);
    sheetNames.push(ws.name);

    // ExcelJS exposes the autofilter as the bare `ref` string it read.
    if (typeof ws.autoFilter === 'string' && ws.autoFilter) {
      filterRefs.set(sheetId, ws.autoFilter);
    }
  });

  if (sheetOrder.length === 0) {
    throw new Error('통합 문서에 시트가 없습니다.');
  }

  const snapshot: IWorkbookData = {
    id: `wb-${sheetOrder.length}-${name}`,
    name,
    appVersion: '0.1.0',
    locale: LocaleType.KO_KR,
    styles: styleTable.styles as Record<string, IStyleData>,
    sheetOrder,
    sheets,
    resources: buildFilterResource(filterRefs),
  };

  return { snapshot, sheetNames };
}
