/**
 * Delimited text <-> Univer snapshot.
 *
 * Type inference is deliberately conservative. Turning text that merely looks
 * numeric into a number is how spreadsheets destroy data: phone numbers lose
 * their leading zero, postal codes lose theirs, long IDs get rounded. The rules
 * below only accept values that are unambiguously numeric and representable
 * exactly, and never guess at dates.
 */
import { CellValueType, LocaleType } from '@univerjs/core';
import type { ICellData, IStyleData, IWorkbookData, Nullable } from '@univerjs/core';
import { resolveStyle } from '../xlsx/style';
import { serialToDate } from '../xlsx/units';
import {
  DEFAULT_COLUMN_WIDTH_CHARS,
  DEFAULT_ROW_HEIGHT_POINTS,
  charsToPixels,
  pointsToPixels,
} from '../xlsx/units';
import { MIN_COLS, MIN_ROWS } from '../xlsx/snapshot';
import { decodeText, encodeTextWithBom } from './encoding';
import type { DetectedEncoding } from './encoding';
import { parseDelimited, serialiseDelimited, sniffDelimiter } from './delimited';
import type { Delimiter } from './delimited';

/**
 * A number literal with no leading zeros. Leading zeros are the signal that the
 * text is an identifier rather than a quantity.
 */
const NUMERIC = /^-?(0|[1-9]\d*)(\.\d+)?([eE][+-]?\d+)?$/;

/** Digits a float64 can hold without losing precision. */
const MAX_EXACT_DIGITS = 15;

/** Converts one raw field to the value to store, leaving anything doubtful as text. */
export function inferValue(raw: string): string | number {
  if (raw === '' || !NUMERIC.test(raw)) return raw;

  const significant = raw.replace(/[^0-9]/g, '').replace(/^0+/, '');
  if (significant.length > MAX_EXACT_DIGITS) return raw;

  const value = Number(raw);
  return Number.isFinite(value) ? value : raw;
}

export interface CsvImportResult {
  snapshot: IWorkbookData;
  encoding: DetectedEncoding;
  delimiter: Delimiter;
  rowCount: number;
  columnCount: number;
}

/** Reads delimited bytes into a single-sheet snapshot. */
export function csvToSnapshot(bytes: Uint8Array, name: string): CsvImportResult {
  const { text, encoding } = decodeText(bytes);
  const delimiter = sniffDelimiter(text);
  const rows = parseDelimited(text, delimiter);

  const cellData: Record<number, Record<number, ICellData>> = {};
  let maxCol = 0;

  rows.forEach((row, r) => {
    row.forEach((raw, c) => {
      if (raw === '') return;
      const value = inferValue(raw);
      (cellData[r] ??= {})[c] = {
        v: value,
        t: typeof value === 'number' ? CellValueType.NUMBER : CellValueType.STRING,
      };
      if (c > maxCol) maxCol = c;
    });
  });

  const sheetName = name.replace(/\.(csv|tsv|txt)$/i, '').slice(0, 31) || 'Sheet1';

  return {
    snapshot: {
      id: `wb-csv-${name}`,
      name,
      appVersion: '0.1.0',
      locale: LocaleType.KO_KR,
      styles: {},
      sheetOrder: ['sheet-1'],
      sheets: {
        'sheet-1': {
          id: 'sheet-1',
          name: sheetName,
          tabColor: '',
          hidden: 0,
          rowCount: Math.max(rows.length, MIN_ROWS),
          columnCount: Math.max(maxCol + 1, MIN_COLS),
          zoomRatio: 1,
          freeze: { xSplit: 0, ySplit: 0, startRow: -1, startColumn: -1 },
          scrollTop: 0,
          scrollLeft: 0,
          defaultColumnWidth: charsToPixels(DEFAULT_COLUMN_WIDTH_CHARS),
          defaultRowHeight: pointsToPixels(DEFAULT_ROW_HEIGHT_POINTS),
          mergeData: [],
          cellData,
          rowData: {},
          columnData: {},
          rowHeader: { width: 46 },
          columnHeader: { height: 20 },
          showGridlines: 1,
          rightToLeft: 0,
        },
      },
      resources: [],
    },
    encoding,
    delimiter,
    rowCount: rows.length,
    columnCount: maxCol + 1,
  };
}

/* ----------------------------------------------------------------- writing */

/**
 * True when a number format renders its value as a date or time.
 *
 * Bracketed sections (`[$-409]`, `[Red]`) and quoted literals are stripped first
 * so that a currency format carrying the letter `d` in a literal is not mistaken
 * for a date.
 */
export function isDateFormat(pattern: string | undefined): boolean {
  if (!pattern) return false;
  const stripped = pattern
    .replace(/"[^"]*"/g, '')
    .replace(/\[[^\]]*\]/g, '')
    .replace(/\\./g, '');
  return /[ymd]/i.test(stripped) || /h+[:.]mm/i.test(stripped);
}

function pad(n: number, width = 2): string {
  return String(n).padStart(width, '0');
}

/**
 * Renders a date serial as ISO-ish text. Only the presence of a time component
 * is taken from the pattern; the rest is normalised, because the point of a CSV
 * export is a value another tool can parse, not a visual copy.
 */
export function formatDateSerial(serial: number, pattern: string | undefined): string {
  const date = serialToDate(serial);
  if (Number.isNaN(date.getTime())) return String(serial);

  const ymd = `${date.getUTCFullYear()}-${pad(date.getUTCMonth() + 1)}-${pad(date.getUTCDate())}`;
  const wantsTime = Boolean(pattern && /h/i.test(pattern.replace(/"[^"]*"/g, '')));
  const hasTime =
    date.getUTCHours() !== 0 || date.getUTCMinutes() !== 0 || date.getUTCSeconds() !== 0;

  if (!wantsTime && !hasTime) return ymd;
  return `${ymd} ${pad(date.getUTCHours())}:${pad(date.getUTCMinutes())}:${pad(date.getUTCSeconds())}`;
}

/** Renders one cell as the text to write. */
function cellToText(
  cell: ICellData | undefined,
  styles: Record<string, Nullable<IStyleData>>,
): string {
  if (!cell || cell.v == null) return '';

  const pattern = resolveStyle(cell.s, styles)?.n?.pattern ?? undefined;
  if (typeof cell.v === 'number' && isDateFormat(pattern)) {
    return formatDateSerial(cell.v, pattern);
  }
  return String(cell.v);
}

export interface CsvExportOptions {
  /** Sheet to write; defaults to the first in `sheetOrder`. */
  sheetId?: string;
  delimiter?: string;
}

/** Serialises one sheet of a snapshot as delimited text (no BOM). */
export function snapshotToCsvText(snapshot: IWorkbookData, options: CsvExportOptions = {}): string {
  const sheetId = options.sheetId ?? snapshot.sheetOrder?.[0];
  const sheet = sheetId ? snapshot.sheets?.[sheetId] : undefined;
  if (!sheet) return '';

  const styles = snapshot.styles ?? {};
  const matrix = (sheet.cellData ?? {}) as unknown as Record<
    string,
    Record<string, ICellData>
  >;

  // Trim to the populated area so an empty 100x26 grid does not become 100 blank
  // lines of commas.
  let maxRow = -1;
  let maxCol = -1;
  for (const [rowKey, row] of Object.entries(matrix)) {
    const r = Number(rowKey);
    if (!Number.isFinite(r) || !row) continue;
    for (const [colKey, cell] of Object.entries(row)) {
      const c = Number(colKey);
      if (!Number.isFinite(c) || !cell || cell.v == null) continue;
      if (r > maxRow) maxRow = r;
      if (c > maxCol) maxCol = c;
    }
  }
  if (maxRow < 0) return '';

  const rows: string[][] = [];
  for (let r = 0; r <= maxRow; r++) {
    const row: string[] = [];
    for (let c = 0; c <= maxCol; c++) {
      row.push(cellToText(matrix[String(r)]?.[String(c)], styles));
    }
    rows.push(row);
  }

  return serialiseDelimited(rows, options.delimiter ?? ',');
}

/** Serialises one sheet as delimited bytes, UTF-8 with a BOM for Excel. */
export function snapshotToCsv(snapshot: IWorkbookData, options: CsvExportOptions = {}): Uint8Array {
  return encodeTextWithBom(snapshotToCsvText(snapshot, options));
}

/** Sheets that a CSV save would leave behind, for warning the user. */
export function sheetsLostInCsv(snapshot: IWorkbookData, sheetId?: string): string[] {
  const keep = sheetId ?? snapshot.sheetOrder?.[0];
  return (snapshot.sheetOrder ?? [])
    .filter((id) => id !== keep)
    .map((id) => snapshot.sheets?.[id]?.name ?? id);
}
