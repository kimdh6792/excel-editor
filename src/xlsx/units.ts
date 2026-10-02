/**
 * Unit conversions between the xlsx file format and Univer's layout model.
 *
 * Excel stores column widths in "characters" of the default font and row
 * heights in points. Univer works in 96-DPI CSS pixels for both.
 */

/** Width in pixels of one character of the default 11pt Calibri grid font. */
const CHAR_WIDTH = 7;
/** Cell padding Excel adds on top of the character count. */
const CELL_PADDING = 5;

/** Excel's own defaults, used when the file does not state one. */
export const DEFAULT_COLUMN_WIDTH_CHARS = 8.43;
export const DEFAULT_ROW_HEIGHT_POINTS = 15;

export function charsToPixels(chars: number): number {
  return Math.round(chars * CHAR_WIDTH + CELL_PADDING);
}

export function pixelsToChars(px: number): number {
  // Round to 2dp: Excel writes widths with limited precision and exact
  // round-tripping of an untouched column matters more than sub-pixel accuracy.
  return Math.max(0, Math.round(((px - CELL_PADDING) / CHAR_WIDTH) * 100) / 100);
}

export function pointsToPixels(pt: number): number {
  return Math.round((pt * 96) / 72);
}

export function pixelsToPoints(px: number): number {
  return Math.round(((px * 72) / 96) * 100) / 100;
}

/**
 * Excel's day-zero offset for the 1900 date system, expressed as a serial
 * number. ExcelJS converts date-formatted cells into UTC `Date` objects, so the
 * inverse is a pure arithmetic shift.
 */
const EXCEL_EPOCH_OFFSET = 25569;
const MS_PER_DAY = 86400000;

export function dateToSerial(date: Date): number {
  return date.getTime() / MS_PER_DAY + EXCEL_EPOCH_OFFSET;
}

/** Inverse of {@link dateToSerial}. Shares its pre-1900-03-01 off-by-one. */
export function serialToDate(serial: number): Date {
  // Round to the nearest millisecond: serials are binary fractions of a day and
  // would otherwise land on 11:59:59.9997 instead of midnight.
  return new Date(Math.round((serial - EXCEL_EPOCH_OFFSET) * MS_PER_DAY));
}

/** Converts a 0-based column index to its spreadsheet letter (0 -> A). */
export function columnToLetter(index: number): string {
  let n = index;
  let out = '';
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
}
