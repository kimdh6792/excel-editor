/**
 * Conversions between A1-style range references and Univer's `IRange`.
 *
 * Univer ranges are 0-based and inclusive on both ends; xlsx references are
 * 1-based letters-and-digits. Both merges and autofilters go through here.
 */
import type { IRange } from '@univerjs/core';
import { columnToLetter } from './units';

/** `AB` -> 27 (1-based). */
function lettersToColumn(letters: string): number {
  return letters
    .toUpperCase()
    .split('')
    .reduce((acc, ch) => acc * 26 + (ch.charCodeAt(0) - 64), 0);
}

/**
 * Parses `A1:B2`, `A1`, or an absolute form like `$A$1:$B$2` into a 0-based
 * inclusive range. Returns undefined for anything else, including the
 * sheet-qualified refs that can appear in defined names.
 */
export function parseRangeRef(ref: string): IRange | undefined {
  const match = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(ref.replace(/\$/g, '').toUpperCase());
  if (!match) return undefined;

  const startColumn = lettersToColumn(match[1]) - 1;
  const startRow = Number(match[2]) - 1;
  const endColumn = match[3] ? lettersToColumn(match[3]) - 1 : startColumn;
  const endRow = match[4] ? Number(match[4]) - 1 : startRow;

  if (endRow < startRow || endColumn < startColumn) return undefined;
  return { startRow, startColumn, endRow, endColumn };
}

/** Formats a 0-based inclusive range as `A1:B2`. */
export function formatRangeRef(range: IRange): string {
  const start = `${columnToLetter(range.startColumn)}${range.startRow + 1}`;
  const end = `${columnToLetter(range.endColumn)}${range.endRow + 1}`;
  return start === end ? start : `${start}:${end}`;
}
