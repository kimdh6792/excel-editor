/**
 * The flattening step, which is where the print rendering can silently lose or
 * duplicate cells: hidden rows, merges that cross the edge of the printed area
 * and styles stored on a row or column rather than a cell.
 */
import { BooleanNumber, HorizontalAlign } from '@univerjs/core';
import type { IStyleData } from '@univerjs/core';
import { describe, expect, it } from 'vitest';
import { buildPrintSheet, clampArea, isBlank } from './area';
import type { Area, SheetData } from './area';
import { charsToPixels, pointsToPixels } from '../xlsx/units';

const AREA: Area = { startRow: 0, endRow: 2, startColumn: 0, endColumn: 2 };

/** A 3x3 sheet whose cells read "r0c0", "r0c1", … so misindexing is visible. */
function grid(sheet: SheetData = {}): SheetData {
  return { rowCount: 10, columnCount: 10, ...sheet };
}

function values(rows: number, columns: number): string[][] {
  return Array.from({ length: rows }, (_, r) =>
    Array.from({ length: columns }, (_, c) => `r${r}c${c}`),
  );
}

function texts(sheet: SheetData, area = AREA, styles: Record<string, IStyleData> = {}) {
  const rows = area.endRow - area.startRow + 1;
  const columns = area.endColumn - area.startColumn + 1;
  const built = buildPrintSheet({
    name: 'Sheet1',
    sheet,
    styles,
    values: values(rows, columns),
    area,
  });
  return built.rows.map((row) => row.map((cell) => cell.text));
}

describe('clampArea', () => {
  it('pulls an area back inside the sheet', () => {
    const clamped = clampArea(
      { startRow: 5, endRow: 999, startColumn: 2, endColumn: 999 },
      { rowCount: 10, columnCount: 4 },
    );
    expect(clamped).toEqual({ startRow: 5, endRow: 9, startColumn: 2, endColumn: 3 });
  });

  it('never returns an inverted area', () => {
    const clamped = clampArea(
      { startRow: 50, endRow: 60, startColumn: 50, endColumn: 60 },
      { rowCount: 3, columnCount: 3 },
    );
    expect(clamped).toEqual({ startRow: 2, endRow: 2, startColumn: 2, endColumn: 2 });
  });

  it('leaves an axis alone when the snapshot states no count', () => {
    expect(clampArea({ startRow: 0, endRow: 7, startColumn: 0, endColumn: 4 }, {})).toEqual({
      startRow: 0,
      endRow: 7,
      startColumn: 0,
      endColumn: 4,
    });
  });
});

describe('buildPrintSheet', () => {
  it('reads display text in area-relative coordinates', () => {
    const built = buildPrintSheet({
      name: 'Sheet1',
      sheet: grid(),
      styles: {},
      values: [['A', 'B'], ['C', 'D']],
      area: { startRow: 4, endRow: 5, startColumn: 7, endColumn: 8 },
    });
    expect(built.rows.map((row) => row.map((cell) => cell.text))).toEqual([
      ['A', 'B'],
      ['C', 'D'],
    ]);
  });

  it('drops hidden rows and columns', () => {
    const sheet = grid({
      rowData: { 1: { hd: BooleanNumber.TRUE } },
      columnData: { 0: { hd: BooleanNumber.TRUE } },
    });
    expect(texts(sheet)).toEqual([
      ['r0c1', 'r0c2'],
      ['r2c1', 'r2c2'],
    ]);
  });

  it('keeps heights and widths in step with the rows that survived', () => {
    const sheet = grid({
      rowData: { 0: { h: 40 }, 1: { hd: BooleanNumber.TRUE }, 2: { h: 60 } },
      columnData: { 0: { w: 120 }, 1: { hd: BooleanNumber.TRUE }, 2: { w: 80 } },
    });
    const built = buildPrintSheet({
      name: 'Sheet1',
      sheet,
      styles: {},
      values: values(3, 3),
      area: AREA,
    });
    expect(built.rowHeights).toEqual([40, 60]);
    expect(built.columnWidths).toEqual([120, 80]);
    expect(built.rows).toHaveLength(2);
    expect(built.rows[0]).toHaveLength(2);
  });

  it('falls back to the sheet default, then Excel’s own default', () => {
    const withSheetDefault = buildPrintSheet({
      name: 'Sheet1',
      sheet: grid({ defaultRowHeight: 33, defaultColumnWidth: 111 }),
      styles: {},
      values: values(3, 3),
      area: AREA,
    });
    expect(withSheetDefault.rowHeights).toEqual([33, 33, 33]);
    expect(withSheetDefault.columnWidths).toEqual([111, 111, 111]);

    const bare = buildPrintSheet({
      name: 'Sheet1',
      sheet: grid(),
      styles: {},
      values: values(3, 3),
      area: AREA,
    });
    expect(bare.rowHeights[0]).toBe(pointsToPixels(15));
    expect(bare.columnWidths[0]).toBe(charsToPixels(8.43));
  });

  it('gives the merge anchor a span and omits the cells it swallows', () => {
    const sheet = grid({
      mergeData: [{ startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 }],
    });
    const built = buildPrintSheet({
      name: 'Sheet1',
      sheet,
      styles: {},
      values: values(3, 3),
      area: AREA,
    });
    expect(built.rows[0][0]).toMatchObject({ text: 'r0c0', rowSpan: 2, colSpan: 2 });
    // Row 0 keeps the anchor plus the column outside the merge; row 1 only the latter.
    expect(built.rows[0].map((c) => c.text)).toEqual(['r0c0', 'r0c2']);
    expect(built.rows[1].map((c) => c.text)).toEqual(['r1c2']);
  });

  it('counts a merge in visible rows, so a hidden row shortens the span', () => {
    const sheet = grid({
      rowData: { 1: { hd: BooleanNumber.TRUE } },
      mergeData: [{ startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 }],
    });
    const built = buildPrintSheet({
      name: 'Sheet1',
      sheet,
      styles: {},
      values: values(3, 3),
      area: AREA,
    });
    expect(built.rows[0][0]).toMatchObject({ text: 'r0c0', rowSpan: 2 });
  });

  it('clips a merge to the printed area and moves its anchor', () => {
    const sheet = grid({
      mergeData: [{ startRow: 0, endRow: 3, startColumn: 0, endColumn: 3 }],
    });
    const area: Area = { startRow: 1, endRow: 2, startColumn: 1, endColumn: 2 };
    const built = buildPrintSheet({
      name: 'Sheet1',
      sheet,
      styles: {},
      values: values(2, 2),
      area,
    });
    expect(built.rows).toHaveLength(2);
    expect(built.rows[0]).toHaveLength(1);
    expect(built.rows[0][0]).toMatchObject({ rowSpan: 2, colSpan: 2 });
    expect(built.rows[1]).toHaveLength(0);
  });

  it('ignores a merge that misses the printed area entirely', () => {
    const sheet = grid({
      mergeData: [{ startRow: 8, endRow: 9, startColumn: 8, endColumn: 9 }],
    });
    const built = buildPrintSheet({
      name: 'Sheet1',
      sheet,
      styles: {},
      values: values(3, 3),
      area: AREA,
    });
    expect(built.rows.every((row) => row.every((cell) => cell.rowSpan === 1))).toBe(true);
    expect(built.rows.flat()).toHaveLength(9);
  });

  it('resolves a style id against the shared table', () => {
    const sheet = grid({ cellData: { 0: { 0: { v: 1, s: 's1' } } } });
    const built = buildPrintSheet({
      name: 'Sheet1',
      sheet,
      styles: { s1: { ht: HorizontalAlign.CENTER } },
      values: values(3, 3),
      area: AREA,
    });
    expect(built.rows[0][0].style).toEqual({ ht: HorizontalAlign.CENTER });
  });

  it('prefers a cell style, then the row’s, then the column’s', () => {
    const sheet = grid({
      cellData: { 0: { 0: { v: 1, s: 'cell' } } },
      rowData: { 0: { s: 'row' }, 1: { s: 'row' } },
      columnData: { 0: { s: 'column' }, 1: { s: 'column' } },
    });
    const styles = {
      cell: { bl: BooleanNumber.TRUE },
      row: { it: BooleanNumber.TRUE },
      column: { ff: 'Courier' },
    };
    const built = buildPrintSheet({
      name: 'Sheet1',
      sheet,
      styles,
      values: values(3, 3),
      area: AREA,
    });
    expect(built.rows[0][0].style).toEqual(styles.cell);
    // Row 0 column 1 has no cell style, so the row's applies.
    expect(built.rows[0][1].style).toEqual(styles.row);
    // Row 2 has no row style, so column 0 falls through to the column's.
    expect(built.rows[2][0].style).toEqual(styles.column);
    expect(built.rows[2][2].style).toBeUndefined();
  });

  it('marks numeric cells, including formula results with no stated type', () => {
    const sheet = grid({
      cellData: {
        0: {
          0: { v: 42 },
          1: { v: '42' },
          2: { v: 7, f: '=1+6' },
        },
      },
    });
    const built = buildPrintSheet({
      name: 'Sheet1',
      sheet,
      styles: {},
      values: values(3, 3),
      area: AREA,
    });
    expect(built.rows[0].map((cell) => cell.numeric)).toEqual([true, false, true]);
  });

  it('reads gridline visibility from the sheet', () => {
    const shown = buildPrintSheet({
      name: 'S',
      sheet: grid({ showGridlines: 1 }),
      styles: {},
      values: values(3, 3),
      area: AREA,
    });
    const hidden = buildPrintSheet({
      name: 'S',
      sheet: grid({ showGridlines: 0 }),
      styles: {},
      values: values(3, 3),
      area: AREA,
    });
    expect(shown.showGridlines).toBe(true);
    expect(hidden.showGridlines).toBe(false);
  });

  it('tolerates a values grid smaller than the area', () => {
    const built = buildPrintSheet({
      name: 'Sheet1',
      sheet: grid(),
      styles: {},
      values: [['only']],
      area: AREA,
    });
    expect(built.rows[0].map((cell) => cell.text)).toEqual(['only', '', '']);
    expect(built.rows[2].map((cell) => cell.text)).toEqual(['', '', '']);
  });
});

describe('isBlank', () => {
  it('treats whitespace-only text as nothing to print', () => {
    const built = buildPrintSheet({
      name: 'Sheet1',
      sheet: grid(),
      styles: {},
      values: [['', '  '], ['', '']],
      area: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
    });
    expect(isBlank(built)).toBe(true);
  });

  it('is false as soon as one cell has text', () => {
    const built = buildPrintSheet({
      name: 'Sheet1',
      sheet: grid(),
      styles: {},
      values: [['', ''], ['', 'x']],
      area: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
    });
    expect(isBlank(built)).toBe(false);
  });
});
