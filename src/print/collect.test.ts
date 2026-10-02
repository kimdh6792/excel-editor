/**
 * Which cells a print run covers.
 *
 * The workbook is faked rather than driven through Univer: what is being tested
 * is the choice of area and the order of sheets, and a fake records exactly
 * which range was asked for, which the real facade does not.
 */
import { BooleanNumber } from '@univerjs/core';
import type { IRange, IWorkbookData } from '@univerjs/core';
import { describe, expect, it } from 'vitest';
import { collectPrintSheets } from './collect';
import type { PrintableWorkbook } from './collect';

interface FakeSheet {
  id: string;
  name: string;
  lastRow: number;
  lastColumn: number;
  hidden?: boolean;
  /** Text of every cell, so a blank sheet can be distinguished. */
  fill?: string;
}

/** Ranges the fake was asked for, newest last, as `sheetId:r,c+rows,cols`. */
type Asked = string[];

function fake(
  sheets: FakeSheet[],
  options: { activeId?: string; selection?: IRange } = {},
): { workbook: PrintableWorkbook; asked: Asked } {
  const asked: Asked = [];
  const activeId = options.activeId ?? sheets[0].id;

  const snapshot: IWorkbookData = {
    id: 'wb',
    name: 'wb',
    appVersion: '0.1.0',
    locale: 'koKR' as IWorkbookData['locale'],
    styles: {},
    sheetOrder: sheets.map((s) => s.id),
    sheets: Object.fromEntries(
      sheets.map((s) => [
        s.id,
        {
          id: s.id,
          name: s.name,
          rowCount: 100,
          columnCount: 20,
          hidden: s.hidden ? BooleanNumber.TRUE : BooleanNumber.FALSE,
        },
      ]),
    ),
  };

  const worksheet = (s: FakeSheet) => ({
    getSheetId: () => s.id,
    getSheetName: () => s.name,
    getLastRow: () => s.lastRow,
    getLastColumn: () => s.lastColumn,
    getRange: (row: number, column: number, rows: number, columns: number) => {
      asked.push(`${s.id}:${row},${column}+${rows},${columns}`);
      return {
        getRange: () => ({ startRow: row, startColumn: column, endRow: 0, endColumn: 0 }),
        getDisplayValues: () =>
          Array.from({ length: rows }, () =>
            Array.from({ length: columns }, () => s.fill ?? 'x'),
          ),
      };
    },
  });

  const workbook: PrintableWorkbook = {
    save: () => snapshot,
    getSheets: () => sheets.map(worksheet),
    getActiveSheet: () => worksheet(sheets.find((s) => s.id === activeId) ?? sheets[0]),
    getActiveRange: () =>
      options.selection
        ? {
            getRange: () => options.selection as IRange,
            getDisplayValues: () => [['sel']],
          }
        : null,
  };

  return { workbook, asked };
}

const ONE: FakeSheet[] = [{ id: 's1', name: 'Sheet1', lastRow: 4, lastColumn: 2 }];

describe('collectPrintSheets', () => {
  it('covers the active sheet’s data extent', () => {
    const { workbook, asked } = fake(ONE);
    const sheets = collectPrintSheets(workbook, 'sheet');
    expect(sheets).toHaveLength(1);
    expect(sheets[0].name).toBe('Sheet1');
    expect(asked).toEqual(['s1:0,0+5,3']);
    expect(sheets[0].rows).toHaveLength(5);
    expect(sheets[0].rows[0]).toHaveLength(3);
  });

  it('prints the active sheet when the second tab is selected', () => {
    const { workbook, asked } = fake(
      [
        { id: 's1', name: 'Sheet1', lastRow: 4, lastColumn: 2 },
        { id: 's2', name: 'Sheet2', lastRow: 1, lastColumn: 1 },
      ],
      { activeId: 's2' },
    );
    const sheets = collectPrintSheets(workbook, 'sheet');
    expect(sheets.map((s) => s.name)).toEqual(['Sheet2']);
    expect(asked).toEqual(['s2:0,0+2,2']);
  });

  it('asks for exactly the selected range', () => {
    const { workbook, asked } = fake(ONE, {
      selection: { startRow: 2, endRow: 6, startColumn: 1, endColumn: 3 },
    });
    const sheets = collectPrintSheets(workbook, 'selection');
    expect(asked).toEqual(['s1:2,1+5,3']);
    expect(sheets[0].rows).toHaveLength(5);
  });

  it('falls back to the whole sheet when the selection is one cell', () => {
    // Univer reports A1 on an untouched grid, and printing one cell because the
    // user never clicked anywhere would be a surprise.
    const { workbook, asked } = fake(ONE, {
      selection: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    });
    collectPrintSheets(workbook, 'selection');
    expect(asked).toEqual(['s1:0,0+5,3']);
  });

  it('falls back to the whole sheet when there is no selection at all', () => {
    const { workbook, asked } = fake(ONE);
    collectPrintSheets(workbook, 'selection');
    expect(asked).toEqual(['s1:0,0+5,3']);
  });

  it('prints every sheet in tab order', () => {
    const { workbook } = fake([
      { id: 's1', name: '1월', lastRow: 1, lastColumn: 1 },
      { id: 's2', name: '2월', lastRow: 1, lastColumn: 1 },
      { id: 's3', name: '3월', lastRow: 1, lastColumn: 1 },
    ]);
    expect(collectPrintSheets(workbook, 'all').map((s) => s.name)).toEqual(['1월', '2월', '3월']);
  });

  it('follows the snapshot’s tab order, not the facade’s', () => {
    const { workbook } = fake([
      { id: 's1', name: 'first', lastRow: 1, lastColumn: 1 },
      { id: 's2', name: 'second', lastRow: 1, lastColumn: 1 },
    ]);
    const snapshot = workbook.save();
    snapshot.sheetOrder = ['s2', 's1'];
    expect(collectPrintSheets(workbook, 'all').map((s) => s.name)).toEqual(['second', 'first']);
  });

  it('skips hidden sheets', () => {
    const { workbook } = fake([
      { id: 's1', name: 'shown', lastRow: 1, lastColumn: 1 },
      { id: 's2', name: 'hidden', lastRow: 1, lastColumn: 1, hidden: true },
    ]);
    expect(collectPrintSheets(workbook, 'all').map((s) => s.name)).toEqual(['shown']);
  });

  it('skips blank sheets rather than spending a page on each', () => {
    const { workbook } = fake([
      { id: 's1', name: 'data', lastRow: 1, lastColumn: 1 },
      { id: 's2', name: 'blank', lastRow: 1, lastColumn: 1, fill: '' },
    ]);
    expect(collectPrintSheets(workbook, 'all').map((s) => s.name)).toEqual(['data']);
  });

  it('still prints one page when every sheet is blank', () => {
    const { workbook } = fake([
      { id: 's1', name: 'a', lastRow: 1, lastColumn: 1, fill: '' },
      { id: 's2', name: 'b', lastRow: 1, lastColumn: 1, fill: '' },
    ]);
    expect(collectPrintSheets(workbook, 'all').map((s) => s.name)).toEqual(['a']);
  });

  it('returns nothing when the active sheet is missing from the snapshot', () => {
    const { workbook } = fake(ONE);
    const snapshot = workbook.save();
    delete snapshot.sheets.s1;
    expect(collectPrintSheets(workbook, 'sheet')).toEqual([]);
  });

  it('handles an empty sheet as a single cell', () => {
    const { workbook, asked } = fake([{ id: 's1', name: 'Sheet1', lastRow: 0, lastColumn: 0 }]);
    const sheets = collectPrintSheets(workbook, 'sheet');
    expect(asked).toEqual(['s1:0,0+1,1']);
    expect(sheets[0].rows).toEqual([[expect.objectContaining({ text: 'x' })]]);
  });
});
