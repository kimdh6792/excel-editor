/**
 * Validates the snapshot against Univer's own model classes.
 *
 * The other tests check that import and export agree with each other, which
 * would still pass if the snapshot used the wrong key names or 1-based indices
 * throughout. Feeding it to Univer's `Workbook` and reading back through
 * Univer's accessors is what proves the shape is actually the one the grid
 * expects.
 */
import ExcelJS from 'exceljs';
import { Workbook } from '@univerjs/core';
import type { ILogService } from '@univerjs/core';
import { beforeAll, describe, expect, it } from 'vitest';
import { importXlsx } from './import';
import { createEmptySnapshot } from './snapshot';
import { charsToPixels, pointsToPixels } from './units';

/** Univer's Workbook wants a logger; nothing here needs its output. */
const silentLog: ILogService = {
  debug: () => {},
  log: () => {},
  warn: () => {},
  error: () => {},
  deprecate: () => {},
  setLogLevel: () => {},
};

async function fixture(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('시트하나');

  const header = ws.getCell('B2');
  header.value = '머리글';
  header.font = { bold: true, size: 13, color: { argb: 'FFC00000' } };
  header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
  ws.mergeCells('B2:D2');

  ws.getCell('B3').value = 10;
  ws.getCell('C3').value = 20;
  ws.getCell('D3').value = { formula: 'B3+C3', result: 30 };

  ws.getColumn(2).width = 25;
  ws.getRow(2).height = 36;

  wb.addWorksheet('시트둘').getCell('A1').value = '둘';

  return new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer);
}

describe('snapshot is valid Univer model data', () => {
  let workbook: Workbook;

  beforeAll(async () => {
    const { snapshot } = await importXlsx(await fixture(), 'model.xlsx');
    workbook = new Workbook(snapshot, silentLog);
  });

  it('exposes the sheets in file order under their real names', () => {
    expect(workbook.getSheets().map((s) => s.getName())).toEqual(['시트하나', '시트둘']);
  });

  it('places values at the 0-based coordinates Univer reads', () => {
    const sheet = workbook.getSheetBySheetName('시트하나')!;
    // B2 is row 1, column 1 for Univer.
    expect(sheet.getCell(1, 1)?.v).toBe('머리글');
    expect(sheet.getCell(2, 1)?.v).toBe(10);
    expect(sheet.getCell(2, 2)?.v).toBe(20);
  });

  it('keeps formulas in the field Univer evaluates', () => {
    const sheet = workbook.getSheetBySheetName('시트하나')!;
    const cell = sheet.getCell(2, 3);
    expect(cell?.f).toBe('=B3+C3');
    expect(cell?.v).toBe(30);
  });

  it('resolves style ids through the workbook style table', () => {
    const sheet = workbook.getSheetBySheetName('시트하나')!;
    // getCellStyle walks the shared table, which only works if the id matches.
    const style = sheet.getCellStyle(1, 1);
    expect(style?.bl).toBe(1);
    expect(style?.fs).toBe(13);
    expect(style?.cl?.rgb).toBe('#c00000');
    expect(style?.bg?.rgb).toBe('#fff2cc');
  });

  it('reports the merge as a 0-based inclusive range', () => {
    const sheet = workbook.getSheetBySheetName('시트하나')!;
    expect(sheet.getMergeData()).toEqual([
      { startRow: 1, startColumn: 1, endRow: 1, endColumn: 3 },
    ]);
  });

  it('reports column widths and row heights in pixels', () => {
    const sheet = workbook.getSheetBySheetName('시트하나')!;
    expect(sheet.getColumnWidth(1)).toBe(charsToPixels(25));
    expect(sheet.getRowHeight(1)).toBe(pointsToPixels(36));
  });

  it('accepts the empty snapshot used for a new file', () => {
    const empty = new Workbook(createEmptySnapshot(), silentLog);
    const sheet = empty.getSheets()[0];
    expect(sheet.getName()).toBe('Sheet1');
    expect(sheet.getRowCount()).toBeGreaterThan(0);
    expect(sheet.getColumnCount()).toBeGreaterThan(0);
  });
});
