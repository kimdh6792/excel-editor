/**
 * Round-trips a real workbook from disk, if one is pointed at.
 *
 * Synthetic fixtures only prove the code agrees with itself; a workbook written
 * by Excel or an export tool exercises the parts no fixture thinks to produce.
 * Point `EXCEL_EDITOR_TEST_FILE` at one to run this:
 *
 *   EXCEL_EDITOR_TEST_FILE=~/somewhere/real.xlsx pnpm test
 *
 * Without it the suite skips rather than fails, so nothing here depends on a
 * file that only exists on one machine.
 */
import { existsSync, readFileSync } from 'node:fs';
import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { exportXlsx } from './export';
import { importXlsx } from './import';

const path = process.env.EXCEL_EDITOR_TEST_FILE;
const present = Boolean(path) && existsSync(path!);

describe.skipIf(!present)('real workbook round trip', () => {
  it('opens, re-saves and reads back with every value intact', async () => {
    const bytes = new Uint8Array(readFileSync(path!));

    const { snapshot, sheetNames } = await importXlsx(bytes, 'real.xlsx');
    expect(sheetNames.length).toBeGreaterThan(0);

    // Collect what the import produced, per sheet.
    const before = snapshot.sheetOrder.map((id) => {
      const sheet = snapshot.sheets[id]!;
      const matrix = sheet.cellData as unknown as Record<
        string,
        Record<string, { v?: unknown; f?: unknown }>
      >;
      const values: Array<[string, unknown]> = [];
      for (const [r, row] of Object.entries(matrix)) {
        for (const [c, cell] of Object.entries(row)) {
          if (cell.v !== undefined) values.push([`${r},${c}`, cell.v]);
        }
      }
      return { name: sheet.name!, values: values.sort() };
    });

    expect(before.some((s) => s.values.length > 0)).toBe(true);

    // Save, then read the saved bytes back through the same importer.
    const saved = await exportXlsx(snapshot);
    const { snapshot: after } = await importXlsx(saved, 'real.xlsx');

    const afterBySheet = after.sheetOrder.map((id) => {
      const sheet = after.sheets[id]!;
      const matrix = sheet.cellData as unknown as Record<string, Record<string, { v?: unknown }>>;
      const values: Array<[string, unknown]> = [];
      for (const [r, row] of Object.entries(matrix)) {
        for (const [c, cell] of Object.entries(row)) {
          if (cell.v !== undefined) values.push([`${r},${c}`, cell.v]);
        }
      }
      return { name: sheet.name!, values: values.sort() };
    });

    expect(afterBySheet.map((s) => s.name)).toEqual(before.map((s) => s.name));
    for (let i = 0; i < before.length; i++) {
      expect(afterBySheet[i].values).toEqual(before[i].values);
    }
  });

  it('produces a file Excel-compatible readers can parse', async () => {
    const bytes = new Uint8Array(readFileSync(path!));
    const { snapshot } = await importXlsx(bytes, 'real.xlsx');
    const saved = await exportXlsx(snapshot);

    // A stock ExcelJS instance, with no normalisation, must accept our output.
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(
      saved.buffer.slice(saved.byteOffset, saved.byteOffset + saved.byteLength) as ArrayBuffer,
    );
    expect(wb.worksheets.length).toBeGreaterThan(0);
    expect(wb.worksheets[0].rowCount).toBeGreaterThan(0);
  });
});
