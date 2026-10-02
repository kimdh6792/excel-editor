/**
 * Round-trip fidelity: build a workbook that exercises every feature the app
 * claims to preserve, push it through import -> export, and read it back with a
 * fresh ExcelJS instance to check what actually survived.
 */
import ExcelJS from 'exceljs';
import { strFromU8, unzipSync } from 'fflate';
import { beforeAll, describe, expect, it } from 'vitest';
import type { ICellData, IWorkbookData } from '@univerjs/core';
import { importXlsx } from './import';
import { exportXlsx } from './export';
import { dateToSerial } from './units';

const SHEET = '데이터';
const HIDDEN_SHEET = '숨김';

/** Builds the fixture workbook as xlsx bytes. */
async function buildFixture(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();

  const ws = wb.addWorksheet(SHEET, {
    views: [{ state: 'frozen', xSplit: 1, ySplit: 1, topLeftCell: 'B2' }],
    properties: { tabColor: { argb: 'FF00B050' } } as ExcelJS.WorksheetProperties,
  });

  // A1:C1 — merged, heavily styled title.
  const title = ws.getCell('A1');
  title.value = '제목';
  title.font = { name: 'Arial', size: 14, bold: true, color: { argb: 'FFC00000' } };
  title.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
  title.alignment = { horizontal: 'center', vertical: 'middle' };
  title.border = {
    top: { style: 'thin', color: { argb: 'FF000000' } },
    left: { style: 'thin', color: { argb: 'FF000000' } },
    bottom: { style: 'double', color: { argb: 'FF0070C0' } },
    right: { style: 'thick', color: { argb: 'FF000000' } },
  };
  ws.mergeCells('A1:C1');

  ws.getCell('A2').value = '텍스트';

  const number = ws.getCell('B2');
  number.value = 1234.5;
  number.numFmt = '#,##0.00';

  const date = ws.getCell('C2');
  date.value = new Date(Date.UTC(2026, 2, 15));
  date.numFmt = 'yyyy-mm-dd';

  ws.getCell('A3').value = { formula: 'B2*2', result: 2469 };
  ws.getCell('B3').value = true;

  ws.getCell('C3').value = {
    richText: [
      { text: '빨강', font: { color: { argb: 'FFFF0000' } } },
      { text: '파랑', font: { color: { argb: 'FF0000FF' } } },
    ],
  };

  ws.getCell('A4').value = { text: '앤트로픽', hyperlink: 'https://anthropic.com' };

  const decorated = ws.getCell('B4');
  decorated.value = '장식';
  decorated.font = { italic: true, underline: 'double', strike: true };

  const wrapped = ws.getCell('C4');
  wrapped.value = '줄바꿈';
  wrapped.alignment = { wrapText: true, textRotation: 45, indent: 2 };

  ws.getColumn(1).width = 20;
  ws.getColumn(2).width = 12.5;
  ws.getColumn(3).hidden = true;
  ws.getRow(2).height = 30;
  // The hidden row needs a cell: ExcelJS cannot serialise a row that has
  // neither cells nor a height, so a fixture without one would be testing
  // nothing. The empty-hidden-row case is covered separately below.
  ws.getCell('A5').value = '숨은 행';
  ws.getRow(5).hidden = true;

  const hidden = wb.addWorksheet(HIDDEN_SHEET, { state: 'hidden' });
  hidden.getCell('A1').value = '숨은 값';

  return new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer);
}

/** Reads bytes back into a fresh ExcelJS workbook. */
async function reload(bytes: Uint8Array): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(
    bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
  );
  return wb;
}

describe('xlsx round trip', () => {
  let snapshot: IWorkbookData;
  let out: ExcelJS.Workbook;
  let ws: ExcelJS.Worksheet;

  beforeAll(async () => {
    const fixture = await buildFixture();
    snapshot = (await importXlsx(fixture, 'fixture.xlsx')).snapshot;
    out = await reload(await exportXlsx(snapshot));
    ws = out.getWorksheet(SHEET)!;
  });

  /* ------------------------------------------------------------- structure */

  it('keeps both sheets, in order, with the hidden one hidden', () => {
    expect(snapshot.sheetOrder).toHaveLength(2);
    expect(out.worksheets.map((w) => w.name)).toEqual([SHEET, HIDDEN_SHEET]);
    expect(out.getWorksheet(SHEET)!.state).toBe('visible');
    expect(out.getWorksheet(HIDDEN_SHEET)!.state).toBe('hidden');
    expect(out.getWorksheet(HIDDEN_SHEET)!.getCell('A1').value).toBe('숨은 값');
  });

  it('deduplicates styles into the shared table', () => {
    const ids = Object.keys(snapshot.styles);
    expect(ids.length).toBeGreaterThan(0);
    // Every id a cell points at must exist in the table.
    const sheet = snapshot.sheets[snapshot.sheetOrder[0]]!;
    const matrix = sheet.cellData as unknown as Record<string, Record<string, ICellData>>;
    for (const row of Object.values(matrix)) {
      for (const cell of Object.values(row)) {
        if (typeof cell.s === 'string') expect(ids).toContain(cell.s);
      }
    }
  });

  /* ----------------------------------------------------------------- values */

  it('preserves plain text, numbers and number formats', () => {
    expect(ws.getCell('A2').value).toBe('텍스트');
    expect(ws.getCell('B2').value).toBe(1234.5);
    expect(ws.getCell('B2').numFmt).toBe('#,##0.00');
  });

  it('preserves dates as serials with their format', () => {
    const cell = ws.getCell('C2');
    expect(cell.numFmt).toBe('yyyy-mm-dd');
    // ExcelJS re-hydrates date-formatted numbers into Date objects.
    const serial = cell.value instanceof Date ? dateToSerial(cell.value) : Number(cell.value);
    expect(serial).toBeCloseTo(dateToSerial(new Date(Date.UTC(2026, 2, 15))), 4);
  });

  it('preserves formulas and their cached results', () => {
    const cell = ws.getCell('A3');
    expect(cell.formula).toBe('B2*2');
    expect(cell.result).toBe(2469);
  });

  it('asks Excel to recalculate on open', async () => {
    // ExcelJS writes `fullCalcOnLoad` but never parses it back, so this has to
    // be checked against the XML rather than a reloaded workbook object.
    const bytes = await exportXlsx(snapshot);
    const workbookXml = strFromU8(unzipSync(bytes)['xl/workbook.xml']);
    expect(workbookXml).toMatch(/<calcPr[^>]*fullCalcOnLoad="1"/);
  });

  it('preserves booleans', () => {
    expect(ws.getCell('B3').value).toBe(true);
  });

  it('restores rich text runs for cells that were not edited', () => {
    const value = ws.getCell('C3').value as { richText: Array<{ text: string; font?: object }> };
    expect(value.richText).toHaveLength(2);
    expect(value.richText.map((r) => r.text)).toEqual(['빨강', '파랑']);
    expect(value.richText[0].font).toMatchObject({ color: { argb: 'FFFF0000' } });
  });

  it('restores hyperlinks for cells that were not edited', () => {
    const value = ws.getCell('A4').value as { hyperlink: string; text: string };
    expect(value.hyperlink).toBe('https://anthropic.com');
    expect(value.text).toBe('앤트로픽');
  });

  it('drops rich text formatting once the text itself changes', async () => {
    // Simulate the user retyping C3: the preserved runs no longer match, so the
    // cell must be written as plain text rather than the stale runs.
    const edited: IWorkbookData = JSON.parse(JSON.stringify(snapshot));
    const sheet = edited.sheets[edited.sheetOrder[0]]!;
    const matrix = sheet.cellData as unknown as Record<string, Record<string, ICellData>>;
    matrix['2']['2'].v = '수정됨';

    const reloaded = await reload(await exportXlsx(edited));
    expect(reloaded.getWorksheet(SHEET)!.getCell('C3').value).toBe('수정됨');
  });

  /* ----------------------------------------------------------------- styles */

  it('preserves font attributes', () => {
    const font = ws.getCell('A1').font;
    expect(font.bold).toBe(true);
    expect(font.size).toBe(14);
    expect(font.name).toBe('Arial');
    expect(font.color?.argb).toBe('FFC00000');
  });

  it('preserves italic, double underline and strikethrough', () => {
    const font = ws.getCell('B4').font;
    expect(font.italic).toBe(true);
    expect(font.underline).toBe('double');
    expect(font.strike).toBe(true);
  });

  it('preserves solid fills', () => {
    const fill = ws.getCell('A1').fill as ExcelJS.FillPattern;
    expect(fill.type).toBe('pattern');
    expect(fill.pattern).toBe('solid');
    expect(fill.fgColor?.argb).toBe('FFFFFF00');
  });

  it('preserves per-edge border styles and colors', () => {
    const border = ws.getCell('A1').border;
    expect(border.top?.style).toBe('thin');
    expect(border.left?.style).toBe('thin');
    expect(border.bottom?.style).toBe('double');
    expect(border.bottom?.color?.argb).toBe('FF0070C0');
    expect(border.right?.style).toBe('thick');
  });

  it('preserves alignment, wrapping, rotation and indent', () => {
    expect(ws.getCell('A1').alignment).toMatchObject({
      horizontal: 'center',
      vertical: 'middle',
    });
    expect(ws.getCell('C4').alignment).toMatchObject({
      wrapText: true,
      textRotation: 45,
      indent: 2,
    });
  });

  /* ------------------------------------------------------- layout & merges */

  it('preserves the merge, and does not duplicate the value into the slaves', () => {
    expect(ws.model.merges).toContain('A1:C1');
    expect(ws.getCell('A1').value).toBe('제목');

    // In the snapshot the covered cells must carry style only, never a value.
    const sheet = snapshot.sheets[snapshot.sheetOrder[0]]!;
    const matrix = sheet.cellData as unknown as Record<string, Record<string, ICellData>>;
    expect(matrix['0']['1']?.v).toBeUndefined();
    expect(matrix['0']['2']?.v).toBeUndefined();
  });

  it('preserves column widths exactly, including fractional ones', () => {
    expect(ws.getColumn(1).width).toBe(20);
    // 12.5 chars does not land on a whole pixel; the original value is carried
    // through `custom` so repeated saves cannot drift.
    expect(ws.getColumn(2).width).toBe(12.5);
  });

  it('preserves hidden columns and rows', () => {
    expect(ws.getColumn(3).hidden).toBe(true);
    expect(ws.getRow(5).hidden).toBe(true);
    expect(ws.getCell('A5').value).toBe('숨은 행');
  });

  it('preserves a hidden row that holds nothing at all', async () => {
    // ExcelJS drops rows with no cells and no height, so export pins the default
    // height on such rows to keep the hidden flag. Without that workaround this
    // row vanishes entirely.
    const edited: IWorkbookData = JSON.parse(JSON.stringify(snapshot));
    const sheet = edited.sheets[edited.sheetOrder[0]]!;
    (sheet.rowData as Record<string, { hd?: number }>)['9'] = { hd: 1 };

    const reloaded = await reload(await exportXlsx(edited));
    expect(reloaded.getWorksheet(SHEET)!.getRow(10).hidden).toBe(true);
  });

  it('preserves row heights exactly', () => {
    expect(ws.getRow(2).height).toBe(30);
  });

  it('preserves frozen panes', () => {
    const view = ws.views[0];
    expect(view.state).toBe('frozen');
    expect((view as ExcelJS.WorksheetViewFrozen).xSplit).toBe(1);
    expect((view as ExcelJS.WorksheetViewFrozen).ySplit).toBe(1);
  });

  it('preserves the sheet tab color', () => {
    expect((ws.properties.tabColor as { argb?: string })?.argb).toBe('FF00B050');
  });

  /* ------------------------------------------------------------- stability */

  it('is stable across repeated saves', async () => {
    const once = await exportXlsx(snapshot);
    const reimported = (await importXlsx(once, 'fixture.xlsx')).snapshot;
    const twice = await reload(await exportXlsx(reimported));
    const ws2 = twice.getWorksheet(SHEET)!;

    expect(ws2.getColumn(1).width).toBe(20);
    expect(ws2.getColumn(2).width).toBe(12.5);
    expect(ws2.getRow(2).height).toBe(30);
    expect(ws2.getCell('A1').value).toBe('제목');
    expect(ws2.getCell('B2').numFmt).toBe('#,##0.00');
    expect(ws2.model.merges).toContain('A1:C1');
  });
});
