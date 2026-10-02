/**
 * Import to printable markup, end to end.
 *
 * The other print tests work from hand-written snapshots, which proves the
 * modules agree with each other but not that they agree with what the importer
 * actually produces. This one starts from a real xlsx: a style the importer
 * stores on the shared table, a merge it parses out of `model.merges`, a hidden
 * row and column — all the places where the two halves could drift apart.
 *
 * Display text is Univer's job and is not available here, so values are passed
 * through as the importer stored them. Formatting is covered where it belongs,
 * in the importer's own tests.
 */
import ExcelJS from 'exceljs';
import type { IStyleData, Nullable } from '@univerjs/core';
import { describe, expect, it } from 'vitest';
import { importXlsx } from '../xlsx/import';
import { buildPrintSheet } from './area';
import { buildPrintHtml } from './html';
import { DEFAULT_PRINT_SETTINGS } from './options';

async function fixture(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('매출');

  ws.mergeCells('A1:D1');
  const title = ws.getCell('A1');
  title.value = '3분기 매출 <요약>';
  title.font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } };
  title.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF305496' } };
  title.alignment = { horizontal: 'center', vertical: 'middle' };

  ['지점', '매출', '증감률', '비고'].forEach((text, index) => {
    const cell = ws.getRow(2).getCell(index + 1);
    cell.value = text;
    cell.font = { bold: true };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
  });

  ws.getCell('A3').value = '서울';
  ws.getCell('B3').value = 128450000;
  ws.getCell('B3').numFmt = '#,##0';
  ws.getCell('D3').value = '신규 채널\n이월 포함';
  ws.getCell('D3').alignment = { wrapText: true };

  // Hidden, and holding data so the importer records the hidden flag at all.
  ws.getCell('A4').value = '숨긴지점';
  ws.getRow(4).hidden = true;
  ws.getCell('E2').value = '내부메모';
  ws.getColumn(5).width = 12;
  ws.getColumn(5).hidden = true;

  ws.getCell('A5').value = '합계';
  ws.getCell('B5').value = 128450000;
  ws.getCell('B5').border = { top: { style: 'double' } };

  ws.getColumn(1).width = 10;
  ws.getColumn(2).width = 20;

  return new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer);
}

/** Raw stored values as text, standing in for Univer's formatter. */
function storedText(
  cellData: Record<number, Record<number, { v?: unknown }>> | undefined,
  rows: number,
  columns: number,
): string[][] {
  return Array.from({ length: rows }, (_, r) =>
    Array.from({ length: columns }, (_, c) => {
      const value = cellData?.[r]?.[c]?.v;
      return value == null ? '' : String(value);
    }),
  );
}

async function render(): Promise<{ html: string; columns: number; rows: number }> {
  const { snapshot } = await importXlsx(await fixture(), 'print.xlsx');
  const sheet = snapshot.sheets[snapshot.sheetOrder[0]];
  const area = { startRow: 0, endRow: 4, startColumn: 0, endColumn: 4 };

  const built = buildPrintSheet({
    name: sheet.name ?? 'Sheet1',
    sheet,
    styles: snapshot.styles as Record<string, Nullable<IStyleData>>,
    values: storedText(sheet.cellData as Record<number, Record<number, { v?: unknown }>>, 5, 5),
    area,
  });

  return {
    html: buildPrintHtml([built], { ...DEFAULT_PRINT_SETTINGS, repeatHeaderRows: 2 }),
    columns: built.columnWidths.length,
    rows: built.rows.length,
  };
}

describe('xlsx to printable markup', () => {
  it('leaves hidden rows and columns off the page', async () => {
    const { html, columns, rows } = await render();
    expect(html).not.toContain('숨긴지점');
    expect(html).not.toContain('내부메모');
    expect(columns).toBe(4);
    expect(rows).toBe(4);
  });

  it('carries the merge across as a span', async () => {
    const { html } = await render();
    expect(html).toMatch(/<td colspan="4"[^>]*>3분기 매출 &lt;요약&gt;<\/td>/);
  });

  it('resolves styles the importer put in the shared table', async () => {
    const { html } = await render();
    // Title: white bold 16pt on the dark fill.
    expect(html).toContain('background-color:#305496');
    expect(html).toContain('color:#ffffff');
    expect(html).toContain('font-size:16pt');
    // Header row fill.
    expect(html).toContain('background-color:#d9e1f2');
  });

  it('keeps the column widths the file stated', async () => {
    const { html } = await render();
    const actualSize = buildPrintHtml(
      [
        {
          name: 'x',
          rows: [[{ text: '', numeric: false, rowSpan: 1, colSpan: 1 }]],
          rowHeights: [20],
          columnWidths: [75],
          showGridlines: true,
        },
      ],
      { ...DEFAULT_PRINT_SETTINGS, fitToWidth: false },
    );
    expect(actualSize).toContain('width:75px');
    // Fitted to the page, the proportions are what survive.
    expect(html).toContain('%">');
  });

  it('marks a wrapped cell and keeps its line break', async () => {
    const { html } = await render();
    expect(html).toContain('white-space:pre-wrap');
    expect(html).toContain('신규 채널\n이월 포함');
  });

  it('draws the double rule under the total', async () => {
    const { html } = await render();
    expect(html).toContain('border-top:3px double');
  });

  it('repeats the title and header rows in the thead', async () => {
    const { html } = await render();
    const thead = html.match(/<thead>(.*?)<\/thead>/s)?.[1] ?? '';
    expect(thead).toContain('3분기 매출');
    expect(thead).toContain('증감률');
    expect(thead).not.toContain('서울');
  });
});
