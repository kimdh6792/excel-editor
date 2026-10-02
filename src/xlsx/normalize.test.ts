/**
 * The repair pass that lets ExcelJS read worksheets which omit the optional `r`
 * position attributes. Real exports do this (DataGrip, various server-side
 * writers), and without the pass ExcelJS rejects the entire file.
 */
import ExcelJS from 'exceljs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { importXlsx } from './import';
import { normalizeSheetXml, normalizeXlsx } from './normalize';

describe('normalizeSheetXml', () => {
  it('numbers rows and cells that state no position', () => {
    const xml =
      '<sheetData><row><c t="s"><v>0</v></c><c t="n"><v>1</v></c></row>' +
      '<row><c t="s"><v>2</v></c></row></sheetData>';
    const out = normalizeSheetXml(xml)!;

    expect(out).toContain('<row r="1">');
    expect(out).toContain('<c r="A1" t="s">');
    expect(out).toContain('<c r="B1" t="n">');
    expect(out).toContain('<row r="2">');
    expect(out).toContain('<c r="A2" t="s">');
  });

  it('leaves a fully-positioned sheet untouched', () => {
    const xml = '<sheetData><row r="1"><c r="A1" t="n"><v>1</v></c></row></sheetData>';
    expect(normalizeSheetXml(xml)).toBeUndefined();
  });

  it('continues from a stated position when only some are present', () => {
    const xml = '<sheetData><row r="7"><c><v>1</v></c><c><v>2</v></c></row></sheetData>';
    const out = normalizeSheetXml(xml)!;
    expect(out).toContain('<c r="A7">');
    expect(out).toContain('<c r="B7">');
  });

  it('resumes cell lettering after an explicitly placed cell', () => {
    const xml = '<sheetData><row r="1"><c r="C1"><v>1</v></c><c><v>2</v></c></row></sheetData>';
    const out = normalizeSheetXml(xml)!;
    expect(out).toContain('<c r="D1">');
  });

  it('handles self-closing cells and rows', () => {
    const xml = '<sheetData><row><c/><c><v>1</v></c></row><row/></sheetData>';
    const out = normalizeSheetXml(xml)!;
    expect(out).toContain('<c r="A1"/>');
    expect(out).toContain('<c r="B1">');
    expect(out).toContain('<row r="2"/>');
  });

  it('does not touch elements whose names merely start the same way', () => {
    const xml =
      '<cols><col min="1" max="1" width="9"/></cols><rowBreaks count="0"/>' +
      '<sheetData><row><c><v>1</v></c></row></sheetData>';
    const out = normalizeSheetXml(xml)!;
    expect(out).toContain('<col min="1" max="1" width="9"/>');
    expect(out).toContain('<rowBreaks count="0"/>');
    expect(out).toContain('<row r="1">');
  });

  it('beyond column Z, keeps letters in step', () => {
    const cells = Array.from({ length: 28 }, () => '<c><v>1</v></c>').join('');
    const out = normalizeSheetXml(`<sheetData><row>${cells}</row></sheetData>`)!;
    expect(out).toContain('<c r="Z1">');
    expect(out).toContain('<c r="AA1">');
    expect(out).toContain('<c r="AB1">');
  });
});

/** Removes every `r` attribute from row/cell tags, imitating a minimal writer. */
function stripPositions(xml: string): string {
  return xml.replace(
    /<(row|c)((?:\s+[\w:.-]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>/g,
    (_whole, name: string, attrs: string, selfClose: string) =>
      `<${name}${attrs.replace(/\s+r\s*=\s*(?:"[^"]*"|'[^']*')/, '')}${selfClose}>`,
  );
}

describe('normalizeXlsx', () => {
  it('makes a position-less workbook readable, with values intact', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Result 1');
    ws.getCell('A1').value = '지점명';
    ws.getCell('B1').value = '요일';
    ws.getCell('C1').value = 3;
    ws.getCell('A2').value = '서울지점';
    ws.getCell('B2').value = '월';
    ws.getCell('C2').value = 42;

    const original = new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer);

    // Rewrite the package the way a minimal writer would emit it.
    const entries = unzipSync(original);
    for (const path of Object.keys(entries)) {
      if (/^xl\/worksheets\/sheet[^/]*\.xml$/.test(path)) {
        entries[path] = strToU8(stripPositions(strFromU8(entries[path])));
      }
    }
    const stripped = zipSync(entries);

    // Confirm the premise: ExcelJS cannot read it as-is.
    await expect(
      new ExcelJS.Workbook().xlsx.load(stripped.buffer.slice(0) as ArrayBuffer),
    ).rejects.toThrow(/Invalid row number/);

    // After normalisation it loads, with every value in the right cell.
    const { snapshot } = await importXlsx(stripped, 'stripped.xlsx');
    const sheet = snapshot.sheets[snapshot.sheetOrder[0]]!;
    const matrix = sheet.cellData as unknown as Record<string, Record<string, { v?: unknown }>>;

    expect(sheet.name).toBe('Result 1');
    expect(matrix['0']['0'].v).toBe('지점명');
    expect(matrix['0']['1'].v).toBe('요일');
    expect(matrix['0']['2'].v).toBe(3);
    expect(matrix['1']['0'].v).toBe('서울지점');
    expect(matrix['1']['1'].v).toBe('월');
    expect(matrix['1']['2'].v).toBe(42);
  });

  it('returns the original bytes when nothing needs repair', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('S').getCell('A1').value = 1;
    const bytes = new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer);
    expect(normalizeXlsx(bytes)).toBe(bytes);
  });

  it('passes non-zip input straight through for ExcelJS to reject', () => {
    const garbage = new Uint8Array([1, 2, 3, 4]);
    expect(normalizeXlsx(garbage)).toBe(garbage);
  });
});
