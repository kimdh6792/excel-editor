import ExcelJS from 'exceljs';
import { strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { describeLossyFeatures, detectLossyFeatures } from './inspect';

/** Builds a plain workbook, then injects extra package parts into its zip. */
async function workbookWithParts(parts: string[]): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('S').getCell('A1').value = 1;
  const entries = unzipSync(new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer));
  for (const path of parts) entries[path] = strToU8('<x/>');
  return zipSync(entries);
}

describe('detectLossyFeatures', () => {
  it('finds nothing in a plain workbook', async () => {
    const bytes = await workbookWithParts([]);
    expect(detectLossyFeatures(bytes)).toEqual([]);
  });

  it('reports charts, images, pivot tables and macros', async () => {
    const bytes = await workbookWithParts([
      'xl/charts/chart1.xml',
      'xl/charts/chart2.xml',
      'xl/media/image1.png',
      'xl/pivotTables/pivotTable1.xml',
      'xl/vbaProject.bin',
    ]);

    const found = detectLossyFeatures(bytes);
    const byLabel = Object.fromEntries(found.map((f) => [f.label, f.count]));

    expect(byLabel['차트']).toBe(2);
    expect(byLabel['이미지·도형']).toBe(1);
    expect(byLabel['피벗 테이블']).toBe(1);
    expect(byLabel['매크로(VBA)']).toBe(1);
  });

  it('reports comments under either spelling used by Excel', async () => {
    const bytes = await workbookWithParts([
      'xl/comments1.xml',
      'xl/threadedComments/threadedComment1.xml',
    ]);
    expect(detectLossyFeatures(bytes).find((f) => f.label === '주석(메모)')?.count).toBe(2);
  });

  it('survives input that is not a zip', () => {
    expect(detectLossyFeatures(new Uint8Array([0, 1, 2]))).toEqual([]);
  });
});

describe('describeLossyFeatures', () => {
  it('stays silent when nothing is lost', () => {
    expect(describeLossyFeatures([])).toBeUndefined();
  });

  it('counts only when there is more than one', () => {
    const text = describeLossyFeatures([
      { label: '차트', count: 3 },
      { label: '매크로(VBA)', count: 1 },
    ])!;
    expect(text).toContain('차트 3개');
    expect(text).toContain('매크로(VBA)');
    expect(text).not.toContain('매크로(VBA) 1개');
  });
});
