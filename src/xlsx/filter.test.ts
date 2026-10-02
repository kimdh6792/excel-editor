import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { exportXlsx } from './export';
import { FILTER_RESOURCE_NAME, buildFilterResource, readFilterRefs } from './filter';
import { importXlsx } from './import';

describe('filter resource', () => {
  it('builds the resource Univer expects', () => {
    const resource = buildFilterResource(new Map([['sheet-1', 'A1:D10']]));
    expect(resource).toHaveLength(1);
    expect(resource[0].name).toBe(FILTER_RESOURCE_NAME);
    expect(JSON.parse(resource[0].data)).toEqual({
      'sheet-1': { ref: { startRow: 0, startColumn: 0, endRow: 9, endColumn: 3 } },
    });
  });

  it('produces nothing when no sheet has a filter', () => {
    expect(buildFilterResource(new Map())).toEqual([]);
  });

  it('skips refs it cannot parse', () => {
    expect(buildFilterResource(new Map([['sheet-1', 'not-a-range']]))).toEqual([]);
  });

  it('reads refs back out', () => {
    const resource = buildFilterResource(new Map([
      ['sheet-1', 'A1:D10'],
      ['sheet-2', 'B2:C5'],
    ]));
    expect(readFilterRefs(resource)).toEqual(
      new Map([
        ['sheet-1', 'A1:D10'],
        ['sheet-2', 'B2:C5'],
      ]),
    );
  });

  it('survives absent, foreign and corrupt resources', () => {
    expect(readFilterRefs(undefined)).toEqual(new Map());
    expect(readFilterRefs([])).toEqual(new Map());
    expect(readFilterRefs([{ name: 'SHEET_OTHER_PLUGIN', data: '{}' }])).toEqual(new Map());
    expect(readFilterRefs([{ name: FILTER_RESOURCE_NAME, data: 'not json' }])).toEqual(new Map());
    expect(readFilterRefs([{ name: FILTER_RESOURCE_NAME, data: '{"s":{}}' }])).toEqual(new Map());
    expect(readFilterRefs([{ name: FILTER_RESOURCE_NAME, data: 'null' }])).toEqual(new Map());
  });
});

describe('autofilter round trip', () => {
  it('carries the filter range through open and save', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('데이터');
    ws.getCell('A1').value = '이름';
    ws.getCell('B1').value = '값';
    ws.getCell('A2').value = '가';
    ws.getCell('B2').value = 1;
    ws.autoFilter = 'A1:B2';
    const bytes = new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer);

    const { snapshot } = await importXlsx(bytes, 'filter.xlsx');
    // The filter must reach the snapshot as a Univer resource.
    expect(readFilterRefs(snapshot.resources)).toEqual(new Map([['sheet-1', 'A1:B2']]));

    const saved = await exportXlsx(snapshot);
    const out = new ExcelJS.Workbook();
    await out.xlsx.load(saved.buffer.slice(0) as ArrayBuffer);
    expect(out.getWorksheet('데이터')!.autoFilter).toBe('A1:B2');
  });

  it('leaves files without a filter untouched', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('S').getCell('A1').value = 1;
    const bytes = new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer);

    const { snapshot } = await importXlsx(bytes, 'plain.xlsx');
    expect(snapshot.resources).toEqual([]);

    const out = new ExcelJS.Workbook();
    await out.xlsx.load((await exportXlsx(snapshot)).buffer.slice(0) as ArrayBuffer);
    expect(out.getWorksheet('S')!.autoFilter).toBeFalsy();
  });
});
