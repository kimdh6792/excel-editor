import { describe, expect, it } from 'vitest';
import type { ICellData, IWorkbookData } from '@univerjs/core';
import { Workbook } from '@univerjs/core';
import type { ILogService } from '@univerjs/core';
import {
  csvToSnapshot,
  formatDateSerial,
  inferValue,
  isDateFormat,
  sheetsLostInCsv,
  snapshotToCsv,
  snapshotToCsvText,
} from './convert';
import { decodeText, encodeTextWithBom } from './encoding';
import { dateToSerial } from '../xlsx/units';

const silentLog: ILogService = {
  debug: () => {},
  log: () => {},
  warn: () => {},
  error: () => {},
  deprecate: () => {},
  setLogLevel: () => {},
};

function cells(snapshot: IWorkbookData) {
  const sheet = snapshot.sheets[snapshot.sheetOrder[0]]!;
  return sheet.cellData as unknown as Record<string, Record<string, ICellData>>;
}

function fromText(text: string, name = 'data.csv') {
  return csvToSnapshot(encodeTextWithBom(text), name);
}

/* ------------------------------------------------------------- inference */

describe('inferValue', () => {
  it('converts plain numbers', () => {
    expect(inferValue('42')).toBe(42);
    expect(inferValue('-7')).toBe(-7);
    expect(inferValue('3.14')).toBe(3.14);
    expect(inferValue('0')).toBe(0);
    expect(inferValue('0.5')).toBe(0.5);
    expect(inferValue('1.50')).toBe(1.5);
    expect(inferValue('1e3')).toBe(1000);
  });

  it('keeps leading-zero values as text', () => {
    // Phone numbers, postal codes and account numbers all die here otherwise.
    expect(inferValue('01012345678')).toBe('01012345678');
    expect(inferValue('06236')).toBe('06236');
    expect(inferValue('007')).toBe('007');
  });

  it('keeps values too long for exact float representation as text', () => {
    expect(inferValue('12345678901234567')).toBe('12345678901234567');
    expect(inferValue('9999999999999999')).toBe('9999999999999999');
    // 15 significant digits still fits.
    expect(inferValue('123456789012345')).toBe(123456789012345);
  });

  it('keeps anything non-numeric as text', () => {
    expect(inferValue('02-123-4567')).toBe('02-123-4567');
    expect(inferValue('1,234')).toBe('1,234');
    expect(inferValue('서울')).toBe('서울');
    expect(inferValue('')).toBe('');
    expect(inferValue(' 42')).toBe(' 42');
    expect(inferValue('42 ')).toBe('42 ');
    expect(inferValue('+42')).toBe('+42');
    expect(inferValue('.5')).toBe('.5');
    expect(inferValue('1.2.3')).toBe('1.2.3');
    expect(inferValue('Infinity')).toBe('Infinity');
    expect(inferValue('NaN')).toBe('NaN');
  });

  it('does not guess at dates or booleans', () => {
    // Guessing here is exactly what corrupts data in other tools.
    expect(inferValue('2026-03-15')).toBe('2026-03-15');
    expect(inferValue('TRUE')).toBe('TRUE');
    expect(inferValue('3/15/2026')).toBe('3/15/2026');
  });
});

/* ----------------------------------------------------------------- import */

describe('csvToSnapshot', () => {
  it('places values at the right coordinates', () => {
    const { snapshot } = fromText('이름,수량\n서울,3\n부산,5');
    const grid = cells(snapshot);
    expect(grid['0']['0'].v).toBe('이름');
    expect(grid['0']['1'].v).toBe('수량');
    expect(grid['1']['0'].v).toBe('서울');
    expect(grid['1']['1'].v).toBe(3);
    expect(grid['2']['1'].v).toBe(5);
  });

  it('reports what it detected', () => {
    const result = fromText('a\tb\n1\t2', 'export.tsv');
    expect(result.delimiter).toBe('\t');
    expect(result.encoding).toBe('utf-8-bom');
    expect(result.rowCount).toBe(2);
    expect(result.columnCount).toBe(2);
  });

  it('names the sheet after the file, without the extension', () => {
    const { snapshot } = fromText('a,b', '주소록.csv');
    expect(snapshot.sheets[snapshot.sheetOrder[0]]!.name).toBe('주소록');
  });

  it('skips empty fields rather than storing blank cells', () => {
    const grid = cells(fromText('a,,c').snapshot);
    expect(grid['0']['0'].v).toBe('a');
    expect(grid['0']['1']).toBeUndefined();
    expect(grid['0']['2'].v).toBe('c');
  });

  it('produces a snapshot Univer accepts', () => {
    const { snapshot } = fromText('이름,수량\n서울,3');
    const workbook = new Workbook(snapshot, silentLog);
    const sheet = workbook.getSheets()[0];
    expect(sheet.getCell(0, 0)?.v).toBe('이름');
    expect(sheet.getCell(1, 1)?.v).toBe(3);
  });

  it('handles an empty file', () => {
    const { snapshot, rowCount } = fromText('');
    expect(rowCount).toBe(0);
    expect(cells(snapshot)).toEqual({});
  });
});

/* ----------------------------------------------------------------- export */

describe('isDateFormat', () => {
  it('recognises date and time patterns', () => {
    expect(isDateFormat('yyyy-mm-dd')).toBe(true);
    expect(isDateFormat('m/d/yy')).toBe(true);
    expect(isDateFormat('h:mm:ss')).toBe(true);
    expect(isDateFormat('yyyy"년" m"월"')).toBe(true);
  });

  it('does not mistake number formats for dates', () => {
    expect(isDateFormat('#,##0.00')).toBe(false);
    expect(isDateFormat('0.00%')).toBe(false);
    expect(isDateFormat('General')).toBe(false);
    expect(isDateFormat(undefined)).toBe(false);
    // Letters hiding inside literals and bracketed sections must not count.
    expect(isDateFormat('0.00"년"')).toBe(false);
    expect(isDateFormat('[$-409]#,##0')).toBe(false);
  });
});

describe('formatDateSerial', () => {
  it('writes a bare date when the format has no time', () => {
    const serial = dateToSerial(new Date(Date.UTC(2026, 2, 15)));
    expect(formatDateSerial(serial, 'yyyy-mm-dd')).toBe('2026-03-15');
  });

  it('includes the time when the format asks for it', () => {
    const serial = dateToSerial(new Date(Date.UTC(2026, 2, 15, 13, 45, 30)));
    expect(formatDateSerial(serial, 'yyyy-mm-dd h:mm:ss')).toBe('2026-03-15 13:45:30');
  });

  it('includes the time when the value has one even if the format hides it', () => {
    const serial = dateToSerial(new Date(Date.UTC(2026, 2, 15, 9, 0, 0)));
    expect(formatDateSerial(serial, 'yyyy-mm-dd')).toBe('2026-03-15 09:00:00');
  });

  it('does not drift at midnight', () => {
    for (const day of [1, 15, 28, 29, 30, 31]) {
      const serial = dateToSerial(new Date(Date.UTC(2026, 0, day)));
      expect(formatDateSerial(serial, 'yyyy-mm-dd')).toBe(`2026-01-${String(day).padStart(2, '0')}`);
    }
  });
});

describe('snapshotToCsvText', () => {
  it('round-trips a grid of values', () => {
    const { snapshot } = fromText('이름,수량\n서울,3\n부산,5');
    expect(snapshotToCsvText(snapshot)).toBe('이름,수량\r\n서울,3\r\n부산,5');
  });

  it('trims to the populated area instead of padding the whole grid', () => {
    const { snapshot } = fromText('a,b');
    // The sheet is 100x26 but only one row holds anything.
    expect(snapshotToCsvText(snapshot).split('\r\n')).toHaveLength(1);
  });

  it('keeps identifier-looking text intact through a full round trip', () => {
    const original = '이름,전화,우편\n서울,01012345678,06236';
    const { snapshot } = fromText(original);
    expect(snapshotToCsvText(snapshot)).toBe(original.replace(/\n/g, '\r\n'));
  });

  it('re-quotes fields that need it', () => {
    const original = 'a,"b,c","say ""hi"""';
    const { snapshot } = fromText(original);
    expect(snapshotToCsvText(snapshot)).toBe(original);
  });

  it('writes date-formatted numbers as dates, not serials', () => {
    const { snapshot } = fromText('x');
    const sheet = snapshot.sheets['sheet-1']!;
    snapshot.styles = { s1: { n: { pattern: 'yyyy-mm-dd' } } };
    (sheet.cellData as unknown as Record<string, Record<string, ICellData>>)['0']['0'] = {
      v: dateToSerial(new Date(Date.UTC(2026, 2, 15))),
      s: 's1',
    };
    expect(snapshotToCsvText(snapshot)).toBe('2026-03-15');
  });

  it('leaves plain numbers alone even when styled', () => {
    const { snapshot } = fromText('x');
    const sheet = snapshot.sheets['sheet-1']!;
    snapshot.styles = { s1: { n: { pattern: '#,##0.00' } } };
    (sheet.cellData as unknown as Record<string, Record<string, ICellData>>)['0']['0'] = {
      v: 1234.5,
      s: 's1',
    };
    expect(snapshotToCsvText(snapshot)).toBe('1234.5');
  });

  it('returns empty text for an empty sheet', () => {
    expect(snapshotToCsvText(fromText('').snapshot)).toBe('');
  });

  it('honours a chosen delimiter', () => {
    const { snapshot } = fromText('a,b\n1,2');
    expect(snapshotToCsvText(snapshot, { delimiter: '\t' })).toBe('a\tb\r\n1\t2');
  });
});

describe('snapshotToCsv', () => {
  it('emits UTF-8 with a BOM so Excel reads Korean correctly', () => {
    const { snapshot } = fromText('이름\n서울');
    const bytes = snapshotToCsv(snapshot);
    expect([bytes[0], bytes[1], bytes[2]]).toEqual([0xef, 0xbb, 0xbf]);
    expect(decodeText(bytes).text).toBe('이름\r\n서울');
  });
});

describe('sheetsLostInCsv', () => {
  it('names the sheets a CSV save would drop', () => {
    const snapshot = {
      sheetOrder: ['a', 'b', 'c'],
      sheets: { a: { name: '첫째' }, b: { name: '둘째' }, c: { name: '셋째' } },
    } as unknown as IWorkbookData;
    expect(sheetsLostInCsv(snapshot)).toEqual(['둘째', '셋째']);
    expect(sheetsLostInCsv(snapshot, 'b')).toEqual(['첫째', '셋째']);
  });

  it('reports nothing for a single-sheet workbook', () => {
    expect(sheetsLostInCsv(fromText('a').snapshot)).toEqual([]);
  });
});
