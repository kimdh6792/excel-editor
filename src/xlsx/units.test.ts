import { describe, expect, it } from 'vitest';
import {
  charsToPixels,
  columnToLetter,
  dateToSerial,
  pixelsToChars,
  pixelsToPoints,
  pointsToPixels,
} from './units';

describe('column width', () => {
  it('round-trips whole character widths exactly', () => {
    for (const chars of [1, 5, 8, 10, 20, 50]) {
      expect(pixelsToChars(charsToPixels(chars))).toBe(chars);
    }
  });

  it("matches Excel's default width of 64px", () => {
    expect(charsToPixels(8.43)).toBe(64);
    expect(pixelsToChars(64)).toBe(8.43);
  });

  it('never produces a negative width', () => {
    expect(pixelsToChars(0)).toBe(0);
    expect(pixelsToChars(3)).toBe(0);
  });
});

describe('row height', () => {
  it('converts points to 96-DPI pixels', () => {
    expect(pointsToPixels(15)).toBe(20);
    expect(pointsToPixels(12)).toBe(16);
    expect(pointsToPixels(30)).toBe(40);
  });

  it('round-trips heights that land on whole pixels', () => {
    for (const pt of [15, 12, 30, 45, 60]) {
      expect(pixelsToPoints(pointsToPixels(pt))).toBe(pt);
    }
  });

  it('drifts by less than a point when the height does not divide evenly', () => {
    // 20pt -> 27px -> 20.25pt. Export keeps the original value for untouched
    // rows precisely because of this; the fallback still has to stay close.
    expect(pixelsToPoints(pointsToPixels(20))).toBeCloseTo(20, 0);
  });
});

describe('dateToSerial', () => {
  it('maps the Unix epoch to Excel serial 25569', () => {
    expect(dateToSerial(new Date(Date.UTC(1970, 0, 1)))).toBeCloseTo(25569, 6);
  });

  it('is off by one before 1900-03-01, matching ExcelJS', () => {
    // Excel's 1900 system contains a non-existent 1900-02-29, so dates before
    // 1900-03-01 are one lower than a linear day count. Excel shows serial 1 for
    // 1900-01-01; the linear formula gives 2. ExcelJS converts the same way, so
    // the app stays internally consistent and only pre-1900 dates read one day
    // off — which no spreadsheet in practice depends on.
    expect(dateToSerial(new Date(Date.UTC(1900, 0, 1)))).toBeCloseTo(2, 6);
  });

  it('maps a modern date', () => {
    // 2026-03-15 — cross-checked against Excel's DATEVALUE.
    expect(dateToSerial(new Date(Date.UTC(2026, 2, 15)))).toBeCloseTo(46096, 6);
  });

  it('keeps the time of day in the fraction', () => {
    const noon = dateToSerial(new Date(Date.UTC(2026, 2, 15, 12, 0, 0)));
    expect(noon % 1).toBeCloseTo(0.5, 6);
  });
});

describe('columnToLetter', () => {
  it('handles single and multi-letter columns', () => {
    expect(columnToLetter(0)).toBe('A');
    expect(columnToLetter(25)).toBe('Z');
    expect(columnToLetter(26)).toBe('AA');
    expect(columnToLetter(27)).toBe('AB');
    expect(columnToLetter(51)).toBe('AZ');
    expect(columnToLetter(52)).toBe('BA');
    expect(columnToLetter(701)).toBe('ZZ');
    expect(columnToLetter(702)).toBe('AAA');
  });
});
