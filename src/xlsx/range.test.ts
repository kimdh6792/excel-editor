import { describe, expect, it } from 'vitest';
import { formatRangeRef, parseRangeRef } from './range';

describe('parseRangeRef', () => {
  it('parses a two-cell range as 0-based inclusive', () => {
    expect(parseRangeRef('A1:B2')).toEqual({
      startRow: 0,
      startColumn: 0,
      endRow: 1,
      endColumn: 1,
    });
  });

  it('treats a single cell as a one-cell range', () => {
    expect(parseRangeRef('C3')).toEqual({
      startRow: 2,
      startColumn: 2,
      endRow: 2,
      endColumn: 2,
    });
  });

  it('ignores absolute markers', () => {
    expect(parseRangeRef('$A$1:$B$2')).toEqual(parseRangeRef('A1:B2'));
  });

  it('handles multi-letter columns', () => {
    expect(parseRangeRef('AA1')?.startColumn).toBe(26);
    expect(parseRangeRef('AB1')?.startColumn).toBe(27);
    expect(parseRangeRef('ZZ1')?.startColumn).toBe(701);
    expect(parseRangeRef('AAA1')?.startColumn).toBe(702);
  });

  it('is case-insensitive', () => {
    expect(parseRangeRef('a1:b2')).toEqual(parseRangeRef('A1:B2'));
  });

  it('rejects reversed and malformed refs', () => {
    expect(parseRangeRef('B2:A1')).toBeUndefined();
    expect(parseRangeRef('Sheet1!A1')).toBeUndefined();
    expect(parseRangeRef('A')).toBeUndefined();
    expect(parseRangeRef('1')).toBeUndefined();
    expect(parseRangeRef('')).toBeUndefined();
  });
});

describe('formatRangeRef', () => {
  it('collapses a one-cell range to a single address', () => {
    expect(formatRangeRef({ startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 })).toBe('A1');
  });

  it('formats a multi-cell range', () => {
    expect(formatRangeRef({ startRow: 0, startColumn: 0, endRow: 9, endColumn: 3 })).toBe('A1:D10');
  });

  it('round-trips with parseRangeRef', () => {
    for (const ref of ['A1', 'A1:D10', 'B2:B2', 'Z100:AA200', 'AAA1:AAB2']) {
      expect(formatRangeRef(parseRangeRef(ref)!)).toBe(
        // `B2:B2` collapses, which is the same range.
        ref === 'B2:B2' ? 'B2' : ref,
      );
    }
  });
});
