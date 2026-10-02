import { describe, expect, it } from 'vitest';
import { excelColorToHex, hexToExcelColor } from './color';

describe('excelColorToHex', () => {
  it('strips the alpha channel from argb', () => {
    expect(excelColorToHex({ argb: 'FFFF0000' })).toBe('#ff0000');
    expect(excelColorToHex({ argb: 'FF123ABC' })).toBe('#123abc');
  });

  it('accepts 6-digit values, which some writers emit', () => {
    expect(excelColorToHex({ argb: '00FF00' })).toBe('#00ff00');
  });

  it('resolves theme indices against the default Office theme', () => {
    expect(excelColorToHex({ theme: 0 })).toBe('#ffffff'); // lt1
    expect(excelColorToHex({ theme: 1 })).toBe('#000000'); // dk1
    expect(excelColorToHex({ theme: 4 })).toBe('#4472c4'); // accent1
    expect(excelColorToHex({ theme: 9 })).toBe('#70ad47'); // accent6
  });

  it('lightens with a positive tint and darkens with a negative one', () => {
    const base = excelColorToHex({ theme: 4 })!;
    const lighter = excelColorToHex({ theme: 4, tint: 0.6 })!;
    const darker = excelColorToHex({ theme: 4, tint: -0.5 })!;

    const luminance = (hex: string) => {
      const h = hex.slice(1);
      return (
        Number.parseInt(h.slice(0, 2), 16) +
        Number.parseInt(h.slice(2, 4), 16) +
        Number.parseInt(h.slice(4, 6), 16)
      );
    };

    expect(luminance(lighter)).toBeGreaterThan(luminance(base));
    expect(luminance(darker)).toBeLessThan(luminance(base));
  });

  it("reproduces Excel's published accent1 tint swatches", () => {
    const channels = (hex: string) =>
      [1, 3, 5].map((i) => Number.parseInt(hex.slice(i, i + 2), 16));

    // The tint values are the ones Excel writes for its theme-color dropdown;
    // the expected results are the swatches it renders for them.
    const cases: Array<[string, number, number[]]> = [
      ['Lighter 80%', 0.7999, [0xd9, 0xe2, 0xf3]],
      ['Lighter 60%', 0.5999, [0xb4, 0xc7, 0xe7]],
      ['Lighter 40%', 0.3999, [0x8f, 0xaa, 0xdc]],
      ['Darker 25%', -0.2499, [0x2f, 0x55, 0x97]],
      ['Darker 50%', -0.4999, [0x20, 0x39, 0x64]],
    ];

    for (const [label, tint, expected] of cases) {
      const actual = channels(excelColorToHex({ theme: 4, tint })!);
      // Within one 8-bit step of Excel's own rounding.
      expect(actual.map((c, i) => Math.abs(c - expected[i]) <= 1), label).toEqual([
        true,
        true,
        true,
      ]);
    }
  });

  it('resolves the legacy indexed palette', () => {
    expect(excelColorToHex({ indexed: 0 })).toBe('#000000');
    expect(excelColorToHex({ indexed: 2 })).toBe('#ff0000');
    expect(excelColorToHex({ indexed: 22 })).toBe('#c0c0c0');
    expect(excelColorToHex({ indexed: 64 })).toBe('#000000');
  });

  it('returns undefined for absent or unrecognised colors', () => {
    expect(excelColorToHex(undefined)).toBeUndefined();
    expect(excelColorToHex(null)).toBeUndefined();
    expect(excelColorToHex({})).toBeUndefined();
    expect(excelColorToHex({ argb: 'nonsense' })).toBeUndefined();
    expect(excelColorToHex({ theme: 99 })).toBeUndefined();
    expect(excelColorToHex({ indexed: 999 })).toBeUndefined();
  });
});

describe('hexToExcelColor', () => {
  it('adds an opaque alpha channel', () => {
    expect(hexToExcelColor('#ff0000')).toEqual({ argb: 'FFFF0000' });
    expect(hexToExcelColor('123abc')).toEqual({ argb: 'FF123ABC' });
  });

  it('expands 3-digit shorthand', () => {
    expect(hexToExcelColor('#f00')).toEqual({ argb: 'FFFF0000' });
  });

  it('rejects garbage', () => {
    expect(hexToExcelColor(undefined)).toBeUndefined();
    expect(hexToExcelColor('')).toBeUndefined();
    expect(hexToExcelColor('rgb(1,2,3)')).toBeUndefined();
  });

  it('round-trips through excelColorToHex', () => {
    for (const hex of ['#ff0000', '#00ff00', '#123abc', '#ffffff', '#000000']) {
      expect(excelColorToHex(hexToExcelColor(hex))).toBe(hex);
    }
  });
});
