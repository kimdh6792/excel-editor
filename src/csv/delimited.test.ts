import { describe, expect, it } from 'vitest';
import { parseDelimited, serialiseDelimited, sniffDelimiter } from './delimited';

describe('parseDelimited', () => {
  it('splits plain rows', () => {
    expect(parseDelimited('a,b\nc,d', ',')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ]);
  });

  it('keeps empty fields', () => {
    expect(parseDelimited('a,,b', ',')).toEqual([['a', '', 'b']]);
    expect(parseDelimited(',,', ',')).toEqual([['', '', '']]);
  });

  it('unwraps quoted fields and doubled quotes', () => {
    expect(parseDelimited('"a,b",c', ',')).toEqual([['a,b', 'c']]);
    expect(parseDelimited('"he said ""hi""",x', ',')).toEqual([['he said "hi"', 'x']]);
  });

  it('keeps newlines inside quoted fields', () => {
    expect(parseDelimited('"line1\nline2",b', ',')).toEqual([['line1\nline2', 'b']]);
  });

  it('treats CRLF as one break', () => {
    expect(parseDelimited('a,b\r\nc,d', ',')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ]);
  });

  it('accepts a lone CR as a break', () => {
    expect(parseDelimited('a\rb', ',')).toEqual([['a'], ['b']]);
  });

  it('does not emit a row for a trailing newline', () => {
    expect(parseDelimited('a,b\n', ',')).toEqual([['a', 'b']]);
    expect(parseDelimited('a,b\r\n', ',')).toEqual([['a', 'b']]);
  });

  it('treats a mid-field quote as literal', () => {
    // Exporters emit this by accident; discarding the row would be worse.
    expect(parseDelimited('5" pipe,x', ',')).toEqual([['5" pipe', 'x']]);
  });

  it('handles tabs as the delimiter', () => {
    expect(parseDelimited('a\tb\nc\td', '\t')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ]);
  });

  it('returns nothing for empty input', () => {
    expect(parseDelimited('', ',')).toEqual([]);
  });
});

describe('sniffDelimiter', () => {
  it('finds commas', () => {
    expect(sniffDelimiter('a,b,c\n1,2,3')).toBe(',');
  });

  it('finds tabs, as DB exports use', () => {
    expect(sniffDelimiter('name\tday\n서울\t월')).toBe('\t');
  });

  it('finds semicolons', () => {
    expect(sniffDelimiter('a;b;c\n1;2;3')).toBe(';');
  });

  it('prefers the delimiter that splits every row evenly', () => {
    // A tab-separated export whose values contain commas: splitting on comma
    // gives 1, 2, 1 fields, while tab gives a clean 3 every row.
    const text = 'name\tday\tcount\n서울, 강남\t월\t3\n부산\t화\t5';
    expect(sniffDelimiter(text)).toBe('\t');
  });

  it('still picks comma when the tab is the one inside a value', () => {
    const text = 'name,note\n서울,a\tb\n부산,c\td';
    expect(sniffDelimiter(text)).toBe(',');
  });

  it('ignores delimiters that only appear inside quotes', () => {
    expect(sniffDelimiter('"a;b"\t"c;d"\n"e;f"\t"g;h"')).toBe('\t');
  });

  it('falls back to comma for single-column input', () => {
    expect(sniffDelimiter('one\ntwo\nthree')).toBe(',');
    expect(sniffDelimiter('')).toBe(',');
  });
});

describe('serialiseDelimited', () => {
  it('writes CRLF line endings', () => {
    expect(serialiseDelimited([['a', 'b'], ['c', 'd']], ',')).toBe('a,b\r\nc,d');
  });

  it('quotes only what needs it', () => {
    expect(serialiseDelimited([['plain', 'a,b', 'say "hi"', 'two\nlines', ' pad ']], ',')).toBe(
      'plain,"a,b","say ""hi""","two\nlines"," pad "',
    );
  });

  it('leaves empty fields unquoted', () => {
    expect(serialiseDelimited([['', '']], ',')).toBe(',');
  });

  it('round-trips every awkward field', () => {
    const rows = [
      ['a,b', 'say "hi"', 'two\nlines', ' pad ', '', 'plain'],
      ['tab\there', 'semi;colon', '"leading quote', 'trailing\\'],
    ];
    expect(parseDelimited(serialiseDelimited(rows, ','), ',')).toEqual(rows);
  });
});
