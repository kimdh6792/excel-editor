import { describe, expect, it } from 'vitest';
import { decodeText, encodeTextWithBom } from './encoding';

/**
 * Encodes text as CP949 by inverting the platform's own `euc-kr` decoder.
 *
 * Hand-written byte constants are too easy to get wrong by one — and a wrong
 * constant still decodes to *some* Korean syllable, so the test would fail while
 * looking like a decoder bug. Deriving the bytes keeps the fixture honest.
 */
const cp949Table = new Map<string, [number, number]>();
function cp949(text: string): Uint8Array {
  if (cp949Table.size === 0) {
    const decoder = new TextDecoder('euc-kr');
    for (let lead = 0x81; lead <= 0xfd; lead++) {
      for (let trail = 0x41; trail <= 0xfe; trail++) {
        const char = decoder.decode(new Uint8Array([lead, trail]));
        if (char.length === 1 && char !== '�' && !cp949Table.has(char)) {
          cp949Table.set(char, [lead, trail]);
        }
      }
    }
  }

  const out: number[] = [];
  for (const char of text) {
    const code = char.codePointAt(0)!;
    if (code < 0x80) {
      out.push(code);
      continue;
    }
    const pair = cp949Table.get(char);
    if (!pair) throw new Error(`CP949로 표현할 수 없는 문자: ${char}`);
    out.push(pair[0], pair[1]);
  }
  return new Uint8Array(out);
}

describe('cp949 test helper', () => {
  it('produces bytes the decoder reads back as the original text', () => {
    for (const text of ['서울', '부산', '대구', '가나다라마'])
      expect(new TextDecoder('euc-kr').decode(cp949(text))).toBe(text);
  });

  it('leaves ASCII as single bytes', () => {
    expect([...cp949('a,1')]).toEqual([0x61, 0x2c, 0x31]);
  });
});

describe('decodeText', () => {
  it('decodes plain UTF-8', () => {
    const bytes = new TextEncoder().encode('이름,요일\n서울,월');
    const { text, encoding } = decodeText(bytes);
    expect(encoding).toBe('utf-8');
    expect(text).toBe('이름,요일\n서울,월');
  });

  it('strips a UTF-8 BOM', () => {
    const bytes = encodeTextWithBom('이름,요일');
    const { text, encoding } = decodeText(bytes);
    expect(encoding).toBe('utf-8-bom');
    expect(text).toBe('이름,요일');
    expect(text.charCodeAt(0)).not.toBe(0xfeff);
  });

  it('decodes UTF-16LE via its BOM', () => {
    const body = '서울';
    const bytes = new Uint8Array(2 + body.length * 2);
    bytes[0] = 0xff;
    bytes[1] = 0xfe;
    for (let i = 0; i < body.length; i++) {
      const code = body.charCodeAt(i);
      bytes[2 + i * 2] = code & 0xff;
      bytes[3 + i * 2] = code >> 8;
    }
    const { text, encoding } = decodeText(bytes);
    expect(encoding).toBe('utf-16le');
    expect(text).toBe('서울');
  });

  it('falls back to CP949 for bytes that are not valid UTF-8', () => {
    const { text, encoding } = decodeText(cp949('서울'));
    expect(encoding).toBe('cp949');
    expect(text).toBe('서울');
  });

  it('decodes a whole CP949 row, ASCII and Korean mixed', () => {
    // The shape a CSV exported by Excel on Windows actually has.
    const { text, encoding } = decodeText(cp949('서울,강남지점,02-123-4567,3'));
    expect(encoding).toBe('cp949');
    expect(text).toBe('서울,강남지점,02-123-4567,3');
  });

  it('prefers UTF-8 when the bytes are valid UTF-8', () => {
    // CP949 would also "decode" these bytes, just into garbage — so order matters.
    const bytes = new TextEncoder().encode('해운대');
    expect(decodeText(bytes)).toEqual({ text: '해운대', encoding: 'utf-8' });
  });

  it('handles empty input', () => {
    expect(decodeText(new Uint8Array())).toEqual({ text: '', encoding: 'utf-8' });
  });
});

describe('encodeTextWithBom', () => {
  it('prefixes the UTF-8 BOM so Excel reads Korean correctly', () => {
    const bytes = encodeTextWithBom('서울');
    expect([bytes[0], bytes[1], bytes[2]]).toEqual([0xef, 0xbb, 0xbf]);
    expect(new TextDecoder('utf-8').decode(bytes.subarray(3))).toBe('서울');
  });

  it('round-trips through decodeText', () => {
    const text = '이름,요일\r\n서울,월';
    expect(decodeText(encodeTextWithBom(text)).text).toBe(text);
  });
});
