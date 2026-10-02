/**
 * Text decoding for delimited files.
 *
 * Korean CSVs in the wild are usually CP949 (what Excel on Windows writes) or
 * UTF-8, and nothing in the file says which. Guessing wrong turns every Korean
 * column into mojibake, so detection matters more here than it would elsewhere.
 */

export type DetectedEncoding = 'utf-8' | 'utf-8-bom' | 'utf-16le' | 'utf-16be' | 'cp949';

export interface DecodedText {
  text: string;
  encoding: DetectedEncoding;
}

function startsWith(bytes: Uint8Array, prefix: number[]): boolean {
  if (bytes.length < prefix.length) return false;
  return prefix.every((b, i) => bytes[i] === b);
}

/**
 * Decodes bytes to text, detecting the encoding.
 *
 * A BOM is taken at its word. Otherwise the bytes are decoded as strict UTF-8;
 * CP949 is the fallback, because any byte sequence is valid CP949 so trying it
 * first would never fail and would silently mangle real UTF-8.
 */
export function decodeText(bytes: Uint8Array): DecodedText {
  if (startsWith(bytes, [0xef, 0xbb, 0xbf])) {
    return {
      text: new TextDecoder('utf-8').decode(bytes.subarray(3)),
      encoding: 'utf-8-bom',
    };
  }
  if (startsWith(bytes, [0xff, 0xfe])) {
    return {
      text: new TextDecoder('utf-16le').decode(bytes.subarray(2)),
      encoding: 'utf-16le',
    };
  }
  if (startsWith(bytes, [0xfe, 0xff])) {
    return {
      text: new TextDecoder('utf-16be').decode(bytes.subarray(2)),
      encoding: 'utf-16be',
    };
  }

  try {
    return {
      text: new TextDecoder('utf-8', { fatal: true }).decode(bytes),
      encoding: 'utf-8',
    };
  } catch {
    // The `euc-kr` label maps to the windows-949 index, which is the superset
    // Excel actually writes — so this covers CP949 too.
    return {
      text: new TextDecoder('euc-kr').decode(bytes),
      encoding: 'cp949',
    };
  }
}

/**
 * Encodes text as UTF-8 with a BOM.
 *
 * The BOM is deliberate: without it Excel reads a UTF-8 CSV using the system
 * code page and Korean text arrives broken. Every other tool ignores it.
 */
export function encodeTextWithBom(text: string): Uint8Array {
  const body = new TextEncoder().encode(text);
  const out = new Uint8Array(body.length + 3);
  out[0] = 0xef;
  out[1] = 0xbb;
  out[2] = 0xbf;
  out.set(body, 3);
  return out;
}
