/**
 * Repair pass applied to xlsx bytes before ExcelJS parses them.
 *
 * In SpreadsheetML the `r` attribute on `<row>` and `<c>` is optional — when it
 * is absent the position is implied by document order. Excel and Numbers honour
 * that; ExcelJS does not, and throws `Invalid row number in model` on the whole
 * file. Several export libraries (and plenty of government data portals) emit
 * exactly this minimal form, so the files most worth opening are the ones that
 * fail. Rewriting the worksheet XML with explicit positions makes them readable.
 *
 * The result is only ever fed to the parser; nothing normalised here is written
 * back to the user's file.
 */
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { columnToLetter } from './units';

/**
 * Matches a `<row>` or `<c>` start tag with its attributes.
 *
 * Attribute values must be quoted in XML, so allowing `>` inside the quotes is
 * what keeps this from stopping at the wrong character. The element name is
 * followed by an attribute, `/` or `>`, which is why `<cols>` and `<rowBreaks>`
 * do not match.
 */
const TAG = /<(row|c)((?:\s+[\w:.-]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>/g;

const HAS_R = /\sr\s*=\s*(?:"([^"]*)"|'([^']*)')/;

function readR(attrs: string): string | undefined {
  const match = HAS_R.exec(attrs);
  if (!match) return undefined;
  return match[1] ?? match[2];
}

/** `A1` / `$AB$12` -> 1-based column number. */
function columnOf(ref: string): number | undefined {
  const letters = /^\$?([A-Za-z]+)/.exec(ref);
  if (!letters) return undefined;
  return letters[1]
    .toUpperCase()
    .split('')
    .reduce((acc, ch) => acc * 26 + (ch.charCodeAt(0) - 64), 0);
}

/**
 * Adds missing `r` attributes to rows and cells. Returns undefined when the
 * sheet already states every position, so untouched files stay byte-identical.
 */
export function normalizeSheetXml(xml: string): string | undefined {
  let changed = false;
  let row = 0;
  let col = 0;

  const out = xml.replace(TAG, (whole, name: string, attrs: string, selfClose: string) => {
    if (name === 'row') {
      const stated = readR(attrs);
      if (stated !== undefined) {
        const parsed = Number.parseInt(stated, 10);
        row = Number.isFinite(parsed) ? parsed : row + 1;
        col = 0;
        return whole;
      }
      row += 1;
      col = 0;
      changed = true;
      return `<row r="${row}"${attrs}${selfClose}>`;
    }

    // name === 'c'
    const stated = readR(attrs);
    if (stated !== undefined) {
      col = columnOf(stated) ?? col + 1;
      return whole;
    }
    col += 1;
    changed = true;
    return `<c r="${columnToLetter(col - 1)}${row}"${attrs}${selfClose}>`;
  });

  return changed ? out : undefined;
}

/**
 * Normalises every worksheet part in the package. Returns the original bytes
 * untouched when no repair was needed.
 */
export function normalizeXlsx(bytes: Uint8Array): Uint8Array {
  let entries: Record<string, Uint8Array>;
  try {
    entries = unzipSync(bytes);
  } catch {
    // Not a zip we can read — let ExcelJS produce the error message.
    return bytes;
  }

  let changed = false;
  for (const [path, data] of Object.entries(entries)) {
    if (!/^xl\/worksheets\/sheet[^/]*\.xml$/.test(path)) continue;
    const repaired = normalizeSheetXml(strFromU8(data));
    if (repaired !== undefined) {
      entries[path] = strToU8(repaired);
      changed = true;
    }
  }

  if (!changed) return bytes;
  return zipSync(entries);
}
