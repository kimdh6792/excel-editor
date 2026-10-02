/**
 * Delimited-text parsing and serialising (RFC 4180 and the usual deviations).
 *
 * Handles quoted fields, doubled quotes inside them, embedded newlines, and
 * either line ending. The delimiter is sniffed rather than assumed, because
 * "CSV" from a Korean locale is often semicolon-separated and DB exports are
 * often tab-separated.
 */

export const DELIMITERS = [',', '\t', ';', '|'] as const;
export type Delimiter = (typeof DELIMITERS)[number];

/** Parses delimited text into a grid of raw strings. */
export function parseDelimited(text: string, delimiter: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQuotes = false;
  let i = 0;
  // Tracks whether the current row has seen anything at all, so that a trailing
  // newline does not produce a spurious final row.
  let rowStarted = false;

  const endField = () => {
    row.push(field);
    field = '';
    rowStarted = true;
  };
  const endRow = () => {
    endField();
    rows.push(row);
    row = [];
    rowStarted = false;
  };

  while (i < text.length) {
    const ch = text[i];

    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        inQuotes = false;
        i += 1;
        continue;
      }
      field += ch;
      i += 1;
      continue;
    }

    if (ch === '"' && field === '') {
      // A quote only opens a quoted field at the start of one; mid-field quotes
      // are literal, which is what spreadsheet exporters produce by accident.
      inQuotes = true;
      i += 1;
      continue;
    }

    if (ch === delimiter) {
      endField();
      i += 1;
      continue;
    }

    if (ch === '\r' || ch === '\n') {
      endRow();
      // Consume CRLF as one break.
      i += ch === '\r' && text[i + 1] === '\n' ? 2 : 1;
      continue;
    }

    field += ch;
    i += 1;
  }

  if (field !== '' || rowStarted || row.length > 0) endRow();
  return rows;
}

/**
 * Picks the delimiter that yields the most consistent table.
 *
 * Scored on the first handful of lines by how many rows agree on the field count
 * and how many fields that is. A file with one column scores zero everywhere, in
 * which case comma is returned and the result is a single column either way.
 */
export function sniffDelimiter(text: string): Delimiter {
  const sample = text.split(/\r\n|\r|\n/).slice(0, 20).join('\n');
  if (!sample) return ',';

  let best: Delimiter = ',';
  let bestScore = -1;

  for (const delimiter of DELIMITERS) {
    const rows = parseDelimited(sample, delimiter).filter((r) => r.length > 0);
    if (rows.length === 0) continue;

    const counts = rows.map((r) => r.length);
    const first = counts[0];
    if (first < 2) continue;

    const agreeing = counts.filter((c) => c === first).length;
    // Consistency dominates; field count only breaks ties, so a stray character
    // that happens to appear often cannot outrank the real delimiter.
    const score = (agreeing / counts.length) * 1000 + Math.min(first, 100);
    if (score > bestScore) {
      bestScore = score;
      best = delimiter;
    }
  }

  return best;
}

/** True when a field has to be quoted to survive a round trip. */
function needsQuoting(field: string, delimiter: string): boolean {
  if (field === '') return false;
  if (field.includes(delimiter) || field.includes('"')) return true;
  if (field.includes('\n') || field.includes('\r')) return true;
  // Leading or trailing spaces are silently eaten by some readers.
  return field !== field.trim();
}

function serialiseField(field: string, delimiter: string): string {
  if (!needsQuoting(field, delimiter)) return field;
  return `"${field.replace(/"/g, '""')}"`;
}

/** Serialises a grid with CRLF line endings, the convention Excel expects. */
export function serialiseDelimited(rows: string[][], delimiter: string): string {
  return rows.map((row) => row.map((f) => serialiseField(f, delimiter)).join(delimiter)).join('\r\n');
}
