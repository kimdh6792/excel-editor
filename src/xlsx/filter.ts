/**
 * Autofilter mapping.
 *
 * Univer's filter plugin keeps its state in the snapshot's `resources` array
 * under `SHEET_FILTER_PLUGIN`, as a JSON string of `{ [sheetId]: IAutoFilter }`.
 * xlsx stores the equivalent as `<autoFilter ref="A1:D10"/>` on the worksheet.
 *
 * Only the *range* round-trips. ExcelJS models the autofilter as a bare ref and
 * drops `<filterColumn>` criteria entirely, so which values are currently
 * filtered out cannot survive a save — the filter buttons come back, the
 * selection behind them does not.
 */
import type { IRange, IResources } from '@univerjs/core';
import { formatRangeRef, parseRangeRef } from './range';

/** Resource name the Univer filter plugin registers under. */
export const FILTER_RESOURCE_NAME = 'SHEET_FILTER_PLUGIN';

/** The subset of Univer's `IAutoFilter` this app populates. */
interface AutoFilterData {
  ref: IRange;
  [key: string]: unknown;
}

/** Builds the filter resource entry from per-sheet autofilter refs. */
export function buildFilterResource(refsBySheetId: Map<string, string>): IResources {
  if (refsBySheetId.size === 0) return [];

  const data: Record<string, AutoFilterData> = {};
  for (const [sheetId, ref] of refsBySheetId) {
    const range = parseRangeRef(ref);
    if (range) data[sheetId] = { ref: range };
  }

  if (Object.keys(data).length === 0) return [];
  return [{ name: FILTER_RESOURCE_NAME, data: JSON.stringify(data) }];
}

/**
 * Reads back the autofilter range of each sheet as an A1 ref, ready to assign to
 * `worksheet.autoFilter`. Malformed resource data is skipped rather than thrown:
 * a corrupt filter must not block saving the actual cells.
 */
export function readFilterRefs(resources: IResources | undefined): Map<string, string> {
  const out = new Map<string, string>();
  const entry = resources?.find((r) => r.name === FILTER_RESOURCE_NAME);
  if (!entry?.data) return out;

  let parsed: unknown;
  try {
    parsed = JSON.parse(entry.data);
  } catch {
    return out;
  }
  if (!parsed || typeof parsed !== 'object') return out;

  for (const [sheetId, value] of Object.entries(parsed as Record<string, unknown>)) {
    const ref = (value as AutoFilterData | undefined)?.ref;
    if (
      ref &&
      typeof ref.startRow === 'number' &&
      typeof ref.startColumn === 'number' &&
      typeof ref.endRow === 'number' &&
      typeof ref.endColumn === 'number'
    ) {
      out.set(sheetId, formatRangeRef(ref));
    }
  }
  return out;
}
