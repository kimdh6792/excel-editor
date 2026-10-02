/**
 * Pulls the printable region out of the live workbook.
 *
 * Two sources are needed and neither covers the other: the snapshot carries
 * styles, widths, merges and hidden flags, while only the facade can say what a
 * cell *reads* as once its number format is applied. This module joins them and
 * hands the result to {@link ./area}.
 *
 * The workbook is typed structurally rather than as `FWorkbook` so the contract
 * is the handful of methods actually used — which also means the whole module
 * can be tested without standing up Univer.
 */
import { BooleanNumber } from '@univerjs/core';
import type { IRange, IWorkbookData } from '@univerjs/core';
import { buildPrintSheet, isBlank } from './area';
import type { Area, PrintSheet } from './area';
import type { PrintArea } from './options';

interface PrintableRange {
  getRange(): IRange;
  getDisplayValues(): string[][];
}

interface PrintableWorksheet {
  getSheetId(): string;
  getSheetName(): string;
  /** Zero-based index of the last row holding data or formatting. */
  getLastRow(): number;
  getLastColumn(): number;
  getRange(row: number, column: number, numRows: number, numColumns: number): PrintableRange;
}

export interface PrintableWorkbook {
  save(): IWorkbookData;
  getSheets(): PrintableWorksheet[];
  getActiveSheet(): PrintableWorksheet;
  getActiveRange(): PrintableRange | null;
}

/** The area a sheet's own content occupies, which is what "the sheet" means here. */
function dataArea(worksheet: PrintableWorksheet): Area {
  return {
    startRow: 0,
    startColumn: 0,
    endRow: Math.max(0, worksheet.getLastRow()),
    endColumn: Math.max(0, worksheet.getLastColumn()),
  };
}

function toArea(range: IRange): Area {
  return {
    startRow: range.startRow,
    startColumn: range.startColumn,
    endRow: range.endRow,
    endColumn: range.endColumn,
  };
}

function isSingleCell(area: Area): boolean {
  return area.startRow === area.endRow && area.startColumn === area.endColumn;
}

/**
 * Builds one print sheet, reading display text for exactly the chosen area.
 *
 * Returns undefined when the sheet is missing from the snapshot, which should
 * not happen but is not worth aborting the whole print over.
 */
function sheetFor(
  worksheet: PrintableWorksheet,
  snapshot: IWorkbookData,
  area: Area,
): PrintSheet | undefined {
  const data = snapshot.sheets[worksheet.getSheetId()];
  if (!data) return undefined;

  const rows = area.endRow - area.startRow + 1;
  const columns = area.endColumn - area.startColumn + 1;
  const values = worksheet
    .getRange(area.startRow, area.startColumn, rows, columns)
    .getDisplayValues();

  return buildPrintSheet({
    name: worksheet.getSheetName(),
    sheet: data,
    styles: snapshot.styles ?? {},
    values,
    area,
  });
}

export function collectPrintSheets(
  workbook: PrintableWorkbook,
  area: PrintArea,
): PrintSheet[] {
  const snapshot = workbook.save();

  if (area === 'selection') {
    const active = workbook.getActiveSheet();
    const selected = workbook.getActiveRange()?.getRange();
    // Univer always has a selection, so an untouched grid reports A1. Printing a
    // single cell because the user never clicked anywhere would be a surprise;
    // fall back to the whole sheet, which is what they asked for in spirit.
    const chosen = selected ? toArea(selected) : undefined;
    const target = chosen && !isSingleCell(chosen) ? chosen : dataArea(active);
    const sheet = sheetFor(active, snapshot, target);
    return sheet ? [sheet] : [];
  }

  if (area === 'all') {
    const bySheetId = new Map(workbook.getSheets().map((s) => [s.getSheetId(), s]));
    // Follow the workbook's own tab order rather than whatever order the facade
    // hands back, so a printed stack matches the tabs on screen.
    const ordered = (snapshot.sheetOrder ?? [...bySheetId.keys()])
      .map((id) => bySheetId.get(id))
      .filter((s): s is PrintableWorksheet => Boolean(s))
      .filter((s) => snapshot.sheets[s.getSheetId()]?.hidden !== BooleanNumber.TRUE);

    const sheets = ordered
      .map((s) => sheetFor(s, snapshot, dataArea(s)))
      .filter((s): s is PrintSheet => Boolean(s));

    // Blank sheets would each cost a blank page. Dropping them all would leave
    // nothing to print, so in that case keep the first and print one page.
    const withContent = sheets.filter((s) => !isBlank(s));
    return withContent.length > 0 ? withContent : sheets.slice(0, 1);
  }

  const active = workbook.getActiveSheet();
  const sheet = sheetFor(active, snapshot, dataArea(active));
  return sheet ? [sheet] : [];
}
