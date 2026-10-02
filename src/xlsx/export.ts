/**
 * Univer workbook snapshot -> xlsx bytes.
 *
 * ExcelJS rewrites the whole package, so only what this module writes survives.
 * `inspect.ts` detects the parts that will be lost and the UI warns before the
 * first save of such a file.
 */
import ExcelJS from 'exceljs';
import type { ICellData, IRange, IStyleData, IWorkbookData, Nullable } from '@univerjs/core';
import { readFilterRefs } from './filter';
import { resolveStyle, toExcelStyle } from './style';
import type { PreservedCell, PreservedColumn, PreservedRow } from './import';
import {
  DEFAULT_ROW_HEIGHT_POINTS,
  charsToPixels,
  columnToLetter,
  pixelsToChars,
  pixelsToPoints,
  pointsToPixels,
} from './units';

type CellMatrix = Record<string, Record<string, ICellData>>;

/**
 * Returns the row height in points, favouring the value the file was opened
 * with. Pixel<->point conversion rounds, so re-deriving an untouched height
 * would nudge it on every save.
 */
function rowHeightPoints(px: number, custom: PreservedRow | undefined): number {
  if (custom && pointsToPixels(custom.points) === px) return custom.points;
  return pixelsToPoints(px);
}

/** Column width in character units, favouring the originally loaded value. */
function columnWidthChars(px: number, custom: PreservedColumn | undefined): number {
  if (custom && charsToPixels(custom.chars) === px) return custom.chars;
  return pixelsToChars(px);
}

function hexToArgb(hex: string | undefined): string | undefined {
  if (!hex) return undefined;
  const h = hex.replace('#', '');
  return /^[0-9a-f]{6}$/i.test(h) ? `FF${h.toUpperCase()}` : undefined;
}

/**
 * Chooses the value to write. When the cell still holds exactly the text we
 * flattened on import, the original rich-text runs or hyperlink are restored so
 * untouched cells round-trip losslessly.
 */
function excelValue(cell: ICellData): ExcelJS.CellValue {
  if (cell.f) {
    const formula = cell.f.startsWith('=') ? cell.f.slice(1) : cell.f;
    const result = cell.v as string | number | boolean | undefined;
    return { formula, result: result ?? undefined } as ExcelJS.CellValue;
  }

  const preserved = cell.custom as PreservedCell | undefined;
  const text = cell.v == null ? '' : String(cell.v);

  if (preserved?.richText && preserved.richText.text === text) {
    return { richText: preserved.richText.runs } as ExcelJS.CellValue;
  }
  if (preserved?.hyperlink && preserved.hyperlink.text === text) {
    return { hyperlink: preserved.hyperlink.target, text } as ExcelJS.CellValue;
  }

  if (cell.v == null) return null;
  return cell.v as ExcelJS.CellValue;
}

function writeSheet(
  wb: ExcelJS.Workbook,
  sheet: NonNullable<IWorkbookData['sheets'][string]>,
  styles: Record<string, Nullable<IStyleData>>,
  forceVisible: boolean,
  autoFilterRef: string | undefined,
): void {
  const state = forceVisible || !sheet.hidden ? 'visible' : sheet.hidden === 2 ? 'veryHidden' : 'hidden';

  const freeze = sheet.freeze;
  const frozen = freeze && (freeze.xSplit > 0 || freeze.ySplit > 0);
  const view: Partial<ExcelJS.WorksheetView> = frozen
    ? {
        state: 'frozen',
        xSplit: freeze.xSplit,
        ySplit: freeze.ySplit,
        topLeftCell: `${columnToLetter(freeze.xSplit)}${freeze.ySplit + 1}`,
      }
    : { state: 'normal' };
  view.showGridLines = sheet.showGridlines !== 0;
  if (sheet.rightToLeft) view.rightToLeft = true;

  const tabArgb = hexToArgb(sheet.tabColor);
  // ExcelJS renders a bare `ref` string straight into `<autoFilter ref="…"/>`.
  const autoFilter = autoFilterRef ?? null;
  const ws = wb.addWorksheet(sheet.name ?? 'Sheet', {
    state,
    views: [view as ExcelJS.WorksheetView],
    properties: {
      tabColor: tabArgb ? { argb: tabArgb } : undefined,
      defaultRowHeight: sheet.defaultRowHeight ? pixelsToPoints(sheet.defaultRowHeight) : undefined,
      defaultColWidth: sheet.defaultColumnWidth ? pixelsToChars(sheet.defaultColumnWidth) : undefined,
    } as ExcelJS.WorksheetProperties,
  });
  ws.autoFilter = autoFilter as ExcelJS.Worksheet['autoFilter'];

  /* ------------------------------------------- column metadata before cells */

  const columnStyles = new Map<number, IStyleData>();
  for (const [key, col] of Object.entries(sheet.columnData ?? {})) {
    const index = Number(key);
    if (!Number.isFinite(index)) continue;
    const target = ws.getColumn(index + 1);
    if (typeof col?.w === 'number') {
      target.width = columnWidthChars(col.w, col.custom as PreservedColumn | undefined);
    }
    if (col?.hd) target.hidden = true;
    const resolved = resolveStyle(col?.s, styles);
    if (resolved) columnStyles.set(index, resolved);
  }

  /* ------------------------------------------------------------------ cells */

  const matrix = (sheet.cellData ?? {}) as unknown as CellMatrix;
  for (const [rowKey, row] of Object.entries(matrix)) {
    const r = Number(rowKey);
    if (!Number.isFinite(r) || !row) continue;

    for (const [colKey, cell] of Object.entries(row)) {
      const c = Number(colKey);
      if (!Number.isFinite(c) || !cell) continue;

      const target = ws.getCell(r + 1, c + 1);
      const value = excelValue(cell);
      if (value !== null) target.value = value;

      // Fall back to the column's style so a value-only cell does not lose the
      // column fill: Excel applies a column style only to cells it has no
      // explicit format for, and every cell we write gets one.
      const resolved = resolveStyle(cell.s, styles) ?? columnStyles.get(c);
      const style = toExcelStyle(resolved);
      if (style) target.style = style as ExcelJS.Style;
    }
  }

  /* ----------------------------------------------------- row metadata, cols */

  const defaultHeightPoints = sheet.defaultRowHeight
    ? pixelsToPoints(sheet.defaultRowHeight)
    : DEFAULT_ROW_HEIGHT_POINTS;

  for (const [key, row] of Object.entries(sheet.rowData ?? {})) {
    const index = Number(key);
    if (!Number.isFinite(index)) continue;
    const target = ws.getRow(index + 1);
    if (typeof row?.h === 'number') {
      target.height = rowHeightPoints(row.h, row.custom as PreservedRow | undefined);
    }
    if (row?.hd) {
      target.hidden = true;
      // ExcelJS serialises a row to nothing unless it has a height or at least
      // one cell, which would silently drop a hidden-but-empty row. Pinning the
      // default height makes the row real without changing how it looks.
      if (target.height === undefined && !target.hasValues) {
        target.height = defaultHeightPoints;
      }
    }
  }

  // Assigning `column.style` directly does not propagate to existing cells in
  // ExcelJS, so this is safe to do after the cells are written.
  for (const [index, style] of columnStyles) {
    const excelStyle = toExcelStyle(style);
    if (excelStyle) ws.getColumn(index + 1).style = excelStyle as ExcelJS.Style;
  }

  /* ----------------------------------------------------------------- merges */

  // Merging copies the master cell's style over the covered range, so it has to
  // happen once the cells themselves are final.
  for (const range of (sheet.mergeData ?? []) as IRange[]) {
    try {
      ws.mergeCells(
        range.startRow + 1,
        range.startColumn + 1,
        range.endRow + 1,
        range.endColumn + 1,
      );
    } catch {
      // Overlapping or degenerate ranges are skipped rather than failing the save.
    }
  }
}

/** Serialises a Univer snapshot to xlsx bytes. */
export async function exportXlsx(snapshot: IWorkbookData): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  wb.created = new Date();
  wb.modified = new Date();
  // We write cached formula results, but asking Excel to recalculate on open
  // guards against anything Univer's engine evaluated differently.
  wb.calcProperties.fullCalcOnLoad = true;

  const styles = snapshot.styles ?? {};
  const order = snapshot.sheetOrder?.length
    ? snapshot.sheetOrder
    : Object.keys(snapshot.sheets ?? {});

  const anyVisible = order.some((id) => !snapshot.sheets?.[id]?.hidden);
  const filterRefs = readFilterRefs(snapshot.resources);

  order.forEach((id, index) => {
    const sheet = snapshot.sheets?.[id];
    if (!sheet) return;
    // Excel rejects a workbook with every sheet hidden.
    writeSheet(wb, sheet, styles, !anyVisible && index === 0, filterRefs.get(id));
  });

  const buffer = await wb.xlsx.writeBuffer();
  return new Uint8Array(buffer as ArrayBuffer);
}
