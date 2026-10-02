/**
 * Snapshot helpers that do not need a parser.
 *
 * Kept apart from `import.ts` so that showing an empty grid at startup does not
 * pull ExcelJS into the initial bundle — it is only needed once the user
 * actually opens or saves a file.
 */
import { LocaleType } from '@univerjs/core';
import type { IWorkbookData } from '@univerjs/core';
import {
  DEFAULT_COLUMN_WIDTH_CHARS,
  DEFAULT_ROW_HEIGHT_POINTS,
  charsToPixels,
  pointsToPixels,
} from './units';

/** Grid size a blank workbook presents, so it still looks like a spreadsheet. */
export const MIN_ROWS = 100;
export const MIN_COLS = 26;

/** An empty one-sheet workbook, used for "new file". */
export function createEmptySnapshot(name = '제목 없음'): IWorkbookData {
  return {
    id: `wb-new-${name}`,
    name,
    appVersion: '0.1.0',
    locale: LocaleType.KO_KR,
    styles: {},
    sheetOrder: ['sheet-1'],
    sheets: {
      'sheet-1': {
        id: 'sheet-1',
        name: 'Sheet1',
        tabColor: '',
        hidden: 0,
        rowCount: MIN_ROWS,
        columnCount: MIN_COLS,
        zoomRatio: 1,
        freeze: { xSplit: 0, ySplit: 0, startRow: -1, startColumn: -1 },
        scrollTop: 0,
        scrollLeft: 0,
        defaultColumnWidth: charsToPixels(DEFAULT_COLUMN_WIDTH_CHARS),
        defaultRowHeight: pointsToPixels(DEFAULT_ROW_HEIGHT_POINTS),
        mergeData: [],
        cellData: {},
        rowData: {},
        columnData: {},
        rowHeader: { width: 46 },
        columnHeader: { height: 20 },
        showGridlines: 1,
        rightToLeft: 0,
      },
    },
  };
}
