/**
 * Print settings, kept in their own module so the toolbar can read the defaults
 * without pulling in the renderer. Everything else under `src/print/` only
 * loads once the user actually prints.
 */

/** What a print run covers. */
export type PrintArea = 'sheet' | 'selection' | 'all';

export interface PrintOptions {
  /** Draw a light rule around every cell, like Excel's "print gridlines". */
  gridlines: boolean;
  /** How many rows from the top of the area repeat on every page. */
  repeatHeaderRows: number;
  /** Scale columns so the sheet's full width lands on one page. */
  fitToWidth: boolean;
}

export interface PrintSettings extends PrintOptions {
  area: PrintArea;
}

/** Upper bound on repeated header rows; past this the repeat eats the page. */
export const MAX_REPEAT_HEADER_ROWS = 10;

export const DEFAULT_PRINT_SETTINGS: PrintSettings = {
  area: 'sheet',
  gridlines: true,
  repeatHeaderRows: 1,
  fitToWidth: true,
};
