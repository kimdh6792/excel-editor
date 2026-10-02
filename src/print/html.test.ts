/**
 * The markup the printer actually sees.
 *
 * Two kinds of mistake matter here. One is layout — a header row that does not
 * repeat, a column that loses its width — and the other is injection: every
 * colour, font name and cell value came out of a file someone else wrote, and
 * all three end up inside an attribute.
 */
import {
  BaselineOffset,
  BooleanNumber,
  BorderStyleTypes,
  HorizontalAlign,
  TextDecoration,
  VerticalAlign,
  WrapStrategy,
} from '@univerjs/core';
import type { IStyleData } from '@univerjs/core';
import { describe, expect, it } from 'vitest';
import type { PrintCell, PrintSheet } from './area';
import { buildPrintHtml } from './html';
import { DEFAULT_PRINT_SETTINGS } from './options';
import type { PrintOptions } from './options';

const OPTIONS: PrintOptions = { ...DEFAULT_PRINT_SETTINGS, repeatHeaderRows: 0 };

function cell(text: string, style?: IStyleData, extra: Partial<PrintCell> = {}): PrintCell {
  return { text, style, numeric: false, rowSpan: 1, colSpan: 1, ...extra };
}

function sheet(rows: PrintCell[][], overrides: Partial<PrintSheet> = {}): PrintSheet {
  return {
    name: 'Sheet1',
    rows,
    rowHeights: rows.map(() => 20),
    columnWidths: (rows[0] ?? []).map(() => 100),
    showGridlines: true,
    ...overrides,
  };
}

/** Renders one cell and returns its `<td …>` tag, attributes included. */
function td(value: PrintCell, options: PrintOptions = OPTIONS): string {
  const html = buildPrintHtml([sheet([[value]])], options);
  return html.match(/<td[^>]*>/)?.[0] ?? '';
}

describe('buildPrintHtml', () => {
  it('escapes text that would otherwise be markup', () => {
    const html = buildPrintHtml([sheet([[cell('<b>a</b> & "b"')]])], OPTIONS);
    expect(html).toContain('&lt;b&gt;a&lt;/b&gt; &amp; &quot;b&quot;');
    expect(html).not.toContain('<b>');
  });

  it('keeps line breaks inside a cell as text, for `white-space: pre` to render', () => {
    const html = buildPrintHtml([sheet([[cell('first\nsecond')]])], OPTIONS);
    expect(html).toContain('first\nsecond');
  });

  /* ------------------------------------------------------------- injection */

  it('drops a colour that is not a plain hex value', () => {
    const attacked = td(cell('x', { cl: { rgb: 'red" onmouseover="alert(1)' } }));
    // No style attribute at all, so there is nothing to break out of.
    expect(attacked).toBe('<td>');
    expect(attacked).not.toContain('onmouseover');
  });

  it('accepts the hex forms Univer produces', () => {
    expect(td(cell('x', { cl: { rgb: '#ff0000' } }))).toContain('color:#ff0000');
    expect(td(cell('x', { bg: { rgb: '#FFF' } }))).toContain('background-color:#FFF');
    expect(td(cell('x', { bg: { rgb: '#11223344' } }))).toContain('background-color:#11223344');
    expect(td(cell('x', { cl: { rgb: 'rgb(255,0,0)' } }))).not.toContain('color:');
  });

  it('strips punctuation from a font name so it cannot end the declaration', () => {
    const html = td(cell('x', { ff: "Arial'; background: url(evil)" }));
    expect(html).toContain("font-family:'Arial background urlevil'");
    expect(html).not.toContain('background:');
  });

  /* ----------------------------------------------------------------- style */

  it('maps the text styles', () => {
    const html = td(
      cell('x', {
        fs: 14,
        bl: BooleanNumber.TRUE,
        it: BooleanNumber.TRUE,
        ul: { s: BooleanNumber.TRUE, t: TextDecoration.DOUBLE },
        st: { s: BooleanNumber.TRUE },
      }),
    );
    expect(html).toContain('font-size:14pt');
    expect(html).toContain('font-weight:700');
    expect(html).toContain('font-style:italic');
    expect(html).toContain('text-decoration:underline line-through double');
  });

  it('right-aligns numbers unless the cell says otherwise', () => {
    expect(td(cell('1,234', undefined, { numeric: true }))).toContain('text-align:right');
    expect(td(cell('abc'))).not.toContain('text-align');
    expect(
      td(cell('1,234', { ht: HorizontalAlign.LEFT }, { numeric: true })),
    ).toContain('text-align:left');
  });

  it('passes vertical alignment and wrapping through', () => {
    expect(td(cell('x', { vt: VerticalAlign.MIDDLE }))).toContain('vertical-align:middle');
    expect(td(cell('x', { tb: WrapStrategy.WRAP }))).toContain('white-space:pre-wrap');
    expect(td(cell('x', { tb: WrapStrategy.CLIP }))).not.toContain('white-space');
  });

  it('draws borders per side, with the dash-dot family falling back to dashed', () => {
    const black = { rgb: '#000000' };
    const html = td(
      cell('x', {
        bd: {
          t: { s: BorderStyleTypes.THICK, cl: { rgb: '#123456' } },
          b: { s: BorderStyleTypes.DOUBLE, cl: black },
          l: { s: BorderStyleTypes.MEDIUM_DASH_DOT, cl: black },
          r: { s: BorderStyleTypes.NONE, cl: black },
        },
      }),
    );
    expect(html).toContain('border-top:3px solid #123456');
    expect(html).toContain('border-bottom:3px double #000000');
    expect(html).toContain('border-left:2px dashed #000000');
    expect(html).not.toContain('border-right');
  });

  it('falls back to black for a border with no usable colour', () => {
    // Univer's own type demands `cl`, but files written elsewhere reach this
    // code with it missing or carrying a theme reference.
    const bd = { t: { s: BorderStyleTypes.THIN } } as NonNullable<IStyleData['bd']>;
    expect(td(cell('x', { bd }))).toContain('border-top:1px solid #000000');
  });

  it('wraps super- and subscript text in an element, not a declaration', () => {
    expect(buildPrintHtml([sheet([[cell('2', { va: BaselineOffset.SUPERSCRIPT })]])], OPTIONS))
      .toContain('<sup>2</sup>');
    expect(buildPrintHtml([sheet([[cell('2', { va: BaselineOffset.SUBSCRIPT })]])], OPTIONS))
      .toContain('<sub>2</sub>');
  });

  it('converts an indent to left padding', () => {
    expect(td(cell('x', { pd: { l: 16, t: 0, b: 0, r: 0 } }))).toContain('padding-left:16px');
  });

  it('leaves an unstyled cell without a style attribute', () => {
    expect(td(cell('plain'))).toBe('<td>');
  });

  /* ------------------------------------------------------------------ spans */

  it('emits colspan and rowspan only when they exceed one', () => {
    expect(td(cell('x', undefined, { colSpan: 3, rowSpan: 2 }))).toBe(
      '<td colspan="3" rowspan="2">',
    );
    expect(td(cell('x'))).toBe('<td>');
  });

  /* ----------------------------------------------------------------- widths */

  it('fits to width by giving each column its share as a percentage', () => {
    const html = buildPrintHtml(
      [sheet([[cell('a'), cell('b')]], { columnWidths: [300, 100] })],
      { ...OPTIONS, fitToWidth: true },
    );
    expect(html).toContain('style="width:100%"');
    expect(html).toContain('<col style="width:75.0000%">');
    expect(html).toContain('<col style="width:25.0000%">');
  });

  it('uses the sheet’s own pixel widths at actual size', () => {
    const html = buildPrintHtml(
      [sheet([[cell('a'), cell('b')]], { columnWidths: [300, 100] })],
      { ...OPTIONS, fitToWidth: false },
    );
    expect(html).toContain('style="width:auto"');
    expect(html).toContain('<col style="width:300px">');
    expect(html).toContain('<col style="width:100px">');
  });

  it('spreads the width evenly when every column is zero-width', () => {
    const html = buildPrintHtml([sheet([[cell('a'), cell('b')]], { columnWidths: [0, 0] })], {
      ...OPTIONS,
      fitToWidth: true,
    });
    expect(html.match(/width:50\.0000%/g)).toHaveLength(2);
  });

  /* ------------------------------------------------------------ header rows */

  it('puts the repeated rows in a thead and the rest in the tbody', () => {
    const rows = [[cell('h1')], [cell('h2')], [cell('a')], [cell('b')]];
    const html = buildPrintHtml([sheet(rows)], { ...OPTIONS, repeatHeaderRows: 2 });
    const thead = html.match(/<thead>(.*?)<\/thead>/s)?.[1] ?? '';
    const tbody = html.match(/<tbody>(.*?)<\/tbody>/s)?.[1] ?? '';
    expect(thead).toContain('h1');
    expect(thead).toContain('h2');
    expect(thead).not.toContain('>a<');
    expect(tbody).toContain('>a<');
    expect(tbody).toContain('>b<');
  });

  it('keeps heights aligned with rows after the header split', () => {
    const rows = [[cell('h')], [cell('a')]];
    const html = buildPrintHtml(
      [sheet(rows, { rowHeights: [31, 77] })],
      { ...OPTIONS, repeatHeaderRows: 1 },
    );
    expect(html).toContain('<thead><tr style="height:31px">');
    expect(html).toContain('<tbody><tr style="height:77px">');
  });

  it('omits the thead entirely when nothing repeats', () => {
    const html = buildPrintHtml([sheet([[cell('a')]])], { ...OPTIONS, repeatHeaderRows: 0 });
    expect(html).not.toContain('<thead>');
  });

  it('never repeats every row, which would leave nothing to paginate', () => {
    const html = buildPrintHtml([sheet([[cell('a')], [cell('b')]])], {
      ...OPTIONS,
      repeatHeaderRows: 5,
    });
    const thead = html.match(/<thead>(.*?)<\/thead>/s)?.[1] ?? '';
    expect(thead).toContain('>a<');
    expect(thead).not.toContain('>b<');
  });

  it('ignores a negative or fractional repeat count', () => {
    expect(
      buildPrintHtml([sheet([[cell('a')], [cell('b')]])], { ...OPTIONS, repeatHeaderRows: -3 }),
    ).not.toContain('<thead>');
    const fractional = buildPrintHtml([sheet([[cell('a')], [cell('b')]])], {
      ...OPTIONS,
      repeatHeaderRows: 1.9,
    });
    expect(fractional.match(/<thead>(.*?)<\/thead>/s)?.[1]).toContain('>a<');
  });

  /* -------------------------------------------------------------- gridlines */

  it('adds the gridline class only when the option and the sheet agree', () => {
    expect(buildPrintHtml([sheet([[cell('a')]])], { ...OPTIONS, gridlines: true })).toContain(
      'class="print-table gridlines"',
    );
    expect(buildPrintHtml([sheet([[cell('a')]])], { ...OPTIONS, gridlines: false })).toContain(
      'class="print-table"',
    );
    // A sheet with gridlines turned off in Excel prints without them either way.
    expect(
      buildPrintHtml([sheet([[cell('a')]], { showGridlines: false })], {
        ...OPTIONS,
        gridlines: true,
      }),
    ).toContain('class="print-table"');
  });

  /* ----------------------------------------------------------- sheet titles */

  it('titles each sheet only when more than one prints', () => {
    const one = buildPrintHtml([sheet([[cell('a')]])], OPTIONS);
    expect(one).not.toContain('<h2>');

    const many = buildPrintHtml(
      [sheet([[cell('a')]], { name: '1월' }), sheet([[cell('b')]], { name: '2월' })],
      OPTIONS,
    );
    expect(many).toContain('<h2>1월</h2>');
    expect(many).toContain('<h2>2월</h2>');
    expect(many.match(/<section class="print-sheet">/g)).toHaveLength(2);
  });

  it('escapes a sheet name too', () => {
    const html = buildPrintHtml(
      [sheet([[cell('a')]], { name: '<x>' }), sheet([[cell('b')]])],
      OPTIONS,
    );
    expect(html).toContain('<h2>&lt;x&gt;</h2>');
  });

  it('renders nothing for no sheets', () => {
    expect(buildPrintHtml([], OPTIONS)).toBe('');
  });
});
