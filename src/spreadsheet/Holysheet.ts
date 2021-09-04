import XSpreadsheet from '@wotaware/x-spreadsheet';

import { SheetCell, SheetRange, ISheet, Sheet } from '../sheet';

import { Spreadsheet, SpreadsheetOptions, ResolvedOptions } from './typing';
import { resolveOptions, createXSpreadsheetInstance } from './helper';

class Holysheet implements Spreadsheet {
  private readonly options: ResolvedOptions;

  private xs: XSpreadsheet = null as any;

  private sheets: ISheet[] = [];

  private sheet: ISheet = null as any;

  constructor({ el, ...others }: SpreadsheetOptions = {}) {
    this.options = resolveOptions(this, others);

    if (el) {
      this.xs = createXSpreadsheetInstance(el, this.options);
    }
  }

  public select(
    colIndex: number,
    rowIndex: number,
    endColIndex?: number,
    endRowIndex?: number,
  ): void {
    let cell: SheetCell | null;
    let range: SheetRange;

    if (endColIndex !== undefined && endRowIndex !== undefined) {
      cell = null;
      range = [colIndex, rowIndex, endColIndex, endRowIndex];
    } else {
      cell = this.sheet.getCell(colIndex, rowIndex) || null;
      range = [colIndex, rowIndex, colIndex, rowIndex];
    }

    this.sheet.setSelection({ cell, range });

    if (cell) {
      (this.xs as any).sheet.selector.set(colIndex, rowIndex);
    }
  }

  public mount(elementOrSelector: HTMLElement | string): void {
    if (this.xs) {
      return;
    }

    this.xs = createXSpreadsheetInstance(elementOrSelector, this.options);

    this.sheet = new Sheet({
      name: 'sheet',
      columnCount: this.options.column.count!,
      rowCount: this.options.row.count!,
    });
  }
}

export default Holysheet;
