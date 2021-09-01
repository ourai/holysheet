import XSpreadsheet from '@wotaware/x-spreadsheet';

import { Spreadsheet, SpreadsheetOptions } from './typing';
import { createXSpreadsheetInstance } from './helper';

class Holysheet implements Spreadsheet {
  private readonly options: Omit<SpreadsheetOptions, 'el'>;

  private xs: XSpreadsheet = null as any;

  constructor({ el, ...others }: SpreadsheetOptions) {
    this.options = others;

    if (el) {
      this.xs = createXSpreadsheetInstance(el, others);
    }
  }

  public mount(elementOrSelector: HTMLElement | string): void {
    if (this.xs) {
      return;
    }

    this.xs = createXSpreadsheetInstance(elementOrSelector, this.options);
  }
}

export default Holysheet;
