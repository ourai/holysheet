import EventEmitter from '@ntks/event-emitter';
import XSpreadsheet from '@wotaware/x-spreadsheet';

import { SheetCell, SheetRow, SheetRange, ISheet, Sheet } from '../sheet';

import { SpreadsheetOptions, ResolvedOptions, Spreadsheet } from './typing';
import { resolveOptions, createXSpreadsheetInstance } from './helper';

class Holysheet extends EventEmitter implements Spreadsheet {
  private readonly options: ResolvedOptions;

  private xs: XSpreadsheet = null as any;

  private sheets: ISheet[] = [];

  private sheet: ISheet = null as any;

  private chosenCell: SheetCell | null = null;
  private chosenRange: SheetRange | null = null;

  private chosenRows: SheetRow[] = [];
  private chosenCols: any[] = [];

  private rowChosen: boolean = false;
  private colChosen: boolean = false;

  private clearRowAndColStatus(): void {
    this.rowChosen = false;
    this.colChosen = false;

    this.chosenRows = [];
    this.chosenCols = [];
  }

  private handleRangeChoose(range: SheetRange): void {
    this.chosenCell = null;
    this.chosenRange = range;

    this.sheet.setSelection({ cell: null, range });
    this.clearRowAndColStatus();

    const [sci, sri, eci, eri] = range;

    if (sci === 0 && eci === this.sheet.getColumnCount() - 1) {
      this.rowChosen = true;
      this.chosenRows = this.sheet.getRowsInRange() as SheetRow[];
    }

    if (sri === 0 && eri === this.sheet.getRowCount() - 1) {
      this.colChosen = true;
    }

    this.emit('range-change', range);
  }

  private handleCellChoose(colIndex: number, rowIndex: number): void {
    // `rowIndex` 为 `-1` 是列标题区域，
    // `colIndex` 为 `-1` 是行标题区域
    if (rowIndex === -1 || colIndex === -1) {
      let range: SheetRange;

      if (rowIndex === -1) {
        range = [colIndex, 0, colIndex, this.sheet.getRowCount() - 1];
      } else {
        range = [0, rowIndex, this.sheet.getColumnCount() - 1, rowIndex];
      }

      this.handleRangeChoose(range);
    } else {
      const cell = this.sheet.getCell(colIndex, rowIndex) as SheetCell;

      this.chosenCell = cell;
      this.chosenRange = null;

      const [colSpan = 0, rowSpan = 0] = cell.span || [];

      this.sheet.setSelection({
        cell,
        range: [colIndex, rowIndex, colIndex + colSpan, rowIndex + rowSpan],
      });
      this.clearRowAndColStatus();

      this.emit('cell-change', this.chosenCell);
    }
  }

  constructor({ el, ...others }: SpreadsheetOptions = {}) {
    super();

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

    const sheet = new Sheet({
      name: 'sheet',
      columnCount: this.options.column.count!,
      rowCount: this.options.row.count!,
    });

    const xs = createXSpreadsheetInstance(elementOrSelector, this.options);

    xs.on('cell-selected', (_, rowIndex, colIndex) => this.handleCellChoose(colIndex, rowIndex));

    xs.on('cells-selected', (_, { sci, sri, eci, eri }) =>
      this.handleRangeChoose([sci, sri, eci, eri]),
    );

    xs.on('width-resized', (ci, width) => {
      sheet.setColumnWidth(ci, width);
      this.emit('width-change', { index: ci, width });
    });

    xs.on('height-resized', (ri, height) => {
      sheet.setRowHeight(ri, height);
      this.emit('height-change', { index: ri, height });
    });

    this.sheet = sheet;
    this.xs = xs;
  }
}

export default Holysheet;
