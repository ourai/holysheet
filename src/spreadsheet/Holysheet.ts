import EventEmitter from '@ntks/event-emitter';
import XSpreadsheet from '@wotaware/x-spreadsheet';

import {
  CellId,
  TableCell,
  TableRow,
  TableRange,
  ITable,
  SheetData,
  ISheet,
  Sheet,
} from '../sheet';

import { MountEl, SpreadsheetOptions, ResolvedOptions, Spreadsheet } from './typing';
import { resolveOptions, createXSpreadsheetInstance } from './helper';

class Holysheet extends EventEmitter implements Spreadsheet {
  private readonly options: ResolvedOptions;

  private xs: XSpreadsheet = null as any;

  private sheets: ISheet[] = [];

  private sheet: ISheet = null as any;
  private table: ITable = null as any;

  private chosenCell: TableCell | null = null;
  private chosenRange: TableRange | null = null;

  private chosenRows: TableRow[] = [];
  private chosenCols: any[] = [];

  private rowChosen: boolean = false;
  private colChosen: boolean = false;

  private setCurrentSheet(index: number = 0): void {
    const prevSheet = this.sheet;
    const prevTable = this.table;

    if (prevTable) {
      prevTable.clearSelection();
      prevTable.off();
    }

    this.sheet = this.sheets[index];

    if (!this.sheet.hasTable()) {
      const table = this.sheet.createTable({
        columnCount: this.options.column.count!,
        rowCount: this.options.row.count!,
      });

      table.on('cell-update', cell => this.emit('cell-update', cell));
    }

    this.table = this.sheet.getTable()!;
  }

  private clearRowAndColStatus(): void {
    this.rowChosen = false;
    this.colChosen = false;

    this.chosenRows = [];
    this.chosenCols = [];
  }

  private handleRangeChoose(range: TableRange): void {
    this.chosenCell = null;
    this.chosenRange = range;

    this.table.setSelection({ cell: null, range });
    this.clearRowAndColStatus();

    const [sci, sri, eci, eri] = range;

    if (sci === 0 && eci === this.table.getColumnCount() - 1) {
      this.rowChosen = true;
      this.chosenRows = this.table.getRowsInRange() as TableRow[];
    }

    if (sri === 0 && eri === this.table.getRowCount() - 1) {
      this.colChosen = true;
    }

    this.emit('range-change', range);
  }

  private handleCellChoose(colIndex: number, rowIndex: number): void {
    // `rowIndex` 为 `-1` 是列标题区域，
    // `colIndex` 为 `-1` 是行标题区域
    if (rowIndex === -1 || colIndex === -1) {
      let range: TableRange;

      if (rowIndex === -1) {
        range = [colIndex, 0, colIndex, this.table.getRowCount() - 1];
      } else {
        range = [0, rowIndex, this.table.getColumnCount() - 1, rowIndex];
      }

      this.handleRangeChoose(range);
    } else {
      const cell = this.table.getCell(colIndex, rowIndex) as TableCell;

      this.chosenCell = cell;
      this.chosenRange = null;

      const [colSpan = 0, rowSpan = 0] = cell.span || [];

      this.table.setSelection({
        cell,
        range: [colIndex, rowIndex, colIndex + colSpan, rowIndex + rowSpan],
      });
      this.clearRowAndColStatus();

      this.emit('cell-change', this.chosenCell);
    }
  }

  private createXSpreadsheetInstance(elementOrSelector: MountEl): void {
    const xs = createXSpreadsheetInstance(elementOrSelector, this.options);

    xs.on('cell-selected', (_, rowIndex, colIndex) => this.handleCellChoose(colIndex, rowIndex));

    xs.on('cells-selected', (_, { sci, sri, eci, eri }) =>
      this.handleRangeChoose([sci, sri, eci, eri]),
    );

    xs.on('width-resized', (ci, width) => {
      this.table.setColumnWidth(ci, width);
      this.emit('width-change', { index: ci, width });
    });

    xs.on('height-resized', (ri, height) => {
      this.table.setRowHeight(ri, height);
      this.emit('height-change', { index: ri, height });
    });

    this.xs = xs;
  }

  constructor({ el, ...others }: SpreadsheetOptions = {}) {
    super();

    this.options = resolveOptions(this, others);

    if (el) {
      this.createXSpreadsheetInstance(el);
    }
  }

  public mount(elementOrSelector: HTMLElement | string): void {
    if (!this.xs) {
      this.createXSpreadsheetInstance(elementOrSelector);
    }
  }

  public select(
    colIndex: number,
    rowIndex: number,
    endColIndex?: number,
    endRowIndex?: number,
  ): void {
    let cell: TableCell | null;
    let range: TableRange;

    if (endColIndex !== undefined && endRowIndex !== undefined) {
      cell = null;
      range = [colIndex, rowIndex, endColIndex, endRowIndex];
    } else {
      cell = this.table.getCell(colIndex, rowIndex) || null;
      range = [colIndex, rowIndex, colIndex, rowIndex];
    }

    this.table.setSelection({ cell, range });

    if (cell) {
      (this.xs as any).sheet.selector.set(colIndex, rowIndex);
    }
  }

  public getSelectedRows(): TableRow[] {
    return this.chosenRows;
  }

  public setSheets(sheets: SheetData[]): void {
    const sheetMap: Record<string, ISheet> = this.sheets.reduce(
      (prev, sheet) => ({ ...prev, [sheet.getId()]: sheet }),
      {},
    );

    this.sheets = sheets.map(sheet =>
      sheet.id && sheetMap[sheet.id] ? sheetMap[sheet.id] : new Sheet(sheet),
    );

    this.setCurrentSheet();
    this.emit(
      'change',
      this.sheets.map(sheet => ({ id: sheet.getId(), name: sheet.getName() })),
    );
  }

  public changeSheet(index: number): void {
    this.setCurrentSheet(index);
    this.emit('sheet-change', index);
  }

  public updateCell(id: CellId, data: Record<string, any>): void {
    this.table.setCellProperties(id, data);
  }
}

export default Holysheet;
