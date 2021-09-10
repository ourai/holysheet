import { noop, clone } from '@ntks/toolbox';
import EventEmitter from '@ntks/event-emitter';
import XSpreadsheet, { CellData as XSpreadsheetCellData } from '@wotaware/x-spreadsheet';

import {
  TableCell,
  TableRow,
  TableRange,
  Result,
  ITable,
  SheetData,
  ISheet,
  Sheet,
} from '../sheet';

import {
  MountEl,
  CellCreator,
  RowCreator,
  SpreadsheetOptions,
  ResolvedOptions,
  Spreadsheet,
} from './typing';
import { resolveOptions, resolveCellStyle, createXSpreadsheetInstance } from './helper';

class Holysheet extends EventEmitter implements Spreadsheet {
  private readonly options: ResolvedOptions;

  private readonly cellCreator: CellCreator | undefined;
  private readonly rowCreator: RowCreator | undefined;

  private readonly beforeSheetActivate: (prev: ISheet) => boolean;
  private readonly sheetActivated: (current: ISheet, prev: ISheet) => void;

  private readonly beforeSheetRender: () => boolean;
  private readonly sheetRenderer: () => void;

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

    if (this.beforeSheetActivate(prevSheet) === false) {
      return;
    }

    const prevTable = this.table;

    if (prevTable) {
      prevTable.clearSelection();
      prevTable.off();
    }

    this.sheet = this.sheets[index];

    if (!this.sheet.hasTable()) {
      const table = this.sheet.createTable({
        cellCreator: this.cellCreator,
        rowCreator: this.rowCreator,
        columnCount: this.options.column.count!,
        rowCount: this.options.row.count!,
      });

      table.on('cell-update', cell => {
        const [colIndex, rowIndex] = this.table.getCellCoordinate(cell.id) as [number, number];

        (this.xs.cellText(rowIndex, colIndex, cell.text) as any).reRender();

        this.emit('cell-update', cell);
      });
    }

    this.table = this.sheet.getTable()!;

    this.sheetActivated(this.sheet, prevSheet);
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

  private handleOperationResult(result: Result): Result {
    if (result.success) {
      this.render();
    }

    return result;
  }

  constructor({
    el,
    cellCreator,
    rowCreator,
    beforeSheetActivate,
    sheetActivated,
    beforeSheetRender,
    sheetRendered,
    ...others
  }: SpreadsheetOptions = {}) {
    super();

    this.options = resolveOptions(this, others);

    this.cellCreator = cellCreator;
    this.rowCreator = rowCreator;

    this.beforeSheetActivate = beforeSheetActivate || (() => true);
    this.sheetActivated = sheetActivated || noop;

    this.beforeSheetRender = beforeSheetRender || (() => true);
    this.sheetRenderer = sheetRendered || noop;

    if (el) {
      this.createXSpreadsheetInstance(el);
    }
  }

  public mount(elementOrSelector: HTMLElement | string): void {
    if (!this.xs) {
      this.createXSpreadsheetInstance(elementOrSelector);
    }
  }

  public render(): void {
    if (this.beforeSheetRender() === false) {
      return;
    }

    const merges: string[] = [];

    const cols: Record<string, any> = { len: this.table.getColumnCount() };
    const rows: Record<string, any> = { len: this.table.getRowCount() };

    this.table.getColumns().forEach((col, idx) => {
      if (col.width !== undefined) {
        cols[idx] = { width: col.width };
      }
    });

    this.table
      .transformRows(row => {
        const { cells, ...others } = row;

        return {
          ...others,
          cells: cells.reduce((prev, tableCell) => {
            const { id, span, mergedCoord } = tableCell as TableCell;

            const resolved: XSpreadsheetCellData = {
              text: this.table.getCellText(id),
              style: resolveCellStyle(this.table.getCellStyle(id)),
            };

            if (span) {
              resolved.merge = [span[1], span[0]];
            }

            if (mergedCoord) {
              merges.push(mergedCoord);
            }

            return { ...prev, [this.table.getCellCoordinate(id)[0]]: resolved };
          }, {}),
        };
      })
      .forEach((row, idx) => (rows[idx] = clone(row)));

    this.xs.loadData({ styles: [], merges, cols, rows });

    this.sheetRenderer();
  }

  public setSheets(sheets: SheetData[]): void {
    const sheetMap: Record<string, ISheet> = this.sheets.reduce(
      (prev, sheet) => ({ ...prev, [sheet.getId()]: sheet }),
      {},
    );
    const existsIdMap: Record<string, boolean> = {};
    const resolved: ISheet[] = [];

    sheets.forEach(sheet => {
      if (sheet.id && sheetMap[sheet.id]) {
        const { id, name, extra } = sheet;
        const existsSheet = sheetMap[id];

        existsIdMap[id] = true;

        if (existsSheet.getName() !== name) {
          existsSheet.setName(name);
        }

        existsSheet.setExtra(extra);
      } else {
        resolved.push(new Sheet(sheet));
      }
    });

    this.sheets.forEach(sheet => {
      if (!existsIdMap[sheet.getId()]) {
        sheet.destroy();
      }
    });

    this.sheets = resolved;

    this.setCurrentSheet();
    this.emit(
      'change',
      resolved.map(sheet => ({ ...sheet.getExtra(), id: sheet.getId(), name: sheet.getName() })),
    );
  }

  public getSheet(): ISheet {
    return this.sheet;
  }

  public changeSheet(index: number): void {
    this.setCurrentSheet(index);
    this.emit('sheet-change', index);
  }

  public getSelectedRows(): TableRow[] {
    return this.chosenRows;
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
      (this.xs as any).sheet.selector.set(rowIndex, colIndex);
    }
  }

  public merge(): Result {
    return this.handleOperationResult(this.table.mergeCells());
  }

  public unmerge(): Result {
    return this.handleOperationResult(this.table.unmergeCells());
  }

  public insertColumns(startColIndex: number, count?: number): Result {
    return this.handleOperationResult(this.table.insertColumns(startColIndex, count));
  }

  public deleteColumns(): Result {
    return this.handleOperationResult(this.table.deleteColumnsInRange());
  }

  public insertRows(startRowIndex: number, count?: number): Result {
    return this.handleOperationResult(this.table.insertRows(startRowIndex, count));
  }

  public deleteRows(): Result {
    return this.handleOperationResult(this.table.deleteRowsInRange());
  }

  public destroy(): void {
    this.sheet = null as any;
    this.table = null as any;

    this.chosenCell = null;
    this.chosenRange = null;

    this.clearRowAndColStatus();

    this.sheets.forEach(sheet => sheet.destroy());

    this.off();
    this.xs.deleteSheet();
  }
}

export default Holysheet;
