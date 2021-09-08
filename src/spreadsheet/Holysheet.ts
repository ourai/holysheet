import { noop } from '@ntks/toolbox';
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

import {
  MountEl,
  RenderCellResolver,
  CellCreator,
  RowCreator,
  SpreadsheetOptions,
  ResolvedOptions,
  Spreadsheet,
} from './typing';
import { resolveOptions, createXSpreadsheetInstance } from './helper';

class Holysheet extends EventEmitter implements Spreadsheet {
  private readonly options: ResolvedOptions;

  private readonly renderCellResolver: RenderCellResolver;

  private readonly cellCreator: CellCreator | undefined;
  private readonly rowCreator: RowCreator | undefined;

  private readonly beforeSheetActivate: (prev: ISheet) => boolean;
  private readonly sheetActivated: (current: ISheet, prev: ISheet) => void;

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

    if (!this.beforeSheetActivate(prevSheet)) {
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

      table.on('cell-update', cell => this.emit('cell-update', cell));
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

  constructor({
    el,
    cellCreator,
    rowCreator,
    renderCellResolver,
    beforeSheetActivate,
    sheetActivated,
    ...others
  }: SpreadsheetOptions = {}) {
    super();

    this.options = resolveOptions(this, others);

    this.cellCreator = cellCreator;
    this.rowCreator = rowCreator;

    this.renderCellResolver = renderCellResolver || ((() => ({ text: '' })) as RenderCellResolver);

    this.beforeSheetActivate = beforeSheetActivate || (() => true);
    this.sheetActivated = sheetActivated || noop;

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
    const merges: string[] = [];

    const cols: Record<string, any> = { len: this.table.getColumnCount() };
    const rows: Record<string, any> = { len: this.table.getRowCount() };

    this.table.getColumns().forEach((col, idx) => {
      if (col.width !== undefined) {
        cols[idx] = { width: col.width };
      }
    });

    this.table
      .transformRows((row, ri) => {
        const { cells, ...others } = row;

        return {
          ...others,
          cells: cells.reduce((prev, tableCell) => {
            const { span, mergedCoord, ...cell } = tableCell as TableCell;
            const resolved = this.renderCellResolver(cell, row, ri);

            if (span) {
              resolved.merge = [span[1], span[0]];
            }

            if (mergedCoord) {
              merges.push(mergedCoord);
            }

            return { ...prev, [this.table.getCellCoordinate(cell.id)[0]]: resolved };
          }, {}),
        };
      })
      .forEach((row, idx) => (rows[idx] = row));

    this.xs.loadData({ styles: [], merges, cols, rows });
  }

  public getTable(): ITable {
    return this.table;
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
      (this.xs as any).sheet.selector.set(colIndex, rowIndex);
    }
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
      this.sheets.map(sheet => ({ ...sheet.getExtra(), id: sheet.getId(), name: sheet.getName() })),
    );
  }

  public changeSheet(index: number): void {
    this.setCurrentSheet(index);
    this.emit('sheet-change', index);
  }

  public updateCell(id: CellId, data: Record<string, any>): void {
    this.table.setCellProperties(id, data);
  }

  public destroy(): void {
    this.sheet = null as any;
    this.table = null as any;

    this.chosenCell = null;
    this.chosenRange = null;

    this.clearRowAndColStatus();

    this.sheets.forEach(sheet => sheet.destroy());

    this.xs.deleteSheet();
  }
}

export default Holysheet;
