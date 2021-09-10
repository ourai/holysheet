import { TableCell, InternalRow, TableRow, Result, SheetData, SheetStyle, ISheet } from '../sheet';

type MountEl = HTMLElement | string;

type SpreadsheetEvents = 'cell-change' | 'range-change' | 'width-change' | 'height-change';

interface ColumnOptions {
  count?: number;
  width?: number;
}

interface RowOptions {
  count?: number;
  height?: number;
}

interface ContextMenuItem {}

type CellCreator = () => Omit<TableCell, 'id'>;
type RowCreator = () => Omit<InternalRow, 'id' | 'cells'>;

interface SpreadsheetHooks {
  beforeSheetActivate?(prev: ISheet): boolean;
  sheetActivated?(current: ISheet, prev: ISheet): void;
  beforeSheetRender?(): boolean;
  sheetRendered?(): void;
}

interface SpreadsheetOptions extends SpreadsheetHooks {
  el?: MountEl;
  column?: ColumnOptions;
  row?: RowOptions;
  style?: SheetStyle;
  contextMenu?: ContextMenuItem[];
  editable?: boolean;
  cellCreator?: CellCreator;
  rowCreator?: RowCreator;
}

type ResolvedOptions = Required<
  Omit<
    SpreadsheetOptions,
    'el' | 'cellCreator' | 'rowCreator' | 'renderCellResolver' | keyof SpreadsheetHooks
  >
>;

interface Spreadsheet {
  mount(elementOrSelector: MountEl): void;
  render(): void;
  setSheets(sheets: SheetData[]): void;
  getSheet(): ISheet;
  changeSheet(index: number): void;
  getSelectedRows(): TableRow[];
  select(colIndex: number, rowIndex: number): void;
  select(
    startColIndex: number,
    startRowIndex: number,
    endColIndex: number,
    endRowIndex: number,
  ): void;
  merge(): Result;
  unmerge(): Result;
  insertColumns(startColIndex: number, count?: number): Result;
  deleteColumns(): Result;
  insertRows(startRowIndex: number, count?: number): Result;
  deleteRows(): Result;
  destroy(): void;
}

export {
  MountEl,
  SpreadsheetEvents,
  ContextMenuItem,
  CellCreator,
  RowCreator,
  SpreadsheetOptions,
  ResolvedOptions,
  Spreadsheet,
};
