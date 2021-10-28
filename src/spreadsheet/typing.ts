import {
  TableCell,
  CellData,
  InternalRow,
  TableRow,
  Result,
  SheetData,
  SheetStyle,
  ISheet,
} from '../sheet';

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

interface Spreadsheet {
  mount(elementOrSelector: MountEl): void;
  render(): void;
  setSheets(sheets: SheetData[], activeSheetIndex?: number): void;
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

type ContextMenuTriggerPosition = 'col-title' | 'row-title' | 'cell';

type ContextMenuItemAsserter = (
  context: Spreadsheet,
  position: ContextMenuTriggerPosition,
) => boolean;

interface ContextMenuItem {
  key: string;
  title: string;
  available?: boolean | ContextMenuItemAsserter;
  handler?: (context: Spreadsheet, data: any) => any;
}

type CellCreator = () => Omit<TableCell, 'id'>;
type RowCreator = () => Omit<InternalRow, 'id' | 'cells'>;

type CellResolver = (cell: TableCell) => Partial<CellData>;

interface SpreadsheetHooks {
  beforeSheetActivate?: (prev: ISheet) => boolean;
  sheetActivated?: (current: ISheet, prev: ISheet) => void;
  beforeSheetRender?: () => boolean;
  sheetRendered?: () => void;
}

interface SpreadsheetOptions extends SpreadsheetHooks {
  el?: MountEl;
  column?: ColumnOptions;
  row?: RowOptions;
  style?: SheetStyle;
  sheetIndex?: number;
  contextMenu?: ContextMenuItem[];
  editable?: boolean;
  hideContextMenu?: boolean;
  cellCreator?: CellCreator;
  rowCreator?: RowCreator;
  cellResolver?: CellResolver;
}

type ResolvedOptions = Required<
  Omit<
    SpreadsheetOptions,
    'el' | 'cellCreator' | 'rowCreator' | 'cellResolver' | keyof SpreadsheetHooks
  >
>;

export {
  MountEl,
  SpreadsheetEvents,
  ContextMenuTriggerPosition,
  ContextMenuItemAsserter,
  ContextMenuItem,
  CellCreator,
  RowCreator,
  CellResolver,
  SpreadsheetOptions,
  ResolvedOptions,
  Spreadsheet,
};
