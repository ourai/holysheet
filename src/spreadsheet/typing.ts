import { CellStyle, CellData } from '@wotaware/x-spreadsheet';

import {
  CellId,
  TableCell,
  InternalRow,
  TableRow,
  ITable,
  SheetData,
  SheetStyle,
  ISheet,
} from '../sheet';

type MountEl = HTMLElement | string;

interface RenderCell extends Omit<CellData, 'style'> {
  style?: CellStyle;
}

type RenderCellResolver = (cell: TableCell, row: TableRow, rowIndex: number) => RenderCell;

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
  renderCellResolver?: RenderCellResolver;
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
  getTable(): ITable;
  getSelectedRows(): TableRow[];
  select(colIndex: number, rowIndex: number): void;
  select(
    startColIndex: number,
    startRowIndex: number,
    endColIndex: number,
    endRowIndex: number,
  ): void;
  setSheets(sheets: SheetData[]): void;
  changeSheet(index: number): void;
  updateCell(id: CellId, data: Record<string, any>): void;
  updateCellText(id: CellId, text: string): void;
  destroy(): void;
}

export {
  MountEl,
  RenderCell,
  RenderCellResolver,
  SpreadsheetEvents,
  ContextMenuItem,
  CellCreator,
  RowCreator,
  SpreadsheetOptions,
  ResolvedOptions,
  Spreadsheet,
};
