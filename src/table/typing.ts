import {
  CellId,
  CellMeta,
  InternalCell as _InternalCell,
  InternalRow as _InternalRow,
  TableColumn as _TableColumn,
  ITable as _ITable,
  TableInitializer as _TableInitializer,
} from '../abstract-table';

interface ColOverflowCell {
  id: CellId;
  index: number;
}

interface CellStyle {
  align?: 'left' | 'center' | 'right';
  verticalAlign?: 'top' | 'middle' | 'bottom';
  font?: {
    bold?: boolean;
  };
  color?: string;
  backgroundColor?: string;
  border?: {
    top?: string[];
    right?: string[];
    bottom?: string[];
    left?: string[];
  };
  wrap?: boolean;
}

interface InternalCell extends _InternalCell {
  text: string;
  style?: CellStyle;
  mergedCoord?: string;
}

interface TableCell extends Omit<InternalCell, '__meta' | 'text' | 'style'> {}

interface CellData extends Omit<InternalCell, '__meta' | 'id' | 'mergedCoord'> {
  coordinate?: [number, number] | [string, string];
  [key: string]: any;
}

interface TableColumn extends _TableColumn {}

interface InternalRow extends _InternalRow {}

interface TableRow extends Omit<InternalRow, 'cells'> {
  cells: TableCell[];
}

type StartColIndex = number;
type StartRowIndex = number;
type EndColIndex = number;
type EndRowIndex = number;

type TableRange = [StartColIndex, StartRowIndex, EndColIndex, EndRowIndex];

interface TableSelection {
  cell: TableCell | null;
  range: TableRange;
}

interface Result {
  success: boolean;
  message?: string;
}

interface ITable extends _ITable {
  getCellText(id: CellId): string;
  setCellText(id: CellId, text: string): void;
  getCellStyle(id: CellId): CellStyle;
  setCellStyle(id: CellId, style: CellStyle): void;
  fill(cells: CellData[]): void;
  getSelection(): TableSelection | null;
  setSelection(selection: TableSelection): void;
  clearSelection(): void;
  getRowsInRange(): TableRow[];
  getCellsInRange(): TableCell[];
  getModifiedCellsInRange(): TableCell[];
  getMergedInRange(): string[];
  mergeCells(): Result;
  unmergeCells(): Result;
  insertColumns(colIndex: number, count?: number): Result;
  deleteColumns(startColIndex: number, count?: number): Result;
  deleteColumnsInRange(): Result;
  insertRows(rowIndex: number, count?: number): Result;
  deleteRows(startRowIndex: number, count?: number): Result;
  deleteRowsInRange(): Result;
}

interface TableHooks {
  cellInserted?: (cells: TableCell[]) => void;
}

interface TableInitializer extends _TableInitializer, TableHooks {}

export {
  CellId,
  CellMeta,
  ColOverflowCell,
  CellStyle,
  InternalCell,
  TableCell,
  CellData,
  TableColumn,
  InternalRow,
  TableRow,
  TableRange,
  TableSelection,
  Result,
  ITable,
  TableInitializer,
};
