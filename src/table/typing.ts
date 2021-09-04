import {
  CellId,
  CellMeta,
  InternalCell as _InternalCell,
  InternalRow as _InternalRow,
  TableColumn as _TableColumn,
  ITable as _ITable,
  TableInitializer as _TableInitializer,
} from '../abstract-table';

interface InternalCell extends _InternalCell {
  mergedCoord?: string;
}

interface TableCell extends Omit<InternalCell, '__meta'> {}

interface CellData extends Omit<InternalCell, '__meta' | 'id' | 'mergedCoord'> {
  coordinate: [number, number] | [string, string];
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
  insertColumn(colIndex: number, count?: number): Result;
  deleteColumns(startColIndex: number, count?: number): Result;
  deleteColumnsInRange(): Result;
  insertRow(rowIndex: number, count?: number): Result;
  deleteRows(startRowIndex: number, count?: number): Result;
  deleteRowsInRange(): Result;
}

interface TableInitializer extends _TableInitializer {}

export {
  CellId,
  CellMeta,
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
