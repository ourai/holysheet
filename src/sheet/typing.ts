import {
  CellId,
  CellMeta,
  InternalCell,
  InternalRow,
  TableColumn,
  ITable,
  TableInitializer,
} from '../table';

interface InternalSheetCell extends InternalCell {
  mergedCoord?: string;
}

interface SheetCell extends Omit<InternalSheetCell, '__meta'> {}

interface CellData extends Omit<InternalSheetCell, '__meta' | 'id' | 'mergedCoord'> {
  coordinate: [number, number] | [string, string];
  [key: string]: any;
}

interface SheetColumn extends TableColumn {}

interface InternalSheetRow extends InternalRow {}

interface SheetRow extends Omit<InternalSheetRow, 'cells'> {
  cells: SheetCell[];
}

type StartColIndex = number;
type StartRowIndex = number;
type EndColIndex = number;
type EndRowIndex = number;

type SheetRange = [StartColIndex, StartRowIndex, EndColIndex, EndRowIndex];

interface SheetSelection {
  cell: SheetCell | null;
  range: SheetRange;
}

interface Result {
  success: boolean;
  message?: string;
}

interface ISheet extends ITable {
  getName(): string;
  setName(name: string): void;
  fill(cells: CellData[]): void;
  getSelection(): SheetSelection | null;
  setSelection(selection: SheetSelection): void;
  clearSelection(): void;
  getRowsInRange(): SheetRow[];
  getCellsInRange(): SheetCell[];
  getModifiedCellsInRange(): SheetCell[];
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

interface SheetInitializer extends TableInitializer {
  name: string;
}

export {
  CellId,
  CellMeta,
  InternalSheetCell,
  SheetCell,
  CellData,
  SheetColumn,
  InternalSheetRow,
  SheetRow,
  SheetRange,
  SheetSelection,
  Result,
  ISheet,
  SheetInitializer,
};
