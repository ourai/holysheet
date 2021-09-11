type ColSpan = number;
type RowSpan = number;

type CellId = string;

interface CellMeta {
  colIndex: number;
  rowIndex: number;
  modified: boolean;
}

interface InternalCell {
  __meta: CellMeta;
  id: CellId;
  span?: [ColSpan, RowSpan];
}

interface TableCell extends Omit<InternalCell, '__meta'> {}

type ColumnId = string;

interface InternalColumn {
  id: ColumnId;
  width?: number;
}

interface TableColumn extends InternalColumn {}

type RowId = string;

interface InternalRow {
  id: RowId;
  cells: CellId[];
  height?: number;
}

interface TableRow extends Omit<InternalRow, 'cells'> {
  cells: TableCell[];
}

type StartColIndex = number;
type StartRowIndex = number;

type CellCoordinate = [StartColIndex, StartRowIndex] | [string, string];

type RowFilter = (row: TableRow, index: number) => boolean;
type RowMapFn<T> = (row: TableRow, index: number) => T;

interface ITable {
  getCell(id: CellId): TableCell | undefined;
  getCell(colIndex: number, rowIndex: number): TableCell | undefined;
  getCell(colTitle: string, rowTitle: string): TableCell | undefined;
  getCellCoordinate(id: CellId, title?: boolean): CellCoordinate;
  setCellProperties(id: CellId, properties: Record<string, any>): void;
  isCellModified(id: CellId): boolean;
  getColumnCount(): number;
  getColumnWidth(indexOrTitle: number | string): number | undefined;
  setColumnWidth(indexOrTitle: number | string, width: number | 'auto'): void;
  getColumns(): TableColumn[];
  getRowCount(): number;
  getRowHeight(indexOrTitle: number | string): number | undefined;
  setRowHeight(indexOrTitle: number | string, height: number | 'auto'): void;
  getRow(rowIndex: number): TableRow;
  getRows(filter?: RowFilter): TableRow[];
  transformRows<T extends any = TableRow>(mapFn: RowMapFn<T>): T[];
  getRowPropertyValue(rowIndex: number, propertyName: string): any;
  setRowPropertyValue(rowIndex: number, propertyName: string, propertyValue: any): void;
  setRowsPropertyValue(
    startRowIndex: number,
    endRowIndex: number,
    propertyName: string,
    propertyValue: any,
  ): void;
  destroy(): void;
}

interface TableHooks {
  cellUpdated?: (cell: TableCell) => void;
  rowUpdated?: (row: TableRow) => void;
}

type CellCreator = () => Omit<TableCell, 'id'>;
type RowCreator = () => Omit<InternalRow, 'id' | 'cells'>;

interface TableInitializer extends TableHooks {
  columnCount: number;
  rowCount: number;
  cellCreator?: CellCreator;
  rowCreator?: RowCreator;
}

export {
  CellId,
  CellMeta,
  InternalCell,
  TableCell,
  InternalColumn,
  TableColumn,
  InternalRow,
  TableRow,
  CellCoordinate,
  RowFilter,
  RowMapFn,
  ITable,
  CellCreator,
  RowCreator,
  TableInitializer,
};
