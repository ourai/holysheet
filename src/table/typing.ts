import { IEventEmitter } from '@ntks/event-emitter';

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

type TableEvents = 'cell-update' | 'cell-change' | 'row-update' | 'row-change';

interface ITable extends IEventEmitter<TableEvents> {
  getCell(id: CellId): TableCell;
  getCell(colIndex: number, rowIndex: number): TableCell;
  getCell(colTitle: string, rowTitle: string): TableCell;
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
}

type CellCreator = () => Omit<TableCell, 'id'>;
type RowCreator = () => Omit<InternalRow, 'id' | 'cells'>;

interface TableInitializer {
  cellCreator: CellCreator;
  rowCreator: RowCreator;
  colCount: number;
  rowCount: number;
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
  TableEvents,
  ITable,
  CellCreator,
  RowCreator,
  TableInitializer,
};
