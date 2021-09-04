import { isNumber, isString, generateRandomId, includes, omit, clone } from '@ntks/toolbox';
import EventEmitter from '@ntks/event-emitter';

import {
  CellId,
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
} from './typing';
import { getColumnTitle, getColumnIndex } from './helper';

class AbstractTable extends EventEmitter<TableEvents> implements ITable {
  private readonly cellCreator: CellCreator;
  private readonly rowCreator: RowCreator;

  protected columns: InternalColumn[] = [];
  protected rows: InternalRow[] = [];

  protected cells: Record<CellId, InternalCell> = {};

  private resolveColumnIndex(indexOrTitle: number | string): number {
    return isNumber(indexOrTitle)
      ? (indexOrTitle as number)
      : getColumnIndex(indexOrTitle as string);
  }

  private resolveRowIndex(indexOrTitle: number | string): number {
    return isNumber(indexOrTitle) ? (indexOrTitle as number) : Number(indexOrTitle) - 1;
  }

  private createCellOrPlaceholder(
    rowIndex: number,
    colIndex: number,
    empty: boolean = false,
  ): CellId | undefined {
    if (empty) {
      return undefined;
    }

    const id = `${generateRandomId('AbstractTableCell')}${Object.keys(this.cells).length}`;

    this.cells[id] = { ...this.cellCreator(), id, __meta: { rowIndex, colIndex, modified: false } };

    return id;
  }

  protected createCells(
    rowIndex: number,
    startColIndex: number,
    count: number,
    empty?: boolean,
  ): (CellId | undefined)[] {
    const cells: (CellId | undefined)[] = [];

    if (count > 0) {
      for (let ci = 0; ci < count; ci++) {
        cells.push(this.createCellOrPlaceholder(rowIndex, startColIndex + ci, empty));
      }
    }

    return cells;
  }

  protected removeCells(ids: (CellId | undefined)[]): InternalCell[] {
    const cells: InternalCell[] = [];

    ids.forEach(id => {
      if (id && this.cells[id]) {
        cells.push(this.cells[id]);

        delete this.cells[id];
      }
    });

    return cells;
  }

  protected markCellAsModified(id: CellId): void {
    this.cells[id].__meta.modified = true;
  }

  private getTableCell(
    idOrColIndex: CellId | number,
    rowIndexOrTitle?: number | string,
  ): TableCell {
    let cellId: CellId;

    if (isString(idOrColIndex)) {
      if (isString(rowIndexOrTitle)) {
        const colIndex = getColumnIndex(idOrColIndex as string);

        cellId = this.rows[Number(rowIndexOrTitle as string) - 1].cells.find(
          id => this.cells[id].__meta.colIndex === colIndex,
        )!;
      } else {
        cellId = idOrColIndex as CellId;
      }
    } else {
      cellId = this.rows[rowIndexOrTitle as number].cells.find(
        id => this.cells[id].__meta.colIndex === (idOrColIndex as number),
      )!;
    }

    return omit(this.cells[cellId], ['__meta']);
  }

  protected setCellCoordinate(id: CellId, colIndex: number, rowIndex: number): void {
    this.cells[id].__meta.colIndex = colIndex;
    this.cells[id].__meta.rowIndex = rowIndex;
  }

  protected createColumns(colCount: number): InternalColumn[] {
    const cols: InternalColumn[] = [];

    for (let ci = 0; ci < colCount; ci++) {
      cols.push({ id: `${generateRandomId('AbstractTableColumn')}${ci}` });
    }

    return cols;
  }

  protected createRows(
    startRowIndex: number,
    rowCount: number,
    colCount: number,
    placeholderCell?: boolean,
  ): InternalRow[] {
    const rows: InternalRow[] = [];

    for (let ri = 0; ri < rowCount; ri++) {
      rows.push({
        ...this.rowCreator(),
        id: `${generateRandomId('AbstractTableRow')}${ri}`,
        cells: this.createCells(startRowIndex + ri, 0, colCount, placeholderCell) as CellId[],
      });
    }

    return rows;
  }

  protected getTableRows(internalRows: InternalRow[] = this.rows): TableRow[] {
    return internalRows.map(({ cells, ...others }) => ({
      ...others,
      cells: cells.map(id => this.getTableCell(id)),
    }));
  }

  constructor({ cellCreator, rowCreator, colCount, rowCount }: TableInitializer) {
    super();

    this.cellCreator = cellCreator;
    this.rowCreator = rowCreator;

    this.columns = this.createColumns(colCount);
    this.rows = this.createRows(0, rowCount, colCount);
  }

  public getCell(idOrColIndex: CellId | number, rowIndexOrTitle?: number | string): TableCell {
    return this.getTableCell(idOrColIndex, rowIndexOrTitle);
  }

  public getCellCoordinate(id: CellId, title: boolean = false): CellCoordinate {
    const { colIndex, rowIndex } = this.cells[id].__meta;

    return title ? [getColumnTitle(colIndex), `${rowIndex + 1}`] : [colIndex, rowIndex];
  }

  public setCellProperties(id: CellId, properties: Record<string, any>): void {
    const props = omit(properties, ['__meta', 'id', 'span']);
    const propKeys = Object.keys(props);

    if (propKeys.length === 0) {
      return;
    }

    const cell = this.cells[id];

    propKeys.forEach(key => (cell[key] = props[key]));

    this.markCellAsModified(id);

    this.emit('cell-update', this.getCell(id));
  }

  public isCellModified(id: CellId): boolean {
    return this.cells[id].__meta.modified === true;
  }

  public getColumnCount(): number {
    return this.columns.length;
  }

  public getColumnWidth(indexOrTitle: number | string): number | undefined {
    return this.columns[this.resolveColumnIndex(indexOrTitle)].width;
  }

  public setColumnWidth(indexOrTitle: number | string, width: number | 'auto'): void {
    const colIndex = this.resolveColumnIndex(indexOrTitle);

    if (isNumber(width)) {
      this.columns[colIndex].width = width as number;
    } else {
      delete this.columns[colIndex].width;
    }
  }

  public getColumns(): TableColumn[] {
    return clone(this.columns);
  }

  public getRowCount(): number {
    return this.rows.length;
  }

  public getRowHeight(indexOrTitle: number | string): number | undefined {
    return this.rows[this.resolveRowIndex(indexOrTitle)].height;
  }

  public setRowHeight(indexOrTitle: number | string, height: number | 'auto'): void {
    const rowIndex = this.resolveRowIndex(indexOrTitle);

    if (isNumber(height)) {
      this.rows[rowIndex].height = height as number;
    } else {
      delete this.rows[rowIndex].height;
    }
  }

  public getRow(rowIndex: number): TableRow {
    return clone(this.getTableRows([this.rows[rowIndex]])[0]);
  }

  public getRows(filter?: RowFilter): TableRow[] {
    const rows = this.getTableRows();

    return clone(filter ? rows.filter(filter) : rows);
  }

  public transformRows<T extends any = TableRow>(mapFunc: RowMapFn<T>): T[] {
    return clone(this.getTableRows()).map(mapFunc);
  }

  public getRowPropertyValue(rowIndex: number, propertyName: string): any {
    const row = this.rows[rowIndex];

    return row ? clone(row[propertyName]) : undefined;
  }

  public setRowPropertyValue(rowIndex: number, propertyName: string, propertyValue: any): void {
    const row = this.rows[rowIndex];

    if (!row || includes(propertyName, ['id', 'cells', 'height'])) {
      return;
    }

    row[propertyName] = propertyValue;

    this.emit('row-update', this.getRow(rowIndex));
  }

  public setRowsPropertyValue(
    startRowIndex: number,
    endRowIndex: number,
    propertyName: string,
    propertyValue: any,
  ): void {
    let ri = startRowIndex;

    while (ri <= endRowIndex) {
      this.setRowPropertyValue(ri, propertyName, propertyValue);

      ri++;
    }
  }
}

export default AbstractTable;
