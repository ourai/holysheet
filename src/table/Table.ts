import { isString, omit } from '@ntks/toolbox';

import AbstractTable, { getColumnTitle, getColumnIndex } from '../abstract-table';

import {
  CellId,
  InternalCell,
  TableCell,
  InternalRow,
  TableRow,
  CellData,
  TableRange,
  TableSelection,
  Result,
  ITable,
} from './typing';
import { getTitleCoord, getIndexCoord } from './helper';

class Table extends AbstractTable implements ITable {
  private selection: TableSelection | null = null;

  private merged: Record<string, TableRange> = {};

  private getInternalRowsInRange(): InternalRow[] {
    return this.selection
      ? this.rows.slice(this.selection.range[1], this.selection.range[3] + 1)
      : [];
  }

  private updateCellCoordinate(id: CellId, colIndex: number, rowIndex: number): void {
    this.setCellCoordinate(id, colIndex, rowIndex);

    const { span = [], mergedCoord } = this.cells[id] as InternalCell;
    const [colSpan = 0, rowSpan = 0] = span;

    if (colSpan === 0 && rowSpan === 0) {
      return;
    }

    const range: TableRange = [colIndex, rowIndex, colIndex + colSpan, rowIndex + rowSpan];
    const newMergedCoord = getTitleCoord(...range);

    (this.cells[id] as InternalCell).mergedCoord = newMergedCoord;

    delete this.merged[mergedCoord!];
    this.merged[newMergedCoord] = range;
  }

  public static getColumnTitle(index: number): string {
    return getColumnTitle(index);
  }

  public static getColumnIndex(title: string): number {
    return getColumnIndex(title);
  }

  public static getCoordTitle(
    colIndex: number,
    rowIndex: number,
    endColIndex?: number,
    endRowIndex?: number,
  ): string {
    return getTitleCoord(colIndex, rowIndex, endColIndex, endRowIndex);
  }

  public setCellProperties(id: CellId, properties: Record<string, any>): void {
    super.setCellProperties(id, omit(properties, ['mergedCoord']));
  }

  public fill(cells: CellData[]): void {
    const needRemoveCells: { index: number; cellIndexes: number[] }[] = [];

    cells.forEach(cell => {
      const {
        coordinate,
        span = [],
        ...others
      } = omit(cell, ['__meta', 'id', 'mergedCoord']) as CellData;

      const [colIndexOrTitle, rowIndexOrTitle] = coordinate;

      let colIndex: number;
      let rowIndex: number;

      if (isString(colIndexOrTitle) && isString(rowIndexOrTitle)) {
        colIndex = getColumnIndex(colIndexOrTitle as string);
        rowIndex = Number(rowIndexOrTitle) - 1;
      } else {
        colIndex = colIndexOrTitle as number;
        rowIndex = rowIndexOrTitle as number;
      }

      const { id } = this.getCell(colIndex, rowIndex);
      const [colSpan = 0, rowSpan = 0] = span;

      if (colSpan > 0 || rowSpan > 0) {
        const range: TableRange = [colIndex, rowIndex, colIndex + colSpan, rowIndex + rowSpan];
        const mergedCoord = getTitleCoord(...range);

        this.cells[id].span = [colSpan, rowSpan];
        (this.cells[id] as InternalCell).mergedCoord = mergedCoord;

        this.merged[mergedCoord] = range;

        if (colSpan > 0) {
          const indexArr: number[] = [];

          let nextColIndex = colIndex + 1;

          while (nextColIndex <= colIndex + colSpan) {
            indexArr.push(nextColIndex);

            nextColIndex++;
          }

          needRemoveCells.push({ index: rowIndex, cellIndexes: indexArr });
        }

        if (rowSpan > 0) {
          let nextRowIndex = rowIndex + 1;

          while (nextRowIndex <= rowIndex + rowSpan) {
            const indexArr: number[] = [];

            let nextColIndex = colIndex;

            while (nextColIndex <= colIndex + colSpan) {
              indexArr.push(nextColIndex);

              nextColIndex++;
            }

            needRemoveCells.push({ index: nextRowIndex, cellIndexes: indexArr });

            nextRowIndex++;
          }
        }
      }

      this.setCellProperties(id, others);
    });

    needRemoveCells
      .sort((a, b) => (a.index < b.index ? -1 : 1))
      .forEach(({ index, cellIndexes }) =>
        cellIndexes
          .sort((a, b) => (a > b ? -1 : 1))
          .forEach(cellIndex => this.removeCells(this.rows[index].cells.splice(cellIndex, 1))),
      );
  }

  public getSelection(): TableSelection | null {
    return this.selection;
  }

  public setSelection(selection: TableSelection): void {
    this.selection = selection;
  }

  public clearSelection(): void {
    this.selection = null;
  }

  public getRowsInRange(): TableRow[] {
    return this.getTableRows(this.getInternalRowsInRange());
  }

  public getCellsInRange(): TableCell[] {
    if (!this.selection) {
      return [];
    }

    const cellsInRange: TableCell[] = [];
    const [sci, _, eci] = this.selection.range;

    this.getRowsInRange().forEach(({ cells }) =>
      cells.forEach(cell => {
        const [colIndex] = this.getCellCoordinate(cell.id);

        if (colIndex >= sci && colIndex <= eci) {
          cellsInRange.push(cell);
        }
      }),
    );

    return cellsInRange;
  }

  public getModifiedCellsInRange(): TableCell[] {
    return this.getCellsInRange().filter(({ id }) => this.isCellModified(id));
  }

  public getMergedInRange(): string[] {
    return this.selection
      ? Object.keys(this.merged).filter(coord => {
          const [sci, sri, eci, eri] = this.selection!.range;
          const [msci, msri, meci, meri] = this.merged[coord];

          return msci >= sci && msri >= sri && meci <= eci && meri <= eri;
        })
      : [];
  }

  public mergeCells(): Result {
    if (!this.selection) {
      return { success: false, message: '请先选择要合并的单元格' };
    }

    const [sci, sri, eci, eri] = this.selection.range;
    const rowsInRange = this.getInternalRowsInRange();

    rowsInRange.forEach((row, ri) => {
      const colSpan = eci - sci;

      if (ri === 0) {
        const cellId = row.cells[sci];
        const cell = this.cells[cellId] as InternalCell;
        const mergedCoord = getTitleCoord(sci, sri, eci, eri);

        cell.mergedCoord = mergedCoord;
        cell.span = [colSpan, eri - sri];

        this.merged[mergedCoord] = [...this.selection!.range];

        this.markCellAsModified(cellId);
        this.removeCells(row.cells.splice(sci + 1, colSpan));
      } else {
        this.removeCells(row.cells.splice(sci, colSpan + 1));
      }
    });

    this.rows.splice(sri, rowsInRange.length, ...rowsInRange);

    this.clearSelection();

    return { success: true };
  }

  public unmergeCells(): Result {
    if (!this.selection) {
      return { success: false, message: '请先选择要取消合并的单元格' };
    }

    const startRowIndex = this.selection.range[1];
    const rowsInRange = this.getInternalRowsInRange();
    const rows: InternalRow[] = this.createRows(
      startRowIndex,
      rowsInRange.length,
      this.getColumnCount(),
      true,
    );

    rowsInRange.forEach((row, ri) => {
      let targetCellIndex = rows[ri].cells.findIndex(cell => !cell);

      if (targetCellIndex === -1) {
        return;
      }

      row.cells.forEach(cellId => {
        const { span, mergedCoord, ...pureCell } = this.cells[cellId] as InternalCell;

        this.cells[cellId] = pureCell;

        while (rows[ri].cells[targetCellIndex]) {
          targetCellIndex++;
        }

        rows[ri].cells[targetCellIndex] = cellId;

        if (span) {
          const [colSpan = 0, rowSpan = 0] = span;

          if (colSpan > 0) {
            this.removeCells(
              rows[ri].cells.splice(
                targetCellIndex + 1,
                colSpan,
                ...(this.createCells(startRowIndex + ri, targetCellIndex + 1, colSpan) as CellId[]),
              ),
            );
          }

          if (rowSpan > 0) {
            const endRowIndex = ri + rowSpan;

            let nextRowIndex = ri + 1; // FIXME: 当选区包含已合并单元格的一部分时会报错

            while (nextRowIndex <= endRowIndex) {
              const cellCount = colSpan + 1;

              this.removeCells(
                rows[nextRowIndex].cells.splice(
                  targetCellIndex,
                  cellCount,
                  ...(this.createCells(
                    startRowIndex + nextRowIndex,
                    targetCellIndex,
                    cellCount,
                  ) as CellId[]),
                ),
              );

              nextRowIndex++;
            }
          }

          if (mergedCoord) {
            delete this.merged[mergedCoord];
          }

          this.markCellAsModified(cellId);

          targetCellIndex += colSpan + 1;
        } else {
          targetCellIndex++;
        }
      });

      const copyProps = omit(row, ['id', 'cells']);

      Object.keys(copyProps).forEach(prop => (rows[ri][prop] = copyProps[prop]));
    });

    this.rows.splice(this.selection.range[1], rows.length, ...rows);

    this.clearSelection();

    return { success: true };
  }

  public insertColumn(colIndex: number, count: number = 1): Result {
    if (colIndex < 0 || count < 1) {
      return { success: false, message: '插入位置有误或未设置要插入的列数' };
    }

    this.columns.splice(colIndex, 0, ...this.createColumns(count));

    this.rows.forEach((row, ri) => {
      row.cells.splice(colIndex, 0, ...(this.createCells(ri, colIndex, count) as CellId[]));

      row.cells
        .slice(colIndex + count)
        .forEach(cellId =>
          this.updateCellCoordinate(
            cellId,
            (this.getCellCoordinate(cellId)[0] as number) + count,
            ri,
          ),
        );
    });

    return { success: true };
  }

  public deleteColumns(startColIndex: number, count?: number): Result {
    const resolvedCount = count === undefined ? this.getColumnCount() - startColIndex : count;

    this.columns.splice(startColIndex, resolvedCount);

    if (this.getColumnCount() > 0) {
      const endColIndex = startColIndex + resolvedCount - 1;

      this.rows.forEach((row, ri) => {
        const hitCells: { id: CellId; index: number }[] = [];

        for (let idx = 0; idx < row.cells.length; idx++) {
          const cellId = row.cells[idx];
          const {
            span = [],
            __meta: { colIndex },
          } = this.cells[cellId];
          const calcColIndex = colIndex + (span[0] || 0);

          if (calcColIndex >= startColIndex && colIndex <= endColIndex) {
            hitCells.push({ id: cellId, index: idx });
          }

          if (calcColIndex >= endColIndex) {
            break;
          }
        }

        if (hitCells.length > 0) {
          for (let idx = hitCells.length - 1; idx >= 0; idx--) {
            const { id: cellId, index: cellIndex } = hitCells[idx];
            const {
              span = [],
              mergedCoord,
              __meta: { colIndex, rowIndex },
            } = this.cells[cellId] as InternalCell;
            const [colSpan = 0, rowSpan = 0] = span;
            const calcColIndex = colIndex + colSpan;

            if (mergedCoord) {
              delete this.merged[mergedCoord];
            }

            if (colSpan > 0 && (colIndex < startColIndex || calcColIndex > endColIndex)) {
              if (colIndex < startColIndex) {
                const newColSpan =
                  colSpan -
                  (calcColIndex >= endColIndex ? resolvedCount : calcColIndex - startColIndex + 1);

                if (newColSpan === 0 && rowSpan === 0) {
                  delete this.cells[cellId].span;
                  delete (this.cells[cellId] as InternalCell).mergedCoord;
                } else {
                  const range: TableRange = [
                    colIndex,
                    rowIndex,
                    colIndex + newColSpan,
                    rowIndex + rowSpan,
                  ];
                  const newMergedCoord = getTitleCoord(...range);

                  this.cells[cellId].span = [newColSpan, rowSpan];
                  (this.cells[cellId] as InternalCell).mergedCoord = newMergedCoord;

                  this.merged[newMergedCoord] = range;
                }
              } else if (calcColIndex > endColIndex) {
                const remained = calcColIndex - endColIndex;
                const newColSpan = remained - 1;
                const newStartColIndex = endColIndex + 1;
                const newCellId = this.createCells(ri, newStartColIndex, 1).pop()!;

                if (newColSpan > 0 || rowSpan > 0) {
                  const range: TableRange = [
                    newStartColIndex,
                    ri,
                    newStartColIndex + newColSpan,
                    ri + rowSpan,
                  ];
                  const newMergedCoord = getTitleCoord(...range);

                  this.cells[newCellId].span = [newColSpan, rowSpan];
                  (this.cells[newCellId] as InternalCell).mergedCoord = newMergedCoord;

                  this.merged[newMergedCoord] = range;
                }

                this.removeCells(row.cells.splice(cellIndex, 1, newCellId));
              }
            } else {
              this.removeCells(row.cells.splice(cellIndex, 1));
            }
          }
        }

        row.cells.forEach(cellId => {
          const {
            span = [],
            mergedCoord,
            __meta: { colIndex },
          } = this.cells[cellId] as InternalCell;

          if (colIndex > endColIndex) {
            const newColIndex = colIndex - resolvedCount;

            this.cells[cellId].__meta.colIndex = newColIndex;

            if (mergedCoord) {
              delete this.merged[mergedCoord];

              const [sci, sri, eci, eri] = getIndexCoord(mergedCoord);
              const range: TableRange = [newColIndex, sri, newColIndex + (span[0] || 0), eri];
              const newMergedCoord = getTitleCoord(...range);

              (this.cells[cellId] as InternalCell).mergedCoord = newMergedCoord;

              this.merged[newMergedCoord] = range;
            }
          }
        });
      });
    } else {
      this.rows.forEach(row => this.removeCells(row.cells.splice(0)));
      this.merged = {};
    }

    return { success: true };
  }

  public deleteColumnsInRange(): Result {
    if (!this.selection) {
      return { success: false, message: '没有选中区域' };
    }

    const [sci, sri, eci, eri] = this.selection.range;

    return eri - sri + 1 === this.getRowCount()
      ? this.deleteColumns(sci, eci - sci + 1)
      : { success: false, message: '没有选中整列' };
  }

  public insertRow(rowIndex: number, count: number = 1): Result {
    if (rowIndex < 0 || count < 1) {
      return { success: false, message: '插入位置有误或未设置要插入的行数' };
    }

    this.rows.splice(rowIndex, 0, ...this.createRows(rowIndex, count, this.getColumnCount()));

    this.rows.slice(rowIndex + count).forEach(({ cells }) =>
      cells.forEach(cellId => {
        const [colIndex, rowIndex] = this.getCellCoordinate(cellId) as [number, number];

        this.updateCellCoordinate(cellId, colIndex, rowIndex + count);
      }),
    );

    return { success: true };
  }

  public deleteRows(startRowIndex: number, count?: number): Result {
    const resolvedCount = count === undefined ? this.getRowCount() - startRowIndex : count;
    const endRowIndex = startRowIndex + resolvedCount - 1;

    const overflowRowIndex = endRowIndex + 1;
    const overflowCells: CellId[] = [];

    this.rows.slice(startRowIndex, overflowRowIndex).forEach(row => {
      row.cells.forEach(cellId => {
        const {
          span = [],
          __meta: { rowIndex },
        } = this.cells[cellId];
        const rowSpan = span[1] || 0;

        if (rowIndex + rowSpan > endRowIndex) {
          overflowCells.push(cellId);
        }
      });
    });

    if (overflowCells.length > 0) {
      const overflowCellRow = this.rows[overflowRowIndex];

      const cells = this.createCells(overflowRowIndex, 0, this.getColumnCount(), true);

      overflowCellRow.cells.forEach(cellId => (cells[this.cells[cellId].__meta.colIndex] = cellId));

      overflowCells.forEach(cellId => {
        const {
          span = [],
          __meta: { colIndex, rowIndex },
        } = this.cells[cellId];
        const [colSpan = 0, rowSpan = 0] = span;
        const newRowSpan = rowSpan - (endRowIndex - rowIndex + 1);
        const newCellId = this.createCells(overflowRowIndex, colIndex, 1).pop()!;

        if (colSpan > 0 || newRowSpan > 0) {
          const range: TableRange = [
            colIndex,
            overflowRowIndex,
            colIndex + colSpan,
            overflowRowIndex + newRowSpan,
          ];
          const mergedCoord = getTitleCoord(...range);

          this.cells[newCellId].span = [colSpan, newRowSpan];
          (this.cells[newCellId] as InternalCell).mergedCoord = mergedCoord;

          this.merged[mergedCoord] = range;
        }

        cells[colIndex] = newCellId;
      });

      overflowCellRow.cells = cells.filter(cellId => !!cellId) as CellId[];
    }

    this.rows.splice(startRowIndex, resolvedCount).forEach(row => {
      row.cells.forEach(cellId => {
        const { mergedCoord } = this.cells[cellId] as InternalCell;

        if (mergedCoord) {
          delete this.merged[mergedCoord];
        }
      });

      this.removeCells(row.cells);
    });

    if (this.getRowCount() > 0) {
      if (startRowIndex > 0) {
        let noRowSpanCellRowIndex = startRowIndex - 1;

        while (this.rows[noRowSpanCellRowIndex].cells.length !== this.getColumnCount()) {
          noRowSpanCellRowIndex--;
        }

        const baseRowindex = noRowSpanCellRowIndex;

        this.rows.slice(baseRowindex, startRowIndex).forEach((row, ri) => {
          row.cells.forEach(cellId => {
            const { span = [] } = this.cells[cellId];
            const [colSpan = 0, rowSpan = 0] = span;
            const bottomRowIndex = baseRowindex + ri + rowSpan;

            if (rowSpan === 0 || bottomRowIndex < startRowIndex) {
              return;
            }

            delete this.merged[(this.cells[cellId] as InternalCell).mergedCoord!];

            const newRowSpan =
              rowSpan -
              (bottomRowIndex >= endRowIndex ? resolvedCount : bottomRowIndex - startRowIndex + 1);

            if (colSpan === 0 && newRowSpan === 0) {
              delete this.cells[cellId].span;
              delete (this.cells[cellId] as InternalCell).mergedCoord;
            } else {
              this.cells[cellId].span = [colSpan, newRowSpan];

              const [sci, sri, eci] = getIndexCoord(
                (this.cells[cellId] as InternalCell).mergedCoord!,
              );
              const range: TableRange = [sci, sri, eci, sri + newRowSpan];
              const mergedCoord = getTitleCoord(...range);

              (this.cells[cellId] as InternalCell).mergedCoord = mergedCoord;

              this.merged[mergedCoord] = range;
            }
          });
        });
      }

      this.rows.slice(startRowIndex, this.getRowCount()).forEach((row, ri) => {
        row.cells.forEach(cellId => {
          const { mergedCoord, span = [] } = this.cells[cellId] as InternalCell;
          const newRowIndex = startRowIndex + ri;

          this.cells[cellId].__meta.rowIndex = newRowIndex;

          if (mergedCoord) {
            const [sci, _, eci] = getIndexCoord(mergedCoord);
            const range: TableRange = [sci, newRowIndex, eci, newRowIndex + (span[1] || 0)];
            const newMergedCoord = getTitleCoord(...range);

            (this.cells[cellId] as InternalCell).mergedCoord = newMergedCoord;
            this.merged[newMergedCoord] = range;

            delete this.merged[mergedCoord];
          }
        });
      });
    } else {
      this.merged = {};
    }

    return { success: true };
  }

  public deleteRowsInRange(): Result {
    if (!this.selection) {
      return { success: false, message: '没有选中区域' };
    }

    const [sci, sri, eci, eri] = this.selection.range;

    return eci - sci + 1 === this.getColumnCount()
      ? this.deleteRows(sri, eri - sri + 1)
      : { success: false, message: '没有选中整行' };
  }

  public destroy(): void {
    this.selection = null;
    this.merged = {};

    super.destroy();
  }
}

export default Table;
