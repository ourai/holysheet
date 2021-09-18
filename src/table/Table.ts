import { isString, noop, omit } from '@ntks/toolbox';

import AbstractTable, { getColumnTitle, getColumnIndex } from '../abstract-table';

import {
  CellId,
  ColOverflowCell,
  CellStyle,
  InternalCell,
  TableCell,
  InternalRow,
  TableRow,
  CellData,
  TableRange,
  TableSelection,
  Result,
  ITable,
  TableInitializer,
} from './typing';
import { getTitleCoord, getIndexCoord } from './helper';

class Table extends AbstractTable implements ITable {
  private readonly cellInserted: (cells: TableCell[]) => void;

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

  private getColumnOverflowCells(
    row: InternalRow,
    startColIndex: number,
    endColIndex: number,
  ): ColOverflowCell[] {
    const hitCells: ColOverflowCell[] = [];

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

    return hitCells;
  }

  private getRowOverflowCells(startRowIndex: number, endRowIndex: number): CellId[] {
    const cells: CellId[] = [];

    this.rows.slice(startRowIndex, endRowIndex + 1).forEach(row => {
      row.cells.forEach(cellId => {
        const {
          span = [],
          __meta: { rowIndex },
        } = this.cells[cellId];
        const rowSpan = span[1] || 0;

        if (rowIndex + rowSpan > endRowIndex) {
          cells.push(cellId);
        }
      });
    });

    return cells;
  }

  constructor({ cellInserted, ...others }: TableInitializer) {
    super(others);

    this.cellInserted = cellInserted || noop;
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

  public getCellText(id: CellId): string {
    return (this.cells[id] as InternalCell).text || '';
  }

  public setCellText(id: CellId, text: string): void {
    (this.cells[id] as InternalCell).text = text;
  }

  public getCellStyle(id: CellId): CellStyle {
    return (this.cells[id] as InternalCell).style || {};
  }

  public setCellStyle(id: CellId, style: CellStyle): void {
    (this.cells[id] as InternalCell).style = style;
  }

  public setCellProperties(
    id: CellId,
    properties: Record<string, any>,
    override: boolean = false,
  ): void {
    const reservedKeys = ['mergedCoord'];
    const { text = this.getCellText(id), style = this.getCellStyle(id), ...others } = omit(
      properties,
      reservedKeys,
    );

    this.setCellText(id, text);
    this.setCellStyle(id, style);

    const { mergedCoord } = this.cells[id] as InternalCell;

    if (mergedCoord) {
      others.mergedCoord = mergedCoord;
    }

    super.setCellProperties(id, others, override);
  }

  public fill(cells: CellData[]): void {
    const needRemoveCells: { index: number; cellIndexes: number[] }[] = [];
    const rows: Record<string, CellData[]> = {};

    let maxColIndex = 0;

    cells.forEach(cell => {
      const [colIndexOrTitle, rowIndexOrTitle] = cell.coordinate!;

      const colIndex = isString(colIndexOrTitle)
        ? getColumnIndex(colIndexOrTitle as string)
        : (colIndexOrTitle as number);

      if (colIndex > maxColIndex) {
        maxColIndex = colIndex;
      }

      const rowIndex = isString(rowIndexOrTitle)
        ? Number(rowIndexOrTitle) - 1
        : (rowIndexOrTitle as number);

      if (rows[rowIndex] === undefined) {
        rows[rowIndex] = [];
      }

      rows[rowIndex].push(cell);
    });

    const currentColCount = this.getColumnCount();
    const targetColCount = maxColIndex + 1;

    if (currentColCount < targetColCount) {
      this.insertColumns(currentColCount, targetColCount - currentColCount);
    }

    const currentRowCount = this.getRowCount();
    const targetRowCount = Object.keys(rows).length;

    if (currentRowCount < targetRowCount) {
      this.insertRows(currentRowCount, targetRowCount - currentRowCount);
    }

    Object.keys(rows).forEach(rowIndexKey => {
      const rowIndex = Number(rowIndexKey);
      const cells = rows[rowIndex];
      const needRemove = this.rows[rowIndex].cells.length > cells.length;

      cells.forEach(cell => {
        const { coordinate, span = [], ...others } = omit(cell, [
          '__meta',
          'id',
          'mergedCoord',
        ]) as CellData;

        const [colIndexOrTitle] = coordinate!;
        const colIndex = isString(colIndexOrTitle)
          ? getColumnIndex(colIndexOrTitle as string)
          : (colIndexOrTitle as number);
        const { id } = this.getCell(colIndex, rowIndex)!;

        if (needRemove) {
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
        }

        this.setCellProperties(id, others);
      });
    });

    if (needRemoveCells.length === 0) {
      return;
    }

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
      if (row.cells.length === 0) {
        return;
      }

      const colSpan = eci - sci;

      let cellIndex = -1;

      for (let idx = 0; idx < row.cells.length; idx++) {
        const {
          __meta: { colIndex: cellColIndex },
        } = this.cells[row.cells[idx]] as InternalCell;

        if (cellColIndex === sci) {
          cellIndex = idx;

          break;
        }
      }

      if (ri === 0) {
        const cellId = row.cells[cellIndex];
        const cell = this.cells[cellId] as InternalCell;
        const mergedCoord = getTitleCoord(sci, sri, eci, eri);

        cell.mergedCoord = mergedCoord;
        cell.span = [colSpan, eri - sri];

        this.merged[mergedCoord] = [...this.selection!.range];

        this.markCellAsModified(cellId);
        this.removeCells(row.cells.splice(cellIndex + 1, colSpan));
      } else {
        let endColCellIndex = cellIndex;

        if (cellIndex === -1) {
          if ((this.getCellCoordinate(row.cells[0]) as [number, number])[0] > eci) {
            return;
          }

          cellIndex = 0;
        }

        for (let idx = cellIndex; idx < row.cells.length; idx++) {
          const {
            span = [],
            __meta: { colIndex: cellColIndex },
          } = this.cells[row.cells[idx]] as InternalCell;
          const [cellColSpan = 0] = span;

          if (cellColIndex + cellColSpan === eci) {
            endColCellIndex = idx;

            break;
          }
        }

        this.removeCells(row.cells.splice(cellIndex, endColCellIndex - cellIndex + 1));
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

    const [startColIndex, startRowIndex, endColIndex] = this.selection.range;
    const rowsInRange = this.getInternalRowsInRange();
    const rows: InternalRow[] = this.createRows(
      startRowIndex,
      rowsInRange.length,
      this.getColumnCount(),
      true,
    );

    const inserted: TableCell[] = [];

    rowsInRange.forEach((row, ri) => {
      let targetCellIndex = rows[ri].cells.findIndex(cell => !cell);

      if (targetCellIndex === -1) {
        return;
      }

      row.cells.forEach(cellId => {
        const cell = this.cells[cellId] as InternalCell;

        this.cells[cellId] = cell;

        while (rows[ri].cells[targetCellIndex]) {
          targetCellIndex++;
        }

        rows[ri].cells[targetCellIndex] = cellId;

        const [colIndex] = this.getCellCoordinate(cellId) as [number, number];

        if (cell.span && colIndex >= startColIndex && colIndex <= endColIndex) {
          const [colSpan = 0, rowSpan = 0] = cell.span;

          if (colSpan > 0) {
            const cellIds = this.createCells(
              startRowIndex + ri,
              targetCellIndex + 1,
              colSpan,
            ) as CellId[];

            inserted.push(...cellIds.map(id => this.getCell(id)!));

            this.removeCells(rows[ri].cells.splice(targetCellIndex + 1, colSpan, ...cellIds));
          }

          if (rowSpan > 0) {
            const endRowIndex = ri + rowSpan;

            let nextRowIndex = ri + 1; // FIXME: 当选区包含已合并单元格的一部分时会报错

            while (nextRowIndex <= endRowIndex) {
              const cellCount = colSpan + 1;
              const cellIds = this.createCells(
                startRowIndex + nextRowIndex,
                targetCellIndex,
                cellCount,
              ) as CellId[];

              inserted.push(...cellIds.map(id => this.getCell(id)!));

              this.removeCells(
                rows[nextRowIndex].cells.splice(targetCellIndex, cellCount, ...cellIds),
              );

              nextRowIndex++;
            }
          }

          delete this.cells[cellId].span;

          if (cell.mergedCoord) {
            delete this.merged[cell.mergedCoord];
            delete (this.cells[cellId] as InternalCell).mergedCoord;
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

    rows.forEach(row => (row.cells = row.cells.filter(cellId => !!cellId)));

    this.rows.splice(this.selection.range[1], rows.length, ...rows);

    this.clearSelection();

    if (inserted.length > 0) {
      this.cellInserted(inserted);
    }

    return { success: true };
  }

  public insertColumns(colIndex: number, count: number = 1): Result {
    if (colIndex < 0 || count < 1) {
      return { success: false, message: '插入位置有误或未设置要插入的列数' };
    }

    this.columns.splice(colIndex, 0, ...this.createColumns(count));

    const inserted: TableCell[] = [];

    this.rows.forEach((row, ri) => {
      let cellIndex = -1;
      let toInsert = false;
      let colOverflow = false;

      for (let idx = 0; idx < row.cells.length; idx++) {
        const cellId = row.cells[idx];
        const {
          span = [],
          __meta: { colIndex: cellColIndex },
        } = this.cells[cellId] as InternalCell;
        const [colSpan = 0] = span;

        if (cellColIndex === colIndex || cellColIndex + colSpan >= colIndex) {
          toInsert = cellColIndex === colIndex;
          colOverflow = colSpan > 0 && cellColIndex + colSpan >= colIndex;
          cellIndex = idx;

          break;
        }
      }

      // 在跨列单元格的范围内插入列时需要更新跨列信息
      if (cellIndex > -1 && !toInsert && colOverflow) {
        const cellId = row.cells[cellIndex];
        const { span = [], mergedCoord } = this.cells[cellId] as InternalCell;
        const [colSpan = 0, rowSpan = 0] = span;

        if (mergedCoord) {
          const [sci, sri, eci, eri] = this.merged[mergedCoord];
          const newColSpan = colSpan + count;
          const range: TableRange = [sci, sri, sci + newColSpan, eri];
          const newMergedCoord = getTitleCoord(...range);

          this.cells[cellId].span = [newColSpan, rowSpan];
          (this.cells[cellId] as InternalCell).mergedCoord = newMergedCoord;

          delete this.merged[mergedCoord];

          this.merged[newMergedCoord] = range;
        }
      } else {
        const cellIds = this.createCells(ri, colIndex, count) as CellId[];

        inserted.push(...cellIds.map(id => this.getCell(id)!));

        if (cellIndex > -1) {
          row.cells.splice(cellIndex, 0, ...cellIds);
        } else {
          row.cells.push(...cellIds);
        }
      }

      if (cellIndex === -1) {
        return;
      }

      row.cells
        .slice(colOverflow ? cellIndex + 1 : cellIndex + count)
        .forEach(cellId =>
          this.updateCellCoordinate(
            cellId,
            (this.getCellCoordinate(cellId)[0] as number) + count,
            ri,
          ),
        );
    });

    if (inserted.length > 0) {
      this.cellInserted(inserted);
    }

    return { success: true };
  }

  public deleteColumns(startColIndex: number, count?: number): Result {
    const resolvedCount = count === undefined ? this.getColumnCount() - startColIndex : count;
    const inserted: TableCell[] = [];

    this.columns.splice(startColIndex, resolvedCount);

    if (this.getColumnCount() > 0) {
      const endColIndex = startColIndex + resolvedCount - 1;

      this.rows.forEach((row, ri) => {
        const hitCells = this.getColumnOverflowCells(row, startColIndex, endColIndex);

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

                inserted.push(this.getCell(newCellId)!);
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

    if (inserted.length > 0) {
      this.cellInserted(inserted);
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

  public insertRows(rowIndex: number, count: number = 1): Result {
    if (rowIndex < 0 || count < 1) {
      return { success: false, message: '插入位置有误或未设置要插入的行数' };
    }

    const overflowCells = this.getRowOverflowCells(0, rowIndex - 1);

    this.rows.splice(rowIndex, 0, ...this.createRows(rowIndex, count, this.getColumnCount()));

    // 在跨行单元格的范围内插入行时需要更新跨行信息
    if (overflowCells.length > 0) {
      overflowCells
        .sort((a, b) => (this.cells[a].__meta.colIndex > this.cells[b].__meta.colIndex ? -1 : 1))
        .forEach(cellId => {
          const {
            span = [],
            mergedCoord,
            __meta: { colIndex, rowIndex: cellRowIndex },
          } = this.cells[cellId] as InternalCell;
          const [colSpan = 0, rowSpan = 0] = span;

          if (rowSpan === 0) {
            return;
          }

          this.rows
            .slice(rowIndex, Math.min(cellRowIndex + rowSpan, rowIndex + count - 1) + 1)
            .forEach(row => this.removeCells(row.cells.splice(colIndex, colSpan + 1)));

          if (mergedCoord) {
            const [sci, sri, eci] = this.merged[mergedCoord!];
            const newRowSpan = rowSpan + count;
            const range: TableRange = [sci, sri, eci, sri + newRowSpan];
            const newMergedCoord = getTitleCoord(...range);

            this.cells[cellId].span = [colSpan, newRowSpan];
            (this.cells[cellId] as InternalCell).mergedCoord = newMergedCoord;

            delete this.merged[mergedCoord];

            this.merged[newMergedCoord] = range;
          }
        });
    }

    this.rows.slice(rowIndex + count).forEach(({ cells }) =>
      cells.forEach(cellId => {
        const [colIndex, cellRowIndex] = this.getCellCoordinate(cellId) as [number, number];

        this.updateCellCoordinate(cellId, colIndex, cellRowIndex + count);
      }),
    );

    const inserted: TableCell[] = [];

    this.rows
      .slice(rowIndex, rowIndex + count)
      .forEach(({ cells }) => inserted.push(...cells.map(id => this.getCell(id)!)));

    if (inserted.length > 0) {
      this.cellInserted(inserted);
    }

    return { success: true };
  }

  public deleteRows(startRowIndex: number, count?: number): Result {
    const resolvedCount = count === undefined ? this.getRowCount() - startRowIndex : count;
    const endRowIndex = startRowIndex + resolvedCount - 1;
    const overflowCells = this.getRowOverflowCells(startRowIndex, endRowIndex);
    const inserted: TableCell[] = [];

    if (overflowCells.length > 0) {
      const overflowRowIndex = endRowIndex + 1;
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

        inserted.push(this.getCell(newCellId)!);
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

    if (inserted.length > 0) {
      this.cellInserted(inserted);
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
