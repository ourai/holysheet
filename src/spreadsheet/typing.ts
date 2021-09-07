import { CellStyle, CellData } from '@wotaware/x-spreadsheet';

import { CellId, TableCell, TableRow, SheetData } from '../sheet';

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

interface SpreadsheetOptions {
  el?: MountEl;
  column?: ColumnOptions;
  row?: RowOptions;
  contextMenu?: ContextMenuItem[];
  editable?: boolean;
  renderCellResolver?: RenderCellResolver;
}

type ResolvedOptions = Required<Omit<SpreadsheetOptions, 'el' | 'renderCellResolver'>>;

interface Spreadsheet {
  mount(elementOrSelector: MountEl): void;
  render(): void;
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
}

export {
  MountEl,
  RenderCell,
  RenderCellResolver,
  SpreadsheetEvents,
  ContextMenuItem,
  SpreadsheetOptions,
  ResolvedOptions,
  Spreadsheet,
};
