import { TableRow } from '../sheet';

type MountEl = HTMLElement | string;

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
}

type ResolvedOptions = Required<Omit<SpreadsheetOptions, 'el'>>;

interface Spreadsheet {
  getSelectedRows(): TableRow[];
  select(colIndex: number, rowIndex: number): void;
  select(
    startColIndex: number,
    startRowIndex: number,
    endColIndex: number,
    endRowIndex: number,
  ): void;
  mount(elementOrSelector: MountEl): void;
}

export {
  MountEl,
  SpreadsheetEvents,
  ContextMenuItem,
  SpreadsheetOptions,
  ResolvedOptions,
  Spreadsheet,
};
