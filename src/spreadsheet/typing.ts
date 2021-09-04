type MountEl = HTMLElement | string;

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
  select(colIndex: number, rowIndex: number): void;
  select(
    startColIndex: number,
    startRowIndex: number,
    endColIndex: number,
    endRowIndex: number,
  ): void;
  mount(elementOrSelector: MountEl): void;
}

export { MountEl, ContextMenuItem, SpreadsheetOptions, ResolvedOptions, Spreadsheet };
