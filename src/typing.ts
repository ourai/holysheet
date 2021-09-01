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

interface Spreadsheet {
  mount(elementOrSelector: MountEl): void;
}

export { MountEl, SpreadsheetOptions, Spreadsheet };
