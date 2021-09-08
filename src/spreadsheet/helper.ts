import { isString, noop } from '@ntks/toolbox';
import XSpreadsheet from '@wotaware/x-spreadsheet';

import {
  MountEl,
  ContextMenuItem,
  SpreadsheetOptions,
  ResolvedOptions,
  Spreadsheet,
} from './typing';

const DEFAULT_COL_COUNT = 26;
const DEFAULT_ROW_COUNT = 100;

function getDefaultContextMenuItems(inst: any): ContextMenuItem[] {
  return [
    {
      key: 'insert-column-left',
      title: () => '前面添加列',
      available: mode => mode === 'col-title',
      handler: noop,
    },
    {
      key: 'insert-column-right',
      title: () => '后面添加列',
      available: mode => mode === 'col-title',
      handler: noop,
    },
    {
      key: 'delete-selected-columns',
      title: () => '删除列',
      available: mode => mode === 'col-title' && inst.table.getColumnCount() > 1,
      handler: noop,
    },
    {
      key: 'insert-row-above',
      title: () => '上面添加行',
      available: mode => mode === 'row-title',
      handler: noop,
    },
    {
      key: 'insert-row-below',
      title: () => '下面添加行',
      available: mode => mode === 'row-title',
      handler: noop,
    },
    {
      key: 'delete-selected-rows',
      title: () => '删除行',
      available: mode => mode === 'row-title' && inst.table.getRowCount() > 1,
      handler: noop,
    },
  ];
}

function resolveOptions(
  inst: Spreadsheet,
  { column = {}, row = {}, style = {}, contextMenu, editable }: SpreadsheetOptions,
): ResolvedOptions {
  return {
    column: { count: column.count || DEFAULT_COL_COUNT, width: column.width || 100 },
    row: { count: row.count || DEFAULT_ROW_COUNT, height: row.height || 30 },
    style,
    contextMenu: contextMenu || getDefaultContextMenuItems(inst),
    editable: editable !== false,
  };
}

function createXSpreadsheetInstance(
  elementOrSelector: MountEl,
  options: ResolvedOptions,
): XSpreadsheet {
  const el = isString(elementOrSelector)
    ? document.querySelector(elementOrSelector as string)!
    : (elementOrSelector as HTMLElement);

  const { column, row } = options;

  return new XSpreadsheet(elementOrSelector, {
    mode: options.editable === true ? 'edit' : 'read',
    showToolbar: false,
    showBottomBar: false,
    view: {
      width: () => el.clientWidth,
      height: () => el.clientHeight,
    },
    col: {
      len: column.count,
      width: column.width,
      minWidth: 0,
      indexWidth: 40,
    },
    row: {
      len: row.count,
      height: row.height,
    },
    style: options.style,
    contextMenuItems: options.contextMenu,
  } as any);
}

export { resolveOptions, createXSpreadsheetInstance };
