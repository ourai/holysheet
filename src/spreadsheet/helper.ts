import { isString } from '@ntks/toolbox';
import XSpreadsheet, { CellStyle as XSpreadsheetCellStyle } from '@wotaware/x-spreadsheet';

import { CellStyle, Result } from '../sheet';

import {
  MountEl,
  ContextMenuItem,
  SpreadsheetOptions,
  ResolvedOptions,
  Spreadsheet,
} from './typing';

const DEFAULT_COL_COUNT = 26;
const DEFAULT_ROW_COUNT = 100;

function handleOperationResult(result: Result): void {
  if (!result.success) {
    alert(result.message);
  }
}

function getDefaultContextMenuItems(inst: any): ContextMenuItem[] {
  return [
    {
      key: 'insert-column-left',
      title: () => '前面添加列',
      available: mode => mode === 'col-title',
      handler: data => handleOperationResult(inst.insertColumn(data.selector.range.sci)),
    },
    {
      key: 'insert-column-right',
      title: () => '后面添加列',
      available: mode => mode === 'col-title',
      handler: data => handleOperationResult(inst.insertColumn(data.selector.range.sci + 1)),
    },
    {
      key: 'delete-selected-columns',
      title: () => '删除列',
      available: mode => mode === 'col-title' && inst.table.getColumnCount() > 1,
      handler: () => handleOperationResult(inst.deleteColumns()),
    },
    {
      key: 'insert-row-above',
      title: () => '上面添加行',
      available: mode => mode === 'row-title',
      handler: data => handleOperationResult(inst.insertRow(data.selector.range.sri)),
    },
    {
      key: 'insert-row-below',
      title: () => '下面添加行',
      available: mode => mode === 'row-title',
      handler: data => handleOperationResult(inst.insertRow(data.selector.range.sri + 1)),
    },
    {
      key: 'delete-selected-rows',
      title: () => '删除行',
      available: mode => mode === 'row-title' && inst.table.getRowCount() > 1,
      handler: () => handleOperationResult(inst.deleteRows()),
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

function resolveCellStyle(style: CellStyle): XSpreadsheetCellStyle {
  return {
    align: style.align,
    valign: style.verticalAlign,
    font: style.font,
    bgcolor: style.backgroundColor,
    textwrap: style.wrap,
    color: style.color,
    border: style.border,
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

export { resolveOptions, resolveCellStyle, createXSpreadsheetInstance };
