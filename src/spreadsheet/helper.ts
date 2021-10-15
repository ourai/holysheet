import { isString, isFunction, generateRandomId } from '@ntks/toolbox';
import XSpreadsheet, { CellStyle as XSpreadsheetCellStyle } from '@wotaware/x-spreadsheet';

import { CellStyle, Result } from '../sheet';

import {
  MountEl,
  ContextMenuItemAsserter,
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

function getDefaultContextMenuItems(): ContextMenuItem[] {
  return [
    {
      key: 'insert-column-left',
      title: '前面添加列',
      available: (_, mode) => mode === 'col-title',
      handler: (context, data) =>
        handleOperationResult(context.insertColumns(data.selector.range.sci)),
    },
    {
      key: 'insert-column-right',
      title: '后面添加列',
      available: (_, mode) => mode === 'col-title',
      handler: (context, data) =>
        handleOperationResult(context.insertColumns(data.selector.range.sci + 1)),
    },
    {
      key: 'delete-selected-columns',
      title: '删除列',
      available: (context, mode) =>
        mode === 'col-title' && (context as any).table.getColumnCount() > 1,
      handler: context => handleOperationResult(context.deleteColumns()),
    },
    {
      key: 'insert-row-above',
      title: '上面添加行',
      available: (_, mode) => mode === 'row-title',
      handler: (context, data) =>
        handleOperationResult(context.insertRows(data.selector.range.sri)),
    },
    {
      key: 'insert-row-below',
      title: '下面添加行',
      available: (_, mode) => mode === 'row-title',
      handler: (context, data) =>
        handleOperationResult(context.insertRows(data.selector.range.sri + 1)),
    },
    {
      key: 'delete-selected-rows',
      title: '删除行',
      available: (context, mode) =>
        mode === 'row-title' && (context as any).table.getRowCount() > 1,
      handler: context => handleOperationResult(context.deleteRows()),
    },
  ];
}

function resolveOptions({
  column = {},
  row = {},
  style = {},
  contextMenu,
  editable,
  hideContextMenu,
  sheetIndex,
}: SpreadsheetOptions): ResolvedOptions {
  return {
    column: { count: column.count || DEFAULT_COL_COUNT, width: column.width || 100 },
    row: { count: row.count || DEFAULT_ROW_COUNT, height: row.height || 30 },
    style,
    contextMenu: contextMenu || getDefaultContextMenuItems(),
    editable: editable !== false,
    hideContextMenu: !!hideContextMenu,
    sheetIndex: sheetIndex || 0,
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
  context: Spreadsheet,
): XSpreadsheet {
  const el = isString(elementOrSelector)
    ? document.querySelector(elementOrSelector as string)!
    : (elementOrSelector as HTMLElement);

  const { column, row } = options;

  return new XSpreadsheet(elementOrSelector, {
    mode: 'edit',
    showToolbar: false,
    showBottomBar: false,
    showContextmenu: !options.hideContextMenu,
    ...['canCut', 'canCopy', 'canPaste', 'canCellEdit'].reduce(
      (prev, key) => ({ ...prev, [key]: options.editable }),
      {},
    ),
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
    contextMenuItems: options.contextMenu.map(menuItem => ({
      key: `${generateRandomId('HolysheetContextMenuItem')}-${menuItem.key}`,
      title: () => menuItem.title,
      available: isFunction(menuItem.available)
        ? mode =>
            (menuItem.available as ContextMenuItemAsserter)(
              context,
              mode === 'range' ? 'cell' : mode,
            )
        : menuItem.available,
      handler: data => (menuItem.handler ? menuItem.handler(context, data) : undefined),
    })),
  } as any);
}

export { resolveOptions, resolveCellStyle, createXSpreadsheetInstance };
