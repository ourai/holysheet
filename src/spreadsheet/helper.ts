import { isString } from '@ntks/toolbox';
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
      // handler: data => this.insertColumn(data.selector.range.sci),
    },
    {
      key: 'insert-column-right',
      title: () => '后面添加列',
      available: mode => mode === 'col-title',
      // handler: data => this.insertColumn(data.selector.range.sci + 1),
    },
    {
      key: 'delete-selected-columns',
      title: () => '删除列',
      available: mode => mode === 'col-title' && inst.sheet.getColumnCount() > 1,
      // handler: data => {
      //   const { sci, eci } = data.selector.range;

      //   if (eci - sci + 1 === this.activeTable.getColumnCount()) {
      //     this.fail(ERROR_MESSAGE.AT_LEAST_ONE_ROW_OR_COLUMN);
      //   } else if (this.activeTable.getModifiedCellsInRange().length > 0) {
      //     this.confirm('所选列已存在配置内容，确定是否继续删除？', {
      //       confirmButtonText: '删除',
      //     }).then(() => this.deleteColumns());
      //   } else {
      //     this.deleteColumns();
      //   }
      // },
    },
    {
      key: 'insert-row-above',
      title: () => '上面添加行',
      available: mode => mode === 'row-title',
      // handler: data => this.insertRow(data.selector.range.sri),
    },
    {
      key: 'insert-row-below',
      title: () => '下面添加行',
      available: mode => mode === 'row-title',
      // handler: data => this.insertRow(data.selector.range.sri + 1),
    },
    {
      key: 'delete-selected-rows',
      title: () => '删除行',
      available: mode => mode === 'row-title' && inst.sheet.getRowCount() > 1,
      // handler: data => {
      //   const { sri, eri } = data.selector.range;

      //   if (eri - sri + 1 === this.activeTable.getRowCount()) {
      //     this.fail(ERROR_MESSAGE.AT_LEAST_ONE_ROW_OR_COLUMN);
      //   } else if (this.activeTable.getModifiedCellsInRange().length > 0) {
      //     this.confirm('所选行已存在配置内容，确定是否继续删除？', {
      //       confirmButtonText: '删除',
      //     }).then(() => this.deleteRows());
      //   } else {
      //     this.deleteRows();
      //   }
      // },
    },
  ];
}

function resolveOptions(
  inst: Spreadsheet,
  { column = {}, row = {}, contextMenu, editable }: SpreadsheetOptions,
): ResolvedOptions {
  return {
    column: { count: column.count || DEFAULT_COL_COUNT, width: column.width || 100 },
    row: { count: row.count || DEFAULT_ROW_COUNT, height: row.height || 30 },
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
    contextMenuItems: options.contextMenu,
  } as any);
}

export { resolveOptions, createXSpreadsheetInstance };
