import { isString } from '@ntks/toolbox';
import XSpreadsheet from '@wotaware/x-spreadsheet';

import { MountEl, SpreadsheetOptions } from './typing';

const DEFAULT_COL_COUNT = 26;
const DEFAULT_ROW_COUNT = 100;

function createXSpreadsheetInstance(
  elementOrSelector: MountEl,
  options: Omit<SpreadsheetOptions, 'el'>,
): XSpreadsheet {
  const el = isString(elementOrSelector)
    ? document.querySelector(elementOrSelector as string)!
    : (elementOrSelector as HTMLElement);

  const { column = {}, row = {} } = options;

  return new XSpreadsheet(elementOrSelector, {
    mode: options.editable === true ? 'edit' : 'read',
    showToolbar: false,
    showBottomBar: false,
    view: {
      width: () => el.clientWidth,
      height: () => el.clientHeight,
    },
    col: {
      len: column.count || DEFAULT_COL_COUNT,
      width: column.width || 100,
      minWidth: 0,
      indexWidth: 40,
    },
    row: {
      len: row.count || DEFAULT_ROW_COUNT,
      height: row.height || 30,
    },
    contextMenuItems: options.contextMenu || [],
  } as any);
}

export { createXSpreadsheetInstance };
