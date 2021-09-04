import { getColumnTitle, getColumnIndex } from '../abstract-table';
import { TableRange } from './typing';

function getTitleCoord(
  colIndex: number,
  rowIndex: number,
  endColIndex?: number,
  endRowIndex?: number,
): string {
  let coord: string = `${getColumnTitle(colIndex)}${rowIndex + 1}`;

  if (endColIndex !== undefined && endRowIndex !== undefined) {
    coord += `:${getTitleCoord(endColIndex, endRowIndex)}`;
  }

  return coord;
}

function getIndexCoord(titleCoord: string): TableRange {
  const range: number[] = [];

  titleCoord.split(':').forEach(title => {
    const matched = title.match(/([A-Z]+)([0-9]+)/)!;

    range.push(getColumnIndex(matched[1]), Number(matched[2]) - 1);
  });

  return range as TableRange;
}

export { getTitleCoord, getIndexCoord };
