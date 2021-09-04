import { TableCell, InternalRow } from './typing';

function generateCell(): Omit<TableCell, 'id'> {
  return {};
}

function generateRow(): Omit<InternalRow, 'id' | 'cells'> {
  return {};
}

const CHAR_BASIS = 'A'.charCodeAt(0);
const BASE_MAX = 26;

function convertNumberToName(num: number): string {
  return num <= BASE_MAX
    ? String.fromCharCode(CHAR_BASIS - 1 + num)
    : convertNumberToName(~~((num - 1) / BASE_MAX)) +
        convertNumberToName(num % BASE_MAX || BASE_MAX);
}

function getColumnTitle(index: number): string {
  return convertNumberToName(index + 1);
}

function getColumnIndex(title: string): number {
  let index = -1;

  for (let i = 0; i < title.length; i++) {
    index = (index + 1) * BASE_MAX + title.charCodeAt(i) - CHAR_BASIS;
  }

  return index;
}

export { generateCell, generateRow, getColumnTitle, getColumnIndex };
