import { ITable, TableInitializer } from '../table';

type SheetId = string;

interface SheetExtra {
  [key: string]: any;
}

interface SheetData extends SheetExtra {
  id?: SheetId;
  name: string;
}

interface SheetStyle {
  [key: string]: any;
}

interface ISheet {
  getId(): SheetId;
  getName(): string;
  setName(name: string): void;
  getExtra(): SheetExtra;
  setExtra(extra: SheetExtra): void;
  getTable(): ITable | null;
  hasTable(): boolean;
  createTable(initializer: TableInitializer): ITable;
  destroy(): void;
}

interface SheetInitializer {
  name: string;
}

export { SheetId, SheetExtra, SheetData, SheetStyle, ISheet, SheetInitializer };
