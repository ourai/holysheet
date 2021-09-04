import { ITable, TableInitializer } from '../table';

type SheetId = string;

interface SheetData {
  id?: SheetId;
  name: string;
}

interface ISheet {
  getId(): SheetId;
  getName(): string;
  setName(name: string): void;
  getTable(): ITable | null;
  hasTable(): boolean;
  createTable(initializer: TableInitializer): ITable;
}

interface SheetInitializer {
  name: string;
}

export { SheetId, SheetData, ISheet, SheetInitializer };
