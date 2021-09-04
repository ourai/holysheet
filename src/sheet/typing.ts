import { ITable, TableInitializer } from '../table';

interface ISheet {
  getName(): string;
  setName(name: string): void;
  getTable(): ITable | null;
  hasTable(): boolean;
  createTable(initializer: TableInitializer): ITable;
}

interface SheetInitializer {
  name: string;
}

export { ISheet, SheetInitializer };
