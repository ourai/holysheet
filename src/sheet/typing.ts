import { ITable, TableInitializer } from '../table';

interface ISheet extends ITable {
  getName(): string;
  setName(name: string): void;
}

interface SheetInitializer extends TableInitializer {
  name: string;
}

export { ISheet, SheetInitializer };
