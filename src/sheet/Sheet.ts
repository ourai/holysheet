import Table, { ITable, TableInitializer } from '../table';

import { ISheet, SheetInitializer } from './typing';

class Sheet implements ISheet {
  private name: string = '';

  private table: ITable = null as any;

  constructor({ name, ...others }: SheetInitializer) {
    this.name = name;
  }

  public getName(): string {
    return this.name;
  }

  public setName(name: string): void {
    this.name = name;
  }

  public getTable(): ITable | null {
    return this.table;
  }

  public hasTable(): boolean {
    return this.table !== null;
  }

  public createTable(initializer: TableInitializer): ITable {
    if (!this.hasTable()) {
      this.table = new Table(initializer);
    }

    return this.table;
  }
}

export default Sheet;
