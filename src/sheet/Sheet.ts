import AbstractTable from '../table';
import { ISheet, SheetInitializer } from './typing';

class Sheet extends AbstractTable implements ISheet {
  private name: string = '';

  constructor({ name, ...others }: SheetInitializer) {
    super(others);

    this.name = name;
  }

  public getName(): string {
    return this.name;
  }

  public setName(name: string): void {
    this.name = name;
  }
}

export default Sheet;
