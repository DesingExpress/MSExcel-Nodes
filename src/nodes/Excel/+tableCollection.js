import { Pure } from "@design-express/fabrica";

export class tableCollection extends Pure {
  static path = "Office/Excel/TableCollection";
  static title = "TableCollection";
  static description = "";

  constructor() {
    super();
    this.addInput("context", "office::excel::context");

    this.addOutput("tables", "office::excel::tablecollection");
  }

  async onExecute() {
    const context = this.getInputData(1);
    if (!context) return this.setOutputData(1, undefined);
    const tables = context.workbook.worksheets.tables;
    this.setOutputData(1, tables);
  }
}
