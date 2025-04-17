import { Pure } from "@design-express/fabrica";

export class getActiveWorksheet extends Pure {
  static path = "Office/Excel/WorksheetCollection";
  static title = "getActiveWorksheet";
  static description = "";

  constructor() {
    super();
    this.addInput("context", "office::excel::context");

    this.addOutput("worksheet", "office::excel::worksheet");
  }

  async onExecute() {
    const context = this.getInputData(1);
    if (!context) return this.setOutputData(1, undefined);
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    this.setOutputData(1, sheet);
  }
}
