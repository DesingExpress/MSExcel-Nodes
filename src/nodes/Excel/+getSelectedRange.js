import { Pure } from "@design-express/fabrica";

export class getSelectedRange extends Pure {
  static path = "Office/Excel/Range";
  static title = "getSelectedRange";
  static description = "";

  constructor() {
    super();
    this.addInput("context", "office::excel::context");
    this.addOutput("range", "office::excel::range,string");
  }

  async onExecute() {
    const context = this.getInputData(1);
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    this.setOutputData(1, range.address);
  }
}
