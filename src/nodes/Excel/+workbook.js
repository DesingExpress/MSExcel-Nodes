import { Pure } from "@design-express/fabrica";

export class workbook extends Pure {
  static path = "Office/Excel";
  static title = "Workbook";
  static description = "";

  constructor() {
    super();
    this.addInput("context", "office::excel::context");
    this.addOutput("workbook", "office::excel::workbook");
  }

  async onExecute() {
    const context = this.getInputData(1);
    console.log("workbook", context.workbook);
    const workbook = context.workbook;
    this.setOutputData(1, workbook);
  }
}
