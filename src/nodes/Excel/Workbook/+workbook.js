import { Pure } from "@design-express/fabrica";

export class Workbook extends Pure {
  static path = "Office/Excel";
  static title = "Workbook";
  static description = "";

  constructor() {
    super();
    this.addOutput("workbook", "excel::workbook");
  }

  async onExecute() {
    await window.Excel.run(async (context) => {
      this.setOutputData(1, context.workbook);
    });
  }
}
