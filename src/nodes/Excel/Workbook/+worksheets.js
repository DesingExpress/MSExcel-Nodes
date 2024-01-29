import { Pure } from "@design-express/fabrica";

export class Worksheets extends Pure {
  static path = "Office/Excel";
  static title = "Worksheets";
  static description = "";

  constructor() {
    super();
    this.addInput("workbook", "excel::workbook");
    this.addOutput("worksheets", "excel::worksheetcollection");
  }

  async onExecute() {
    const wb = this.getInputData(1);
    if (!wb) return;
    const wsc = wb.worksheets;
    this.setOutputData(1, wsc);
  }
}
