import { Pure } from "@design-express/fabrica";

export class setCustomCellValue extends Pure {
  static path = "Office/Excel/Utils";
  static title = "setCustomCellValue";
  static description = "";

  constructor() {
    super();
    this.addInput("workbook", "excel::workbook");
    this.addInput("targets", "array");
    this.addInput("values", "");

    this.addOutput("workbook", "excel::workbook");
  }

  async onExecute() {
    const wb = this.getInputData(1);
    const targets = this.getInputData(2);
    const values = this.getInputData(3);
    if (!wb) return;

    targets.forEach((t) => {
      const { sheet: sheetName, cell: address, key } = t;
      const v = values[key];

      let range = wb.worksheets.getItem(sheetName).getRange(address);
      range.values = [[v]];
      range.format.set({
        font: {
          color: "red",
        },
      });
    });

    this.setOutputData(1, wb);
  }
}
