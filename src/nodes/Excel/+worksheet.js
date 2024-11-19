import { Pure } from "@design-express/fabrica";

export class worksheet extends Pure {
  static path = "Office/Excel";
  static title = "worksheet";
  static description = "";

  constructor() {
    super();
    this.addInput("context", "office::excel::context");
    this.addInput("workbook", "office::excel::workbook");
    this.addInput("sheetname", "string");
    this.addOutput("worksheet", "office::excel::worksheet");
  }

  async onExecute() {
    const context = this.getInputData(1);
    const workbook = this.getInputData(2);
    const sheetname = this.getInputData(3);

    const worksheet = (await sheetname)
      ? (async function () {
          const sheets = workbook.worksheets;
          sheets.load("items/name");
          await context.sync();
          return sheets.find((i) => i.name === sheetname);
        })()
      : workbook.getActiveWorksheet();
    this.setOutputData(1, worksheet);
  }
}
