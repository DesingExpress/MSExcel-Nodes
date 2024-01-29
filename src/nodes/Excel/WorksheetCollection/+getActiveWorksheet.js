import { Pure } from "@design-express/fabrica";

export class getActiveWorksheet extends Pure {
  static path = "Office/Excel/WorksheetCollection";
  static title = "getActiveWorksheet";
  static description = "";

  constructor() {
    super();
    this.addInput("worksheets", "excel::worksheetcollection");
    this.addOutput("worksheet", "excel::worksheet");
  }

  async onExecute() {
    const wsc = this.getInputData(1);
    if (!wsc) return this.setOutputData(1, undefined);
    const sheet = wsc.getActiveWorksheet();
    this.setOutputData(1, sheet);

    // sheet.set({
    //   tabColor: "",
    //   name: "Sheet1",
    // });
    await sheet.context.sync();
  }
}
