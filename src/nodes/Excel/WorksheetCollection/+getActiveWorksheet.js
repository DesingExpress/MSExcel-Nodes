import { Pure } from "@design-express/fabrica";

export class getActiveWorksheet extends Pure {
  static path = "Office/Excel/WorksheetCollection";
  static title = "getActiveWorksheet";
  static description = "";

  constructor() {
    super();
    this.addInput("worksheets", "office::excel::worksheetcollection");

    this.addOutput("worksheet", "office::excel::worksheet");

    this.handler = undefined;
  }

  async onExecute() {
    const _worksheets = this.getInputData(1);
    if (!_worksheets) return this.setOutputData(1, undefined);

    const sheet = _worksheets.getActiveWorksheet();
    sheet.load(["id", "isNull"]);
    await sheet.context.sync();
    this.setOutputData(1, sheet);
  }
}
