import { Pure } from "@design-express/fabrica";

export class getFirst extends Pure {
  static path = "Office/Excel/WorksheetCollection";
  static title = "getFirst";
  static description = "";

  constructor() {
    super();
    this.addInput("worksheets", "excel::worksheetcollection");
    this.addOutput("worksheet", "excel::worksheet");
  }

  async onExecute() {
    const wsc = this.getInputData(1);
    if (!wsc) return this.setOutputData(1, undefined);
    const sheet = wsc.getFirst();
    this.setOutputData(1, sheet);
  }
}
