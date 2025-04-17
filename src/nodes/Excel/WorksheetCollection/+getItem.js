import { Pure } from "@design-express/fabrica";

export class getItem extends Pure {
  static path = "Office/Excel/WorksheetCollection";
  static title = "getItem";
  static description = "";

  constructor() {
    super();
    this.addInput("worksheets", "excel::worksheetcollection");
    this.addInput("key", "string");

    this.addOutput("worksheet", "excel::worksheet");
  }

  async onExecute() {
    const wsc = this.getInputData(1);
    if (!wsc) return this.setOutputData(1, undefined);
    const key = this.getInputData(2);

    const sheet = wsc.getItem(key);
    this.setOutputData(1, sheet);
  }
}
