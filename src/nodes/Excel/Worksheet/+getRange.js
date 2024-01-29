import { Pure } from "@design-express/fabrica";

export class getRange extends Pure {
  static path = "Office/Excel/Worksheet";
  static title = "getRange";
  static description = "";

  constructor() {
    super();
    this.addInput("worksheet", "excel::worksheet");
    this.addInput("address", "string");

    this.addOutput("range", "excel::range");
  }

  async onExecute() {
    const ws = this.getInputData(1);
    if (!ws) return;
    const address = this.getInputData(2);

    const range = ws.getRange(address);
    this.setOutputData(1, range);
  }
}
