import { Pure } from "@design-express/fabrica";

export class getUsedRange extends Pure {
  static path = "Office/Excel/Worksheet";
  static title = "getUsedRange";
  static description = "";

  constructor() {
    super();
    this.addInput("worksheet", "excel::worksheet");

    this.addOutput("range", "excel::range");
  }

  async onExecute() {
    const ws = this.getInputData(1);
    if (!ws) return;

    let range = ws.getUsedRange();
    // range.select();
    // await range.context.sync();
    this.setOutputData(1, range);
  }
}
