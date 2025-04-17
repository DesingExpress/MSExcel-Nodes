import { Pure } from "@design-express/fabrica";

export class getCell extends Pure {
  static path = "Office/Excel/Worksheet";
  static title = "getCell";
  static description = "";

  constructor() {
    super();
    this.addInput("worksheet", "excel::worksheet");
    this.addInput("row", "number");
    this.addInput("column", "number");

    this.addOutput("range", "excel::range");
  }

  async onExecute() {
    const ws = this.getInputData(1);
    if (!ws) return;
    const r = this.getInputData(2) ?? 0;
    const c = this.getInputData(3) ?? 0;
    const range = ws.getCell(r, c);

    // range.format.autofitColumns();
    // range.format.columnWidth = 50;
    // range.format.rowHeight = 100;

    await range.context.sync();
    this.setOutputData(1, range);
  }
}
