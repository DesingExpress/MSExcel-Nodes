import { Pure } from "@design-express/fabrica";

export class getRangeByIndexes extends Pure {
  static path = "Office/Excel/Worksheet";
  static title = "getRangeByIndexes";
  static description = "";

  constructor() {
    super();
    this.addInput("worksheet", "excel::worksheet");
    this.addInput("startRow", "number");
    this.addInput("startColumn", "number");
    this.addInput("rowCount", "number");
    this.addInput("columnCount", "number");

    this.addOutput("range", "excel::range");
  }

  async onExecute() {
    const ws = this.getInputData(1);
    if (!ws) return;
    const startRow = this.getInputData(2);
    const startColumn = this.getInputData(3);
    const rowCount = this.getInputData(4);
    const columnCount = this.getInputData(5);

    const range = ws.getRangeByIndexes(
      startRow,
      startColumn,
      rowCount,
      columnCount
    );
    this.setOutputData(1, range);
  }
}
