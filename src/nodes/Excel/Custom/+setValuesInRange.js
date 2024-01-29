import { Pure } from "@design-express/fabrica";

export class setValuesInRange extends Pure {
  static path = "Office/Excel/Utils";
  static title = "setValuesInRange";
  static description = "";

  constructor() {
    super();
    this.addInput("worksheet", "excel::worksheet");
    this.addInput("data", "array");
    this.addInput("startRow", "number");
    this.addInput("startCol", "number");

    this.addOutput("range", "excel::range");
  }

  async onExecute() {
    const sheet = this.getInputData(1);
    const data = this.getInputData(2);
    const sr = this.getInputData(3) ?? 0;
    const sc = this.getInputData(4) ?? 0;
    if (!sheet) return;

    let isExist = !!data[0] && Array.isArray(data[0]) && data[0].length > 0;
    if (!isExist) return;

    let rowCount = data.length;
    let colCount = data[0]?.length ?? 0;
    let range = sheet.getRangeByIndexes(sr, sc, rowCount, colCount);
    range.set({
      values: data,
    });
    // await range.context.sync();
    this.setOutputData(1, range);
  }
}
