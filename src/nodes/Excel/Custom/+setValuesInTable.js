import { Pure } from "@design-express/fabrica";

export class setValuesInTable extends Pure {
  static path = "Office/Excel/Utils";
  static title = "setValuesInTable";
  static description = "";

  constructor() {
    super();
    this.addInput("worksheet", "excel::worksheet");
    this.addInput("data", "array");
    this.addInput("startRow", "number");
    this.addInput("startCol", "number");
    this.addInput("style", "string");

    this.addOutput("table", "excel::table");
  }

  async onExecute() {
    const ws = this.getInputData(1);
    const data = this.getInputData(2);
    const sr = this.getInputData(3) ?? 0;
    const sc = this.getInputData(4) ?? 0;
    if (!ws) return;
    const style = this.getInputData(5) ?? "TableStyleMedium2";

    let isExist = !!data[0] && Array.isArray(data[0]) && data[0].length > 0;
    if (!isExist) return;

    let rowCount = data.length;
    let colCount = data[0]?.length ?? 0;
    let range = ws.getRangeByIndexes(sr, sc, rowCount, colCount);

    // ws.tables.getItemAt(0)?.delete();

    let table = ws.tables.add(range, true);
    table.getHeaderRowRange().values = data.slice(0, 1);
    table.getDataBodyRange().values = data.slice(1);
    table.set({
      style,
    });

    this.setOutputData(1, table);
  }
}
