import { Pure } from "@design-express/fabrica";

export class addTable extends Pure {
  static path = "Office/Excel/Worksheet";
  static title = "tables.add";
  static description = "";

  constructor() {
    super();
    this.addInput("worksheet", "office::excel::worksheet");
    this.addInput("data", "");
    this.addInput("", "number");
    this.addInput("", "number");

    this.addOutput("table", "office::excel::table");

    this.properties = { hasHeaders: true, startRow: 0, startColumn: 0 };

    this.widgets_up = true;
    this.widgets_start_y = 50;
    this.addWidget("number", "startRow", this.properties.startRow, "startRow");
    this.addWidget(
      "number",
      "startCol",
      this.properties.startColumn,
      "startColumn"
    );
  }

  async onExecute() {
    const ws = this.getInputData(1);
    const data = this.getInputData(2) ?? {};
    const sr = this.getInputData(3) ?? this.properties.startRow;
    const sc = this.getInputData(4) ?? this.properties.startColumn;
    const style = this.getInputData(4) ?? "TableStyleMedium2";
    if (!ws) return;
    const { columns, rows } = data;

    let rowCount = rows.length + 1;
    let colCount = columns.length ?? 0;
    let range = ws.getRangeByIndexes(sr, sc, rowCount, colCount);
    let table = ws.tables.add(range, true);

    table.getHeaderRowRange().values = [columns];
    table.getDataBodyRange().values = rows;
    table.set({
      style,
    });
    await table.context.sync();
    this.setOutputData(1, table);
  }
}
