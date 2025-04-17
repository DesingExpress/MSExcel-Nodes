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
    // this.addWidget(
    //   "toggle",
    //   "hasHeaders",
    //   this.properties.hasHeaders,
    //   "hasHeaders",
    //   {
    //     on: true,
    //     off: false,
    //   }
    // );
  }

  async onExecute() {
    const ws = this.getInputData(1);
    const data = this.getInputData(2) ?? {};
    const sr = this.getInputData(3) ?? this.properties.startRow;
    const sc = this.getInputData(4) ?? this.properties.startColumn;
    const style = this.getInputData(4) ?? "TableStyleMedium1";
    if (!ws) return;
    const { columns, rows } = data;

    let rowCount = data.length;
    let colCount = data[0]?.length ?? 0;
    let range = ws.getRangeByIndexes(sr, sc, rowCount, colCount);
    let table = ws.tables.add(range, true);
    table.getHeaderRowRange().values = columns;
    table.getDataBodyRange().values = rows;
    table.set({
      style,
    });

    this.setOutputData(1, table);
  }
}
