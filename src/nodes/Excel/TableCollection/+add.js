import { Pure } from "@design-express/fabrica";

export class addTable extends Pure {
  static path = "Office/Excel/TableCollection";
  static title = "add";
  static description = "";

  constructor() {
    super();
    this.addInput("tables", "office::excel::tablecollection");
    this.addInput("range|address", "office::excel::range,string");
    this.addInput("", "boolean");

    this.addOutput("table", "office::excel::table");

    this.properties = { hasHeaders: false };
    this.addWidget(
      "combo",
      "hasHeaders",
      this.properties.hasHeaders,
      "hasHeaders",
      { values: [true, false] }
    );
    this.widgets_up = true;
    this.widgets_start_y = 50;
  }

  async onExecute() {
    const _tables = this.getInputData(1);
    const _range = this.getInputData(2);
    if (!_tables || !_range) {
      this.setOutputData(1, undefined);
      return;
    }
    const hasHeader = this.getInputData(3) ?? this.properties.hasHeaders;
    const tb = _tables.add(_range, hasHeader);
    tb.worksheet.context.sync();
    this.setOutputData(1, tb);
  }
}
