import { Pure } from "@design-express/fabrica";

export class addTable extends Pure {
  static path = "Office/Excel/Worksheet";
  static title = "tables.add";
  static description = "";

  constructor() {
    super();
    this.addInput("worksheet", "excel::worksheet");
    this.addInput("address", "excel::range,string");
    this.addInput("style", "string");

    this.addOutput("table", "excel::table");

    this.properties = { hasHeaders: true };
    this.addWidget(
      "toggle",
      "hasHeaders",
      this.properties.hasHeaders,
      "hasHeaders",
      {
        on: true,
        off: false,
      }
    );
  }

  async onExecute() {
    const ws = this.getInputData(1);
    const address = this.getInputData(2);
    if (!ws || !address) return;
    const style = this.getInputData(3) ?? "TableStyleMedium1";

    let table = ws.tables.add(address, this.properties.hasHeaders);
    table.set({
      style,
    });

    this.setOutputData(1, table);
  }
}
