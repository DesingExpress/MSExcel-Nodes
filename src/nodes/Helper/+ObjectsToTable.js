import { Pure } from "@design-express/fabrica";

export class objectToTable extends Pure {
  static path = "Office/Helper";
  static title = "ObjectToTable";
  static description = "";

  constructor() {
    super();
    this.addInput("data", "");

    this.addOutput("table", "");
  }

  async onExecute() {
    const data = this.getInputData(1) ?? {};
    if (!data && data.length > 0) return this.setOutputData(1, undefined);
    const columns = Object.keys(data[0]);
    const rows = data.map((r) => columns.map((k) => r[k]));
    console.log(columns, rows);
    this.setOutputData(1, { columns, rows });
  }
}
