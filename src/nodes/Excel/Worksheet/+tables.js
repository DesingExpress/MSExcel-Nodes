import { Pure } from "@design-express/fabrica";

export class tableCollection extends Pure {
  static path = "Office/Excel/Worksheet";
  static title = "Tables";
  static description = "";

  constructor() {
    super();
    this.addInput("worksheet", "office::excel::worksheet");

    this.addOutput("tables", "office::excel::tablecollection");
  }

  onExecute() {
    const _worksheet = this.getInputData(1);
    if (!_worksheet || _worksheet.isNull) {
      return this.setOutputData(1, undefined);
    }
    this.setOutputData(1, _worksheet.tables);
  }
}
