import { Pure } from "@design-express/fabrica";

export class getSelectedRange extends Pure {
  static path = "Office/Excel/Workbook";
  static title = "getSelectedRange";
  static description = "";

  constructor() {
    super();
    this.addInput("workbook", "office::excel::workbook");
    this.addOutput("range|address", "office::excel::range,string");
  }

  async onExecute() {
    const _workbook = this.getInputData(1);
    const _range = _workbook.getSelectedRange();
    _range.load("address");
    await _range.context.sync();
    this.setOutputData(1, _range.address);
  }
}
