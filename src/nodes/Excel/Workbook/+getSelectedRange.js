import { Pure } from "@design-express/fabrica";
import { workbook } from "../+workbook";

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
    if (!_workbook) {
      console.error("Workbook is undefined.");
      return this.setOutputData(1, undefined);
    }
    const _range = _workbook.getSelectedRange();
    this.setOutputData(1, _range);
  }
}
