import { Pure } from "@design-express/fabrica";

export class worksheetCollection extends Pure {
  static path = "Office/Excel";
  static title = "Worksheets";
  static description = "";

  constructor() {
    super();
    this.addInput("workbook", "office::excel::workbook");

    this.addOutput("worksheets", "office::excel::worksheetcollection");
    this.addOutput("onDeactivated", -1);
    this.addOutput("deactivated", "");
  }

  dispose() {
    this.handler.remove?.();
    this.handler.context?.sync?.();
    this.handler = undefined;
  }

  handleDeactivated(event) {
    this.setOutputData(3, event);
    this.triggerSlot(2);
  }

  async onExecute() {
    const _workbook = this.getInputData(1);
    if (!_workbook) {
      console.error("Workbook is undefined.");
      return this.setOutputData(1, undefined);
    }

    const _worksheets = _workbook.worksheets;

    if (this.handler) {
      this.dispose();
    }
    this.handler = _worksheets.onDeactivated.add((event) => {
      this.setOutputData(3, event);
      this.triggerSlot(2);
    });
    await _worksheets.context.sync();

    this.setOutputData(1, _workbook.worksheets);
  }
}
