import { Pure } from "@design-express/fabrica";

export class worksheet extends Pure {
  static path = "Office/Excel";
  static title = "Worksheet";
  static description = "";

  constructor() {
    super();
    this.addInput("context", "office::excel::context");
    this.addInput("name|id", "string");

    this.addOutput("worksheet", "office::excel::worksheet");
  }

  async onExecute() {
    const context = this.getInputData(1);
    if (!context) return;

    const key = this.getInputData(2) ?? false;
    const _worksheet = key
      ? await (async function () {
          const wsCollection = context.workbook.worksheets;
          const ws = wsCollection.getItemOrNullObject(key);
          return ws;
        })()
      : context.workbook.worksheets.getActiveWorksheet();

    _worksheet.load("isNull");
    await _worksheet.context.sync();
    this.setOutputData(1, _worksheet);
  }
}
