import { Pure } from "@design-express/fabrica";

export class getActiveWorksheet extends Pure {
  static path = "Office/Excel/WorksheetCollection";
  static title = "getActiveWorksheet";
  static description = "";

  constructor() {
    super();
    // In the Excel Office JavaScript API, context is an object that acts as the bridge between your JavaScript code and the Excel application.
    // It provides access to the Excel workbook and allows you to queue up commands to interact with the workbook’s contents.
    // 1. Access the workbook:
    //    You can access the Excel workbook through context.workbook, which then lets you manipulate worksheets, ranges, tables, charts, etc.
    // 2. Batch command execution:
    //    Office.js uses a queued command model. You issue commands (like reading or writing to cells), and those commands are queued in the context until context.sync() is called.
    // 3. Synchronization with Excel
    //    context.sync() sends all queued commands to Excel for execution and updates your JavaScript objects with the latest values from Excel.
    this.addInput("context", "office::excel::context");

    this.addOutput("worksheet", "office::excel::worksheet");
  }

  async onExecute() {
    const context = this.getInputData(1);
    if (!context) return this.setOutputData(1, undefined);
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    this.setOutputData(1, sheet);
  }
}
