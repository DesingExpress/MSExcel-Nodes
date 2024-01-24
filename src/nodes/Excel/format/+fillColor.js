import { Pure } from "@design-express/fabrica";

export class getSelectedRange extends Pure {
  static path = "Office/Excel/Format";
  static title = "fill";
  static description = "";

  constructor() {
    super();
    this.addInput("range", "office::excel::range,string");
    this.addInput("color", "string");
    // this.addOutput("chain", "");
  }

  async onExecute() {
    const _rng = this.getInputData(1);
    const _color = this.getInputData(2);
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const [_sheetname, _range] = _rng.split("!");
      const sheet = context.workbook.worksheets.getItem(_sheetname);
      // Read the range address
      const range = sheet.getRange(_range);
      // Update the fill color
      range.format.fill.color = _color;
      await context.sync();
    });
  }
}
