import { Pure } from "@design-express/fabrica";

export class getSelectedRange extends Pure {
  static path = "Office/Excel/Range";
  static title = "getSelectedRange";
  static description = "";

  constructor() {
    super();
    this.addOutput("range", "office::excel::range,string");
  }

  async onExecute() {
    // console.log(this.outputs[1], this.getOutputNodes(1));
    await window.Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();
      range.load("address");
      await context.sync();
      this.setOutputData(1, range.address);
      //   console.log(`The range address was ${range.address}.`);
    });
  }
}
