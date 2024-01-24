import { Pure } from "@design-express/fabrica";

export class custom extends Pure {
  static path = "Office/Excel";
  static title = "customMacro";
  static description = "";

  constructor() {
    super();
    this.addInput("site", "");
    this.addInput("d01warn", "");
    this.addInput("d01", "");
    this.addInput("d25warn", "");
    this.addInput("d25", "");
    this.addInput("d10warn", "");
    this.addInput("d10", "");
    this.addInput("dbwarn", "");
    this.addInput("db", "");
  }

  async onExecute() {
    const _inputs = this.inputs
      .slice(1)
      .map((d) => this.graph.links[d.link]?.data);
    // console.log(this.outputs[1], this.getOutputNodes(1));
    // console.log(_inputs.filter((_, i) => !(i % 2)));
    await window.Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const sheet = context.workbook.worksheets.getFirst();
      const usedRng = sheet.getUsedRange(true);
      const lastrowIdx = usedRng.getLastRow();
      lastrowIdx.load("rowIndex");
      await context.sync();
      const appendRng = sheet.getRangeByIndexes(
        lastrowIdx.rowIndex + 1,
        0,
        1,
        5
      );
      appendRng.values = [_inputs.filter((_, i) => !(i % 2))];
      const fillColor = _inputs.filter((_, i) => i % 2);
      for (let i = 0; i < fillColor.length; i++) {
        if (fillColor[i]) {
          appendRng.getColumn(i + 1).format.fill.color = "#dcb57c";
        }
      }
      await context.sync();

      // this.setOutputData(1, range.address);
      //   console.log(`The range address was ${range.address}.`);
    });
  }
}
