import { Pure } from "@design-express/fabrica";

export class sync extends Pure {
  static path = "Office";
  static title = "Sync";
  static description = "";

  constructor() {
    super();
    this.addInput("workbook", "excel::workbook");
    this.addInput("chain", "");
  }

  async onExecute() {
    const wb = this.getInputData(1);
    wb.context.sync();
    // await window.Excel.run(async (context) => {
    //   await wb.context.sync();
    // });
  }
}
