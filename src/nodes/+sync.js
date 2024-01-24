import { Pure } from "@design-express/fabrica";

export class sync extends Pure {
  static path = "Office";
  static title = "Sync";
  static description = "";

  constructor() {
    super();
    this.addInput("chain", "");
  }

  async onExecute() {
    await window.Excel.run(async (context) => {
      await context.sync();
    });
  }
}
