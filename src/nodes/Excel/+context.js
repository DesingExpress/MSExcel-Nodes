import { Pure } from "@design-express/fabrica";

export class context extends Pure {
  static path = "Office/Exel";
  static title = "Context";
  static description = "";

  constructor() {
    super();

    this.addOutput("context", "office::excel::context");
  }

  async onExecute() {
    if (this._waiter) this._waiter();
    const lock = new Promise((r) => {
      this._waiter = r;
    });
    const getResult = { current: undefined };
    const resultPromise = new Promise((r) => (getResult.current = r));
    window.Excel.run(async (context) => {
      getResult.current(context);
      await lock;
    });
    this.setOutputData(1, await resultPromise);
  }
}
