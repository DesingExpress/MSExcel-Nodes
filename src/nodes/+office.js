import { Pure } from "@design-express/fabrica";

export class getHost extends Pure {
  static path = "Office";
  static title = "GetHost";
  static description = "";

  constructor() {
    super();

    this.addOutput("host", "string");
  }

  async onExecute() {
    this.setOutputData(1, window.Office?.context.host);
  }
}
