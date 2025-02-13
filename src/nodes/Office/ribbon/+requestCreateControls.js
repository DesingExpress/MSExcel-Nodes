import { Pure } from "@design-express/fabrica";

export class requestCreateControls extends Pure {
  static path = "Office/ribbon";
  static title = "requestCreateControls";
  static description = "";

  constructor() {
    super({ in: true, out: false });
    this.addInput("tabDefinition", "object");

    this.addOutput("onSuccess", -1);
    this.addOutput("onError", -1);
  }

  async onExecute() {
    const ribbonJson = this.getInputData(1);
    if (!ribbonJson) return;
    await window.Office.ribbon
      .requestCreateControls(ribbonJson)
      .then((r) => {
        window.Office.ribbon.requestUpdate(ribbonJson);
        this.triggerSlot(1);
      })
      .catch((e) => {
        console.error(e);
        this.triggerSlot(2);
      });
  }
}
