import { Pure } from "@design-express/fabrica";

export class addGeometricShape extends Pure {
  static path = "Office/Excel/ShapeCollection";
  static title = "addGeometricShape";
  static description = "";

  constructor() {
    super();
    this.addInput("worksheet", "excel::worksheet");

    this.addOutput("range", "excel::range");
  }

  async onExecute() {
    const ws = this.getInputData(1);
    if (!ws) return;

    let shapes = ws.shapes;

    // let line = shae;
  }
}
