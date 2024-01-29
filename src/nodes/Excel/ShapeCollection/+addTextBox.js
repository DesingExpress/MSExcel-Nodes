import { Pure } from "@design-express/fabrica";

export class addTextBox extends Pure {
  static path = "Office/Excel/ShapeCollection";
  static title = "addTextBox";
  static description = "";

  constructor() {
    super();
    this.addInput("worksheet", "excel::worksheet");
    this.addInput("text", "string");

    this.addOutput("shape", "excel::shape");
  }

  async onExecute() {
    const ws = this.getInputData(1);
    if (!ws) return;
    const text = this.getInputData(2) ?? "";
    const shapes = ws.shapes;

    let shape = shapes.addTextBox(text);

    // shape.textFrame.textRange.font.set({
    //   size: 8,
    //   bold: false,
    //   color: "red",
    // });
    // shape.textFrame.textRange.text = "asdfasdfasdf"
    // shape.textFrame.set({
    //   autoSizeSetting: "AutoSizeShapeToFitText",
    // });
    // shape.load();
    // await shapes.context.sync();

    this.setOutputData(1, shape);

    // await shapes.context.sync();

    // let line = shae;
  }
}
