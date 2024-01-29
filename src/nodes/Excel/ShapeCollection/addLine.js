import { Pure } from "@design-express/fabrica";

export class addLine extends Pure {
  static path = "Office/Excel/ShapeCollection";
  static title = "addLine";
  static description = "";

  constructor() {
    super();
    this.addInput("worksheet", "excel::worksheet");

    this.addOutput("worksheet", "excel::worksheet");
  }

  async onExecute() {
    const ws = this.getInputData(1);
    if (!ws) return;

    let points = [
      [-50, -50],
      [-50, +50],
      [+50, +50],
      [+50, -50],
    ];

    let shapes = ws.shapes;
    // let line = shapes.addLine(0, 0, 100, 200, Excel.ConnectorType.straight);
    let xMin = Math.min(...points.map((e) => e[0]));
    let yMin = Math.min(...points.map((e) => e[1]));
    let xInc = xMin < 0 ? Math.abs(xMin) : 0;
    let yInc = xMin < 0 ? Math.abs(xMin) : 0;
    xInc += 10;
    yInc += 10;

    let shapeArr = [];

    for (let i = 0; i < points.length - 1; i++) {
      const [x1, y1] = points[i];
      const [x2, y2] = points[i + 1];
      console.log(x1 + xInc, y1 + yInc, x2 + xInc, y2 + yInc);
      let l = shapes.addLine(
        x1 + xInc,
        y1 + yInc,
        x2 + xInc,
        y2 + yInc,
        Excel.ConnectorType.straight
      );
      shapeArr.push(l);
    }
    let groupShape = shapes.addGroup(shapeArr);
    let toMoveCell = ws.getCell(9, 1);
    toMoveCell.load();
    await toMoveCell.context.sync();
    let { left, top } = toMoveCell;
    groupShape.set({
      left: left,
      top: top,
      height: 200,
      width: 200,
      rotation: 180,
    });

    groupShape.load();
    await shapes.context.sync();
    console.log(groupShape);

    // let line = shae;
  }
}
