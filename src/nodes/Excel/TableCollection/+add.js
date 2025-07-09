import { Pure } from "@design-express/fabrica";

export class addTable extends Pure {
  static path = "Office/Excel/TableCollection";
  static title = "add";
  static description = "";

  constructor() {
    super();
    this.addInput("tables", "office::excel::tablecollection");
    this.addInput("range|address", "office::excel::range,string");
    this.addInput("", "boolean");

    this.addOutput("table", "office::excel::table");

    this.properties = {
      hasHeaders: false,
      brightness: "Light",
      color: "Black",
    };

    this.widgets_up = true;
    this.widgets_start_y = 50;
    this.addWidget(
      "combo",
      "hasHeaders",
      this.properties.hasHeaders,
      "hasHeaders",
      { values: [true, false] }
    );
    this.addWidget(
      "combo",
      "brightness",
      this.properties.brightness,
      "brightness",
      { values: ["None", "Light", "Medium", "Dark"] }
    );
    this.colors = [
      "Black",
      "Blue",
      "Orange",
      "Green",
      "Blue",
      "Purple",
      "LightGreen",
    ];
    this.addWidget("combo", "color", this.properties.color, "color", {
      values: this.colors,
    });
  }

  computeSize() {
    if (this.mode === 0) return [210, 130];
    else return [210, 155];
  }

  getStyleName() {
    const { brightness, color } = this.properties;
    if (brightness === "None") return null;
    const colorIdx = this.colors.indexOf(color);
    return `TableStyle${brightness}${colorIdx + 1}`;
  }

  async onExecute() {
    const _tables = this.getInputData(1);
    const _range = this.getInputData(2);
    if (!_tables || !_range) {
      console.error("Tables or Range is undefined.");
      this.setOutputData(1, undefined);
      return;
    }
    const hasHeader = this.getInputData(3) ?? this.properties.hasHeaders;
    const tb = _tables.add(_range, hasHeader);
    tb.set({ style: this.getStyleName() });
    await tb.context.sync();

    this.setOutputData(1, tb);
  }
}
