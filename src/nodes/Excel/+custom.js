import { Pure } from "@design-express/fabrica";

export class custom extends Pure {
  static path = "Office/Excel";
  static title = "customMacro";
  static description = "";

  constructor() {
    super();
    this.addInput("context", "office::excel::context");
    this.addInput("values", "array");

    // this.addOutput("range", "office::excel::range,string");
  }

  async onExecute() {
    // const _inputs = this.inputs
    //   .slice(1)
    //   .map((d) => this.graph.links[d.link]?.data);
    // // console.log(this.outputs[1], this.getOutputNodes(1));
    // // console.log(_inputs.filter((_, i) => !(i % 2)));

    // const values = [1, 2, 3, 4];
    const context = this.getInputData(1);
    const values = this.getInputData(2) ?? [];
    console.log("excel", context, values);
    if (!context || values.length < 1) return;

    const sheet = context.workbook.worksheets.getFirst();
    const usedRng = sheet.getUsedRange(true);
    if (!usedRng) {
      console.log("No used range found. Starting at row 1.");
      const range = sheet.getRangeByIndexes(
        0,
        0,
        values.length,
        values[0].length
      );
      range.values = values;
      return;
    }

    usedRng.load("rowIndex, rowCount, columnCount"); // 필요한 속성 로드
    await context.sync(); // 동기화하여 데이터를 로드

    const lastRowIndex = usedRng.rowIndex + usedRng.rowCount; // 마지막 사용된 행의 다음 행 인덱스

    // 마지막 행의 다음 행부터 values의 크기만큼 데이터 삽입
    const targetRange = sheet.getRangeByIndexes(
      lastRowIndex,
      0,
      values.length,
      values[0].length
    );
    targetRange.values = values; // 2D 배열을 설정
    console.log(`2D Values written starting from row ${lastRowIndex + 1}`);
    await context.sync(); // 변경사항 적용
    // await window.Excel.run(async (context) => {
    //   /**
    //    * Insert your Excel code here
    //    */
    //   const sheet = context.workbook.worksheets.getFirst();
    //   const usedRng = sheet.getUsedRange(true);
    //   const lastrowIdx = usedRng.getLastRow();
    //   lastrowIdx.load("rowIndex");
    //   await context.sync();
    //   const appendRng = sheet.getRangeByIndexes(
    //     lastrowIdx.rowIndex + 1,
    //     0,
    //     1,
    //     5
    //   );
    //   appendRng.values = [_inputs.filter((_, i) => !(i % 2))];
    //   const fillColor = _inputs.filter((_, i) => i % 2);
    //   for (let i = 0; i < fillColor.length; i++) {
    //     if (fillColor[i]) {
    //       appendRng.getColumn(i + 1).format.fill.color = "#dcb57c";
    //     }
    //   }
    //   await context.sync();

    //   // this.setOutputData(1, range.address);
    //   //   console.log(`The range address was ${range.address}.`);
    // });
  }
}
