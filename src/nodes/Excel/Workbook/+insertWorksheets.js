import { Pure } from "@design-express/fabrica";

export class insertWorksheets extends Pure {
  static path = "Office/Excel";
  static title = "insertWorksheets";
  static description = "";

  constructor() {
    super();
    this.addOutput("workbook", "excel::workbook");
  }

  async onExecute() {
    const input = document.createElement("input");
    input.type = "file";
    input.onchange = async (e) => {
      const file = e.target.files[0];
      const reader = new FileReader();
      reader.onload = async function (evt) {
        await window.Excel.run(async (context) => {
          // Remove the metadata before the base64-encoded string.
          let startIndex = reader.result.toString().indexOf("base64,");
          let externalWorkbook = reader.result
            .toString()
            .substr(startIndex + 7);

          // Retrieve the current workbook.
          let workbook = context.workbook;

          // Set up the insert options.
          let options = {
            sheetNamesToInsert: [], // Insert all the worksheets from the source workbook.
            positionType: window.Excel.WorksheetPositionType.beginning, // Insert after the `relativeTo` sheet.
            // relativeTo: "Sheet1", // The sheet relative to which the other worksheets will be inserted. Used with `positionType`.
          };

          // Insert the new worksheets into the current workbook.
          workbook.insertWorksheetsFromBase64(externalWorkbook, options);
          // await context.sync();
        });
        input.remove();
      };
      reader.readAsDataURL(file);
    };
    await new Promise((r) =>
      setTimeout(() => {
        input.click();
      }, 100)
    );
  }
}
