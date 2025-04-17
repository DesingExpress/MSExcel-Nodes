import { Pure } from "@design-express/fabrica";

export class associateNode extends Pure {
  static path = "Office/actions";
  static title = "associate";
  static description = "";

  constructor() {
    super();
    this.addInput("functionName", "string");
    this.addInput("callback", "");
  }

  onExecute() {
    const functionName = this.getInputData(1) ?? "onExecuteRibbonAction";
    const cb = this.getInputData(2) ?? sampleAction;
    window.Office.actions.associate(functionName, cb);
  }
}

async function sampleAction(event) {
  try {
    const sourceId = event?.source?.id;
    if (sourceId === "idControl1") {
      await window.Excel.run(async (context) => {
        // Create table
        let sheet = context.workbook.worksheets.getActiveWorksheet();
        let existed = sheet.tables.getItemOrNullObject("SampleTable");
        existed.load("isNull");
        await context.sync();
        if (!existed.isNull) existed.delete();
        let table = sheet.tables.add("A1:D1", true);
        table.name = "SampleTable";
        table.getHeaderRowRange().values = [
          ["Date", "Merchant", "Category", "Amount"],
        ];
        let data = [
          ["1/1/2017", "The Phone Company", "Communications", "$120"],
          ["1/2/2017", "Northwind Electric Cars", "Transportation", "$142"],
          ["1/5/2017", "Best For You Organics Company", "Groceries", "$27"],
          ["1/10/2017", "Coho Vineyard", "Restaurant", "$33"],
          ["1/11/2017", "Bellows College", "Education", "$350"],
          ["1/15/2017", "Trey Research", "Other", "$135"],
          ["1/15/2017", "Best For You Organics Company", "Groceries", "$97"],
        ];
        table.rows.add(null, data);
        await context.sync();
      });
    }
    if (sourceId === "idControl2") {
      await window.Excel.run(async (context) => {
        // Create table
        let sheet = context.workbook.worksheets.getActiveWorksheet();
        let existed = sheet.tables.getItemOrNullObject("SampleTable");
        existed.load("isNull");
        await context.sync();
        if (!existed.isNull) existed.delete();
        await context.sync();
      });
    }
    if (sourceId === "idItem1") {
      await window.Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.getUsedRange().format.autofitColumns();
        await context.sync();
      });
    }
    if (sourceId === "idItem2") {
      await window.Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.getUsedRange().format.autofitRows();
        await context.sync();
      });
    }
    event.completed();
  } catch (e) {
    console.error(e);
    event.completed();
  }
}
