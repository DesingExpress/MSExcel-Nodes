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
    if (!this.domElem) {
      const { promise, resolve } = Promise.withResolvers();
      const historyDomElem = window.document.createElement("script");
      historyDomElem.text = `window._historyCache = {
                replaceState: window.history.replaceState,
                pushState: window.history.pushState,
            };`;
      historyDomElem.type = "text/javascript";
      window.document.head.appendChild(historyDomElem);

      const domElem = (this.domElem = window.document.createElement("script"));
      domElem.src = "/addin/msoffice/dist/office.js";
      window.document.head.appendChild(domElem);
      console.log(this.domElem);
      domElem.onload = () => {
        const runtimeDomElem = window.document.createElement("script");
        runtimeDomElem.text = `const { promise, resolve, reject } = Promise.withResolvers();
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Office js deletes window.history.pushState and window.history.replaceState. Restore them
    window.history.replaceState = window._historyCache.replaceState;
    window.history.pushState = window._historyCache.pushState;
    delete window._historyCache;
    resolve();
  }
  reject();
});
window.promisifyOffice = promise;`;
        runtimeDomElem.type = "text/javascript";
        window.document.body.appendChild(runtimeDomElem);
        resolve();
      };
      await promise;
    }
    await window.promisifyOffice;
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
