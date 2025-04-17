import { Pure } from "@design-express/fabrica";

export class actionRouter extends Pure {
  static path = "Office/actions";
  static title = "Router";
  static description = "";

  constructor() {
    super();
    this.addInput("tabDefinition", "object");

    this.addWidget("button", "Refresh slots", true, () =>
      this.refreshSlots(this)
    );

    this.properties = { actions: [] };
  }

  isAction(v, actions = []) {
    return actions.some(
      (e) => e.id === v.actionId && e.type === "ExecuteFunction"
    );
  }

  refreshSlots(node) {
    const ribbonJson = node.getInputData(1) ?? {};
    const { actions, tabs } = ribbonJson ?? {};
    // actions from current ribbon json
    let _actions = [];
    (tabs ?? []).forEach((t) => {
      (t?.groups ?? []).forEach((g) => {
        (g?.controls ?? []).forEach((c) => {
          if (c.type === "Button" && node.isAction(c, actions)) {
            _actions.push([c.id, -1, { label: c.label }]);
          }
          if (c.type === "Menu") {
            (c?.items ?? []).forEach((e) => {
              if (node.isAction(e, actions)) {
                _actions.push([e.id, -1, { label: e.label }]);
              }
            });
          }
        });
      });
    });
    let isChanged = !_actions.every(
      ([id, _, { label }], i) =>
        node.properties.actions[i]?.[0] === id &&
        node.properties.actions[i]?.[2].label === label
    );
    if (isChanged) {
      this.outputs.forEach((e, i) => {
        if (i > 0) this.removeOutput(i);
      });

      _actions.forEach((e, i) => {
        console.log(e);
        this.addOutput(e[0], e[1]);
        this.outputs[i + 1].label = e[2].label;
      });
      node.properties.actions = _actions;
      console.log("Refreshed!");
    }
  }

  onExecute() {
    const ribbonJson = this.getInputData(1) ?? {};
    const { actions, tabs } = ribbonJson;
    let node = this;
    actions.forEach((v) => {
      const cb = (event) => {
        let slotIdx = this.outputs.findIndex((e) => {
          return e.name === event.source.id;
        });
        if (slotIdx > 0) {
          node.triggerSlot(slotIdx);
        }
        event.completed();
      };
      window.Office.actions.associate(v.functionName, cb);
    });
  }
}
