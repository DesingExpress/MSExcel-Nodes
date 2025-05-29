import { Pure } from "@design-express/fabrica";
import { generateIfcGuid } from "./uuid";
import {
  IFCBOOLEAN,
  IFCINTEGER,
  IFCLOGICAL,
  IFCPROPERTYSET,
  IFCPROPERTYSINGLEVALUE,
  IFCREAL,
  IFCRELDEFINESBYPROPERTIES,
  IFCTEXT,
} from "./web-ifc";

export class editManager extends Pure {
  static path = "ifc";
  static title = "EditManager";
  static description = "";

  constructor() {
    super();
    this.addInput("worksheet", "office::excel::worksheet");
    this.addInput("baseURL", "");
    this.addInput("id", "");

    this.addInput("onApply", -1);
    this.addInput("onDelete", -1);
    this.addInput("onInsertIds", -1);
    this.addInput("ids", "");

    this.addOutput("excel::table", "office::excel::table");
    this.addOutput("toApply", -1);
    this.addOutput("resp", "");
    this.addOutput("onSelection", -1);
    this.addOutput("id", "");

    this.properties = { tableName: "EditingPsetTable" };
    this.selectionHandler = undefined;
    this.actions = { DELETE: 0, CREATE: 1, UPDATE: 2 };
  }

  clear() {
    this.dataMap = new Map();
    this.origin = undefined;
    this.editType = false;
    this.isLinked = false;
    this.rowCount = 0;
    this.prevSelectedRow = undefined;
    this.toDeleteItems = [];
  }
  clearMap() {
    this.dataMap.clear();
  }

  async setTable(tableObj) {
    const worksheet = this.worksheet;
    const tableName = this.properties.tableName;
    const existedTable = worksheet.tables.getItemOrNullObject(tableName);
    await worksheet.tables.context.sync();
    if (!existedTable.isNullObject) existedTable.delete();

    const { header, data } = tableObj;
    if (!!header && !!data) {
      let colCount = header[0].length;
      let rowCount = (this.rowCount = data.length) + 1;
      const tableRange = worksheet.getRangeByIndexes(0, 0, rowCount, colCount);
      const table = worksheet.tables.add(tableRange, true);
      const headerRange = table.getHeaderRowRange();
      headerRange.numberFormat = "@";
      headerRange.values = header;
      const dataRange = table.getDataBodyRange();
      dataRange.numberFormat = "@";
      dataRange.values = data;
      tableRange.format.autofitColumns();
      table.set({
        style: "TableStyleMedium2",
      });
      table.name = tableName;

      if (this.selectionHandler) {
        this.selectionHandler.remove?.();
        this.selectionHandler.context?.sync?.();
        this.selectionHandler = undefined;
      }
      this.selectionHandler = async (event) => {
        const regex = /[A-Z]+([0-9]+)/;
        if (!header[0][0].includes("id") || !event.isInsideTable) {
          return (this.prevSelectedRow = undefined);
        }
        let rowNumber = Number(event.address.match(regex)[1]);
        if (isNaN(rowNumber) || this.prevSelectedRow === rowNumber) return;
        this.prevSelectedRow = rowNumber;
        const tableRange = await this.getTableRange();
        let id = tableRange.values?.[rowNumber - 1]?.[0];
        if (!id) return;
        this.setOutputData(5, id);
        this.triggerSlot(4);
      };
      table.onSelectionChanged.add(this.selectionHandler);

      await worksheet.context.sync();
      return table;
    }
  }

  async getTable() {
    const table = this.worksheet.tables.getItem(this.properties.tableName);
    table.columns.load("items");
    await table.context.sync();
    return table;
  }

  async getTableRange(table) {
    const _table = table ?? (await this.getTable());
    const tableRange = _table.getRange();
    tableRange.load(["values", "text", "rowCount", "columnCount"]);
    await tableRange.context.sync();
    return tableRange;
  }

  async setDeleteIds() {
    const selectedRanges = this.worksheet.context.workbook.getSelectedRanges();
    selectedRanges.load(["areas", "areaCount"]);
    this.worksheet.tables.load("items");
    await this.worksheet.context.sync();
    const table = await this.getTable();
    const tableRange = await this.getTableRange(table);
    let endRowIdx = tableRange.rowCount - 1;

    // 선택된 행들 중 테이블에 포함된 행만 추출
    const processedRowIndices = new Set(); // 중복 행 처리 방지

    let toDeleteRowIdx = [];
    for (let i = 0; i < selectedRanges.areaCount; i++) {
      const area = selectedRanges.areas.getItemAt(i);
      area.load(["rowIndex", "rowCount"]);
      await this.worksheet.context.sync();

      const areaSri = area.rowIndex;
      const areaEri = areaSri + area.rowCount - 1;

      // 이 선택 영역과 테이블 범위가 겹치는지 확인
      if (areaEri >= 1 && areaSri <= endRowIdx) {
        // 겹치는 행 인덱스 계산
        const overlapStartRow = Math.max(areaSri, 1);
        const overlapEndRow = Math.min(areaEri, endRowIdx);

        // 겹치는 각 행에 대해 처리
        for (let ri = overlapStartRow; ri <= overlapEndRow; ri++) {
          // 중복 확인 (같은 행이 여러 선택 영역에 포함될 수 있음)
          if (!processedRowIndices.has(ri)) {
            processedRowIndices.add(ri);
            // 테이블 내 행 인덱스 계산 (헤더 고려)
            let entity = this.origin.entities.find(
              (e) => e.id === tableRange.values[ri][0]
            );
            if (entity) {
              this.toDeleteItems.push(entity);
              toDeleteRowIdx.push(ri);
            }
          }
        }
      }
    }
    let minus = 0;
    toDeleteRowIdx.forEach((ri) => {
      let row = table.rows.getItemAt(ri - 1 - minus);
      row.delete();
      minus++;
    });
    await table.context.sync();
  }

  async insertIds(id) {
    const regex = /[A-Z]+([0-9]+)/;
    const ids = Array.isArray(id) ? id : [id];
    const selectedRange = this.worksheet.context.workbook.getSelectedRanges();
    selectedRange.load("address");
    await this.worksheet.context.sync();
    let address = selectedRange.address;
    let ri = Number(address.match(regex)[1]) - 1;

    const table = await this.getTable();
    const tableRange = await this.getTableRange(table);
    const newValues = tableRange.values;
    const numRow = tableRange.values.length;
    if (ri < numRow) {
      for (let i = 0; i < ids.length; i++) {
        let idx = ri + i;
        if (0 < idx) {
          if (idx >= tableRange.values.length) {
            let toAddRow = new Array(newValues[0].length).fill("");
            newValues.push(toAddRow);
            table.rows.add(null, [toAddRow]);
            newValues[idx][0] = ids[i];
          } else {
            newValues[idx][0] = ids[i];
          }
        }
      }
    }
    const currentTableRange = await this.getTableRange();
    currentTableRange.values = newValues;
    await currentTableRange.context.sync();
  }

  async applyNewPset() {
    this.clearMap();
    const logMap = { norminalValue: [], create_bnd_ids: [] };
    const adds = [];
    if (!this.isLinked) {
      return logMap;
    }
    const tableRange = await this.getTableRange();
    let values = tableRange.values;
    let column = values[0];
    let rows = values.slice(1);
    rows.forEach((r, ri) => {
      let added = Object.fromEntries(
        column
          .filter((e) => !["entity_id"].includes(e))
          .map((k) => [k, r[column.findIndex((e) => e === k)]])
      );
      let addedStr = JSON.stringify(added);
      let entity = {
        id: r[0],
        relation_id: generateIfcGuid(),
        relation_type: IFCRELDEFINESBYPROPERTIES,
      };
      const [entity_id, name] = r.slice(0, 2);
      if (!!entity_id && !!name) {
        if (!this.dataMap.has(addedStr)) {
          let index = this.dataMap.size;
          this.dataMap.set(addedStr, index);
          adds.push({
            psetAction: this.actions.CREATE,
            updateRels: [],
            createRels: [entity],
            data: {
              id: generateIfcGuid(),
              name: added["#pset_name"],
              data: Object.fromEntries(
                Object.entries(added).filter((e) => e[0] !== "#pset_name")
              ),
            },
          });
        } else {
          let index = this.dataMap.get(addedStr);
          adds[index].createRels.push(entity);
        }
      }
    });
    // Set logMap
    adds.forEach(({ createRels, data: changed }) => {
      const { name, data } = changed;
      logMap[changed.id] = [this.actions.CREATE, IFCPROPERTYSET];
      // Create edited pset
      logMap[changed.id].push("Name", 1, null, name);
      Object.entries(data).forEach(([k, v]) => {
        if (!k.includes(".type")) {
          setHasProperties(logMap, changed, k);
        }
      });
      // Create relation
      createRels.forEach((e) => {
        logMap.create_bnd_ids.push(e.relation_id);
        logMap[e.relation_id] = [
          this.actions.CREATE,
          e.relation_type,
          "RelatedObjects",
          50,
          null,
          e.id,
          "RelatingPropertyDefinition",
          50,
          null,
          changed.id,
        ];
      });
    });
    return logMap;
  }

  async applyEditedPset() {
    this.clearMap();
    const logMap = { norminalValue: [], create_bnd_ids: [] };
    if (!this.isLinked) {
      return logMap;
    }
    const tableRange = await this.getTableRange();
    let values = tableRange.values;
    let changedColumn = values[0];
    let changedRows = values.slice(1);
    // Create edits with the edited infos
    const edits = [];
    const {
      key: originMapKey,
      data: originData,
      rawData,
      entities,
    } = this.origin;

    const entityIds = entities.map((e) => e.id);
    const remainRows = changedRows.filter((r) => entityIds.includes(r[0]));
    const addedRows = changedRows.filter((r) => !entityIds.includes(r[0]));
    const toDeleteItems = [...this.toDeleteItems];
    const changedIdItems = entities.filter(
      (e) =>
        !remainRows.map((r) => r[0]).includes(e.id) &&
        !toDeleteItems.map((d) => d.id).includes(e.id)
    );
    // 삭제 및 추가가 아닌 사용자의 entity_id 직접 변경의 경우 기존 데이터 삭제
    toDeleteItems.push(...changedIdItems);

    let editedCount = 0;
    let uneditedEntities = [];
    remainRows.forEach((r, ri) => {
      let edited = Object.fromEntries(
        changedColumn
          .filter((e) => !["entity_id", "id"].includes(e))
          .map((k) => [k, r[changedColumn.findIndex((e) => e === k)]])
      );
      let editedStr = JSON.stringify(edited);
      let isChanged = editedStr !== originMapKey;
      let entity = entities.find((e) => e.id === r[0]);
      if (isChanged) {
        if (!this.dataMap.has(editedStr)) {
          let index = this.dataMap.size;
          this.dataMap.set(editedStr, index);
          edits.push({
            psetAction: this.actions.CREATE,
            updateRels: [entity],
            createRels: [],
            data: {
              id: generateIfcGuid(),
              name: edited["#pset_name"],
              data: Object.fromEntries(
                Object.entries(edited).filter((e) => e[0] !== "#pset_name")
              ),
            },
          });
        } else {
          let index = this.dataMap.get(editedStr);
          edits[index].updateRels.push(entity);
        }
        editedCount += 1;
      } else {
        uneditedEntities.push(entity);
      }
    });
    addedRows.forEach((r, ri) => {
      let added = Object.fromEntries(
        changedColumn
          .filter((e) => !["entity_id", "id"].includes(e))
          .map((k) => [k, r[changedColumn.findIndex((e) => e === k)]])
      );
      let addedStr = JSON.stringify(added);
      let entity = {
        id: r[0],
        relation_id: generateIfcGuid(),
        relation_type: IFCRELDEFINESBYPROPERTIES,
      };
      if (!this.dataMap.has(addedStr)) {
        let index = this.dataMap.size;
        this.dataMap.set(addedStr, index);
        edits.push({
          psetAction: this.actions.CREATE,
          updateRels: [],
          createRels: [entity],
          data: {
            id: generateIfcGuid(),
            name: added["#pset_name"],
            data: Object.fromEntries(
              Object.entries(added).filter((e) => e[0] !== "#pset_name")
            ),
          },
        });
      } else {
        let index = this.dataMap.get(addedStr);
        edits[index].createRels.push(entity);
      }
    });
    const isDirectDeleteRow =
      entities.length - toDeleteItems.length !==
      changedRows.length - addedRows.length;
    if (!isDirectDeleteRow && editedCount > remainRows.length / 2) {
      if (uneditedEntities.length > 0) {
        edits.push({
          psetAction: this.actions.CREATE,
          updateRels: uneditedEntities,
          createRels: [],
          data: {
            id: generateIfcGuid(),
            name: originData["#pset_name"],
            data: Object.fromEntries(
              Object.entries(originData).filter(
                (e) => !["#pset_name"].includes(e[0])
              )
            ),
          },
        });
      }
      let toUpdate = edits.find(
        (e) =>
          e.updateRels.length ===
          Math.max(...edits.map((e) => e.updateRels.length))
      );
      toUpdate.data.id = rawData.id;
      toUpdate.psetAction = this.actions.UPDATE;
    }
    console.log("remainRows", remainRows);
    console.log("addedRows", addedRows);
    console.log("deletes", toDeleteItems);
    console.log("edits", edits);

    // Set logMap
    edits.forEach(({ psetAction, updateRels, createRels, data: changed }) => {
      logMap[changed.id] = [psetAction, IFCPROPERTYSET];
      if (psetAction === this.actions.UPDATE) {
        // Modify Name
        if (rawData.name !== changed.name) {
          logMap[changed.id].push("Name", 1, null, changed.name);
        }
        // Modify(Only can create) HasProperties
        Object.entries(changed.data).forEach(([k, v]) => {
          if (!k.includes(".type")) {
            setHasProperties(logMap, changed, k, rawData);
          }
        });
      }
      if (psetAction === this.actions.CREATE) {
        // Create edited pset
        logMap[changed.id].push("Name", 1, null, changed.name);
        Object.entries(changed.data).forEach(([k, v]) => {
          if (!k.includes(".type")) {
            setHasProperties(logMap, changed, k);
          }
        });
        // Update relation
        updateRels.forEach((e) => {
          logMap[e.relation_id] = [
            this.actions.UPDATE,
            e.relation_type,
            "RelatingPropertyDefinition",
            50,
            null,
            changed.id,
          ];
        });
        // Create relation
        createRels.forEach((e) => {
          logMap.create_bnd_ids.push(e.relation_id);
          logMap[e.relation_id] = [
            this.actions.CREATE,
            e.relation_type,
            "RelatedObjects",
            50,
            null,
            e.id,
            "RelatingPropertyDefinition",
            50,
            null,
            changed.id,
          ];
        });
      }
    });
    // Delete psets
    toDeleteItems.forEach((e) => {
      logMap[e.relation_id] = [this.actions.DELETE, e.relation_type];
    });
    return logMap;
  }

  async requestPsetEntities(baseURL, id) {
    const [pset, entities] = await Promise.all([
      await fetch(new URL("ifc/1.0/search/0/propertyset", baseURL), {
        method: "post",
        body: JSON.stringify({
          conditions: `propertyset.id='${id}'`,
          getRelation: true,
        }),
        headers: {
          "content-type": "application/json",
        },
      }).then((r) => r.json()),
      await fetch(new URL("ifc/1.0/search/0/entity", baseURL), {
        method: "post",
        body: JSON.stringify({
          conditions: `propertyset.id='${id}'`,
          getRelation: true,
        }),
        headers: {
          "content-type": "application/json",
        },
      }).then((r) => r.json()),
    ]);
    return { propertyset: pset?.[0], entities };
  }

  async requestModify(baseURL, data) {
    return await fetch(new URL("ifc/1.0/modify/0", baseURL), {
      method: "post",
      body: JSON.stringify(data),
      headers: {
        "content-type": "application/json",
      },
    }).then((r) => r.json());
  }

  async onExecute() {
    this.clear();

    const ws = this.getInputData(1);
    if (!ws) return;

    let tableObj = {};
    const baseURL = this.getInputData(2);
    const id = this.getInputData(3) ?? 1;

    // Add Pset
    if (!id) {
      tableObj.header = [["entity_id", "#pset_name", "key"]];
      tableObj.data = [["", "", ""]];

      this.isLinked = true;
      this.editType = "propertyset::add";
    }
    // Edit Pset
    else {
      const resp = await this.requestPsetEntities(baseURL, id);
      // const resp = testResp;
      const { propertyset, entities } = resp;
      console.log("search response: ", resp);
      let dataKeys = Object.keys(propertyset.data).filter(
        (k) => !k.includes(".type")
      );
      const column = ["entity_id", "#pset_name", ...dataKeys];
      tableObj.header = [column];
      tableObj.data = entities.map((e) => [
        e.id,
        propertyset.name,
        ...dataKeys.map((k) => propertyset.data[k]),
      ]);

      let originPset = Object.fromEntries(
        tableObj.header[0].map((k, i) => [k, tableObj.data[0][i]])
      );
      let psetKeys = tableObj.header[0].filter(
        (e) => !["entity_id"].includes(e)
      );
      let pset = Object.fromEntries(psetKeys.map((k) => [k, originPset[k]]));
      let mapKey = JSON.stringify(pset);
      this.origin = {
        key: mapKey,
        data: pset,
        rawData: propertyset,
        entities,
      };
      // this.dataMap.set(mapKey, true);

      this.isLinked = true;
      this.editType = "propertyset::edit";
    }
    this.worksheet = ws;
    const table = this.setTable(tableObj);
    this.setOutputData(1, table);
  }

  async onAction(name) {
    if (name === "onApply") {
      let logMap;
      if (this.editType === "propertyset::add") {
        logMap = await this.applyNewPset();
      }
      if (this.editType === "propertyset::edit") {
        logMap = await this.applyEditedPset();
      }
      this.setOutputData(3, logMap);
      /** @todo fetch here!! */
      // const resp = await this.requestModify(this.getInputData(2), logMap);
      // this.setOutputData(3, resp);
      this.triggerSlot(2);
      return;
    }
    if (name === "onDelete") {
      if (this.editType === "propertyset::edit" && this.isLinked) {
        this.setDeleteIds();
      }
      return;
    }
    if (name === "onInsertIds") {
      if (this.isLinked) {
        this.insertIds(this.getInputData(7) ?? []);
      }
      return;
    }
    return super.onAction(...arguments);
  }
}

function setHasProperties(logMap, changed, key, origin = null) {
  const { norminalValue } = logMap;
  const isUpdate = !!origin;
  const id = changed.id;
  if (logMap[id].findIndex((e) => e === "HasProperties") < 0) {
    logMap[id].push("HasProperties", []);
  }
  let idx = logMap[id].findIndex((e) => e === "HasProperties");
  let type, typecode, value;
  if (isUpdate) {
    type = origin.data[`${key}.type`] ?? null;
    typecode = origin.data[`${key}.typecode`] ?? null;
    value = changed.data[key] ?? null;
  } else {
    let vInfo = getValueTypeCode(changed.data[key]);
    type = vInfo.type;
    typecode = vInfo.typecode;
    value = vInfo.value;
  }
  norminalValue.push([
    IFCPROPERTYSINGLEVALUE,
    "Name",
    1,
    null,
    key,
    "NominalValue",
    type,
    typecode,
    value,
  ]);
  logMap[id][idx + 1].push(5, norminalValue.length);
}

function getValueTypeCode(v) {
  if (v === undefined || v === null || v === "") {
    return { typecode: IFCLOGICAL, type: 3, value: null }; // IFCLOGICAL, ENUM, undefined
  }
  if (v.toLowerCase() === "true" || v.toLowerCase() === "false") {
    return {
      typecode: IFCBOOLEAN,
      type: 3,
      value: Boolean(v),
    }; // IFCBOOLEAN, ENUM, .T.||.F.
  }
  if (/^[-+]?\d*\.?\d+$/.test(v)) {
    if (v.includes(".")) {
      return { typecode: IFCREAL, type: 4, value: v }; // IFCREAL, REAL, value
    }
    return { typecode: IFCINTEGER, type: 4, value: Number(v) }; // IFCINEGER, REAL, value
  }
  return { typecode: IFCTEXT, type: 1, value: v }; // IFCTEXT, STRING, value
}
