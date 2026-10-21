// Minimal Office.js proxy fixture: queued writes apply only at context.sync().
const { rangeShape } = require("../../app/office/addin/operations");
const clone = (value) => JSON.parse(JSON.stringify(value));

function excelFixture(initial = [[1]]) {
  const model = { numberFormat: initial.map((row) => row.map(() => "General")), font: { bold: false, italic: false, name: "Arial", size: 11, color: "#000000" }, fillColor: "#FFFFFF", horizontalAlignment: "General", verticalAlignment: "Bottom", wrapText: false, inserted: [], extraSheets: [], tables: [], values: clone(initial), formulas: clone(initial), syncs: 0, writes: 0, beforeSync: null };
  const excel = {
    async run(callback) {
      const ranges = [];
      const pendingStructure = [];
      const sheet = {
        id: "sheet-id", name: "Sheet One", position: 0,
        load() {},
        getUsedRangeOrNullObject() {
          return { isNullObject: model.empty === true, address: `Sheet One!A1:${String.fromCharCode(64 + model.values[0].length)}${model.values.length}`, rowIndex: 0, columnIndex: 0, rowCount: model.values.length, columnCount: model.values[0].length, load() {} };
        },
        tables: { items: [], load(options) { this.items = model.tables.slice(options.skip, options.skip + options.top).map((table) => ({ ...table, getRange: () => ({ ...table.range, load() {} }) })); } },
        getRangeByIndexes(row, column, rows, columns) {
          const whole = (axis) => ({ insert(direction) { pendingStructure.push(() => { model.inserted.push({ axis, row, column, rows, columns, direction }); model.writes++; }); } });
          return { getEntireRow: () => whole("rows"), getEntireColumn: () => whole("columns") };
        },
        getRange(address) {
          const shape = rangeShape(address);
          let values;
          let formulas;
          let text;
          let pending;
          let numberFormat;
          const pendingFormat = {};
          const fields = ["horizontalAlignment", "verticalAlignment", "wrapText"];
          const format = {};
          const font = {};
          for (const key of Object.keys(model.font)) Object.defineProperty(font, key, { enumerable: true, get: () => model.font[key], set: (value) => { pendingFormat.font = { ...pendingFormat.font, [key]: value }; } });
          format.font = font;
          format.fill = { get color() { return model.fillColor; }, set color(value) { pendingFormat.fillColor = value; } };
          for (const key of fields) Object.defineProperty(format, key, { enumerable: true, get: () => model[key], set: (value) => { pendingFormat[key] = value; } });
          const range = {
            rowIndex: Number(address.match(/\d+/)[0]) - 1,
            columnIndex: [...address.match(/^[A-Z]+/)[0]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1,
            format,
            get numberFormat() { return numberFormat; },
            set numberFormat(value) { pendingFormat.numberFormat = clone(value); },
            rowCount: shape.rows, columnCount: shape.columns,
            load() {},
            get values() { return values; },
            set values(data) { pending = { kind: "values", data: clone(data) }; },
            get formulas() { return formulas; },
            set formulas(data) { pending = { kind: "formulas", data: clone(data) }; },
            get text() { return text; },
            sync() {
              if (pending) {
                model.writes++;
                model.formulas = clone(pending.data);
                model.values = clone(pending.kind === "values" ? pending.data : pending.data.map((row) => row.map(() => 3)));
                pending = null;
              }
              if (Object.keys(pendingFormat).length) {
                model.writes++;
                for (const [key, value] of Object.entries(pendingFormat)) {
                  if (key === "font") Object.assign(model.font, value);
                  else model[key] = clone(value);
                  delete pendingFormat[key];
                }
              }
              numberFormat = clone(model.numberFormat);
              values = clone(model.values);
              formulas = clone(model.formulas);
              text = model.values.map((row) => row.map(String));
            },
          };
          ranges.push(range);
          return range;
        },
      };
      const context = {
        workbook: { worksheets: {
          items: [sheet],
          load() { this.items = [sheet, ...model.extraSheets]; },
          getItem: () => sheet,
          add(name) {
            const added = { id: `sheet-${model.extraSheets.length + 2}`, name, position: model.extraSheets.length + 1, load() {} };
            pendingStructure.push(() => { model.extraSheets.push(added); model.writes++; });
            return added;
          },
        } },
        async sync() {
          model.syncs++;
          await model.beforeSync?.(model.syncs);
          for (const work of pendingStructure.splice(0)) work();
          for (const range of ranges) range.sync();
        },
      };
      return callback(context);
    },
  };
  return { excel, model };
}

module.exports = { excelFixture };
