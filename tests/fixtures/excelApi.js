// Minimal Office.js proxy fixture: queued writes apply only at context.sync().
const { rangeShape } = require("../../app/office/addin/operations");
const clone = (value) => JSON.parse(JSON.stringify(value));

function excelFixture(initial = [[1]]) {
  const model = { values: clone(initial), formulas: clone(initial), syncs: 0, writes: 0, beforeSync: null };
  const excel = {
    async run(callback) {
      const ranges = [];
      const sheet = {
        id: "sheet-id", name: "Sheet One", position: 0,
        load() {},
        getRange(address) {
          const shape = rangeShape(address);
          let values;
          let formulas;
          let text;
          let pending;
          const range = {
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
        workbook: { worksheets: { items: [sheet], load() {}, getItem: () => sheet } },
        async sync() {
          model.syncs++;
          await model.beforeSync?.(model.syncs);
          for (const range of ranges) range.sync();
        },
      };
      return callback(context);
    },
  };
  return { excel, model };
}

module.exports = { excelFixture };
