// Minimal Office.js proxy fixture: queued writes apply only at context.sync().
const { rangeShape } = require("../../app/office/addin/operations");
const clone = (value) => JSON.parse(JSON.stringify(value));

function excelFixture(initial = [[1]]) {
  const model = { columnWidths: {}, rowHeights: {}, autoFits: [], numberFormat: initial.map((row) => row.map(() => "General")), font: { bold: false, italic: false, name: "Arial", size: 11, color: "#000000" }, fillColor: "#FFFFFF", horizontalAlignment: "General", verticalAlignment: "Bottom", wrapText: false, inserted: [], extraSheets: [], tables: [], values: clone(initial), formulas: clone(initial), syncs: 0, writes: 0, beforeSync: null };
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
          const firstRow = Number(address.match(/\d+/)[0]) - 1;
          const firstColumn = [...address.match(/^[A-Z]+/)[0]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
          const columnName = (index) => { let name = ""; for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + (n - 1) % 26) + name; return name; };
          const subrange = (row, column, rows, columns) => sheet.getRange(`${columnName(column)}${row + 1}:${columnName(column + columns - 1)}${row + rows}`);
          let columnWidth, rowHeight;
          const pendingFormat = {};
          const fields = ["horizontalAlignment", "verticalAlignment", "wrapText"];
          const format = {};
          const font = {};
          for (const key of Object.keys(model.font)) Object.defineProperty(font, key, { enumerable: true, get: () => model.font[key], set: (value) => { pendingFormat.font = { ...pendingFormat.font, [key]: value }; } });
          format.font = font;
          format.fill = { get color() { return model.fillColor; }, set color(value) { pendingFormat.fillColor = value; } };
          for (const key of fields) Object.defineProperty(format, key, { enumerable: true, get: () => model[key], set: (value) => { pendingFormat[key] = value; } });
          for (const key of ["columnWidth", "rowHeight"]) Object.defineProperty(format, key, { enumerable: true, get: () => key === "columnWidth" ? columnWidth : rowHeight, set: (value) => { pendingFormat[key] = value; } });
          for (const axis of ["Columns", "Rows"]) format[`autofit${axis}`] = () => {
            pendingStructure.push(() => {
              model.autoFits.push({ axis, address });
              model.writes++;
              // Deterministic stand-ins for Office's font-dependent layout.
              if (axis === "Columns") {
                for (let c = firstColumn; c < firstColumn + shape.columns; c++) model.columnWidths[c] = 7 * Math.max(1, ...model.values.slice(firstRow, firstRow + shape.rows).map((row) => String(row[c] ?? "").length));
              } else {
                for (let r = firstRow; r < firstRow + shape.rows; r++) model.rowHeights[r] = 15 * Math.max(1, ...(model.values[r] ?? []).slice(firstColumn, firstColumn + shape.columns).map((value) => String(value).split("\n").length));
              }
            });
          };
          const range = {
            rowIndex: Number(address.match(/\d+/)[0]) - 1,
            columnIndex: [...address.match(/^[A-Z]+/)[0]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1,
            format,
            getColumn(index) { return subrange(firstRow, firstColumn + index, shape.rows, 1); },
            getRow(index) { return subrange(firstRow + index, firstColumn, 1, shape.columns); },
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
                  if (key === "columnWidth") for (let c = firstColumn; c < firstColumn + shape.columns; c++) model.columnWidths[c] = value;
                  else if (key === "rowHeight") for (let r = firstRow; r < firstRow + shape.rows; r++) model.rowHeights[r] = value;
                  else if (key === "font") Object.assign(model.font, value);
                  else model[key] = clone(value);
                  delete pendingFormat[key];
                }
              }
              const widths = Array.from({ length: shape.columns }, (_, i) => model.columnWidths[firstColumn + i] ?? 60);
              const heights = Array.from({ length: shape.rows }, (_, i) => model.rowHeights[firstRow + i] ?? 15);
              columnWidth = widths.every((value) => value === widths[0]) ? widths[0] : null;
              rowHeight = heights.every((value) => value === heights[0]) ? heights[0] : null;
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
