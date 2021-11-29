const XLSX = require('xlsx-style');
const fs = require('fs');

const fills = require('./fills');
const specs = {
  date: { width: 100 },
  size: { width: 200 },
  fee: { width: 300 },
  side: { width: '3' },
};

function build() {
  const writeOptions = { bookType: 'xlsx', bookSST: false, type: 'binary' };
  const sheet = addUsedRange(
    sheetFromArrayOfArrays(datasetToArrayofArrays(fills, specs))
  );
  const mysheet = addColumnWidths(sheet, specs);
  const workbook = {
    SheetNames: ['mysheet'],
    Sheets: {
      mysheet,
    },
  };
  let data = XLSX.write(workbook, writeOptions);
  let buffer = Buffer.from(data, 'binary');
  fs.writeFileSync('file.xls', buffer);
}

build();

function parseWidth({ width }) {
  //   console.log({ width });
  if (!width) {
    return {};
  }
  if (!Number.isInteger(parseInt(width))) {
    throw new Error('Provide column width as a number');
  }
  return typeof width === 'number' ? { wpx: width } : { wch: width };
}

function datasetToArrayofArrays(data, specification = Object.keys(data[0])) {
  return data.map((item) =>
    Object.keys(specification).map((header) => item[header])
  );
}

function sheetFromArrayOfArrays(rows) {
  return rows.reduce((sheet, row, rowIndex) => {
    row.forEach((value, colIndex) => {
      const cellRef = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
      sheet[cellRef] = { v: value, t: 's' };
    });
    return sheet;
  }, {});
}

function addUsedRange(sheet) {
  const rows = [];
  const cols = [];
  Object.keys(sheet).forEach((range) => {
    const [col, row] = [range.slice(0, 1), range.slice(1)];
    rows.push(row);
    cols.push(col);
  });
  const startCell = `${cols.sort()[0]}${rows.sort()[0]}`;
  const endCell = `${cols.sort().slice(-1)[0]}${
    rows.sort((a, b) => a - b).slice(-1)[0]
  }`;
  const sheetWithRef = Object.assign(
    { ['!ref']: `${startCell}:${endCell}` },
    sheet
  );
  return sheetWithRef;
}

function addColumnWidths(sheet, specs) {
  const columnWidths = [];
  for (column in specs) {
    columnWidths.push(parseWidth(specs[column]));
  }
  return { ...sheet, ...{ ['!cols']: columnWidths } };
}
