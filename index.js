const XLSX = require('xlsx-style');
const fs = require('fs');

const fills = require('./fills');

function build() {
  const writeOptions = { bookType: 'xlsx', bookSST: false, type: 'binary' };
  const mysheet = addUsedRange(
    sheetFromArrayOfArrays(
      dataSetToArrayofArrays(fills, ['date', 'price', 'size'])
    )
  );
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

function dataSetToArrayofArrays(data, specification = Object.keys(data[0])) {
  return data.map((item) => specification.map((header) => item[header]));
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
