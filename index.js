const XLSX = require('xlsx-style');
const fs = require('fs');

function build() {
  const writeOptions = { bookType: 'xlsx', bookSST: false, type: 'binary' };
  const workbook = {
    SheetNames: ['mysheet'],
    Sheets: {
      mysheet: {
        A1: { v: 6, t: 'n' },
        ["!ref"]: XLSX.utils.encode_range({
          s: { c: 0, r: 0 },
          e: { c: 1, r: 1 },
        }),
      },
    },
  };
  let data = XLSX.write(workbook, writeOptions);
  let buffer = Buffer.from(data, 'binary');
fs.writeFileSync("file.xls", buffer)
}

build();
