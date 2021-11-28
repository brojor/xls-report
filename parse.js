const XLSX = require('xlsx');
var workbook = XLSX.readFile('file1.xlsx');

console.log({ workbook });
