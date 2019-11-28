const xlsl = require('xlsx');

const test_file = xlsl.readFile('./test.xlsx');

const sheets = test_file.SheetNames;

// console.log(test_file.Props);
// console.log(test_file.Workbook);

const dec = sheets[13];

console.log(dec);

const worksheet = test_file.Sheets[dec];

console.log(worksheet);
