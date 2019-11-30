const xlsl = require('xlsx');

const test_file = xlsl.readFile('./test.xlsx');

const sheets = test_file.SheetNames;

// console.log(test_file.Props);
// console.log(test_file.Workbook);

const dec = sheets[13];

console.log(dec);

const worksheet = test_file.Sheets[dec];

console.log(worksheet.A1);

// { t: 's',
//   v: '삼일교회 청년 1부 현금 출납부  (2019.12월-1)',
//   f: '표지!$M$70&" 현금 출납부  ("&LEFT(표지!E19,4)&"."&L1&"-1)"',
//   h: '삼일교회 청년 1부 현금 출납부  (2019.12월-1)',
//   w: '삼일교회 청년 1부 현금 출납부  (2019.12월-1)' }
// t is
// h field는 Parse rich text and save HTML
// w field는 generate formatted text
// f : save formulae to f field
