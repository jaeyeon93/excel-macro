const xlsl = require('xlsx');
const test_file = xlsl.readFile('./test.xlsx');
const sheets = test_file.SheetNames;
const dec = sheets[13];

const worksheet = test_file.Sheets[dec];

// console.log(worksheet.A1);

// { t: 's',
//   v: '삼일교회 청년 1부 현금 출납부  (2019.12월-1)',
//   f: '표지!$M$70&" 현금 출납부  ("&LEFT(표지!E19,4)&"."&L1&"-1)"',
//   h: '삼일교회 청년 1부 현금 출납부  (2019.12월-1)',
//   w: '삼일교회 청년 1부 현금 출납부  (2019.12월-1)' }
// t is
// h field는 Parse rich text and save HTML
// w field는 generate formatted text
// f : save formulae to f field

const getDate = (worksheet, num) => {
    const value = `B${num}`;
    return worksheet[value];
};

const getDescription = (worksheet, num) => {
    const value = `B${num}`;
    return worksheet[value];
};

const getPerson = (worksheet, num) => {
    const value = `C${num}`;
    return worksheet[value];
};

const getIncome = (worksheet, num) => {
    const value = `E${num}`;
    return worksheet[value];
};

const getOutcome = (worksheet, num) => {
    const value = `F${num}`;
    return worksheet[value];
};

const getReceiptNumber = (worksheet, num) => {
    const value = `J${num}`;
    console.log(worksheet[value]);
    return worksheet[value];
};

const getInputOutput = (worksheet, num) => {
    const value = `L${num}`;
    return worksheet[value];
};

const getPurpose = (worksheet, num) => {
    const value = `M${num}`;
    return worksheet[value];
};

const getTotalData = (worksheet, num) => {
    let result = {};
    result.date = getDate(worksheet, num).w;
    result.description = getDescription(worksheet, num).w;
    result.person = getPerson(worksheet, num).w;
    result.income = getIncome(worksheet, num).w;
    result.outcome = getOutcome(worksheet, num).w;
    result.receiptNumber = getReceiptNumber(worksheet, num).w;
    result.inputOutput = getInputOutput(worksheet, num).w;
    result.purpose = getPurpose(worksheet, num).w;
    return result;
};

const setDate = (worksheet, num, data) => {
    const value = `B${num}`;
    worksheet[value].w = data;
    worksheet[value].v = data;
    worksheet[value].h = data;
    xlsl.writeFile(worksheet, 'hello.xlsx');
    return worksheet[value];
};

// const result = getTotalData(worksheet, 5);
// console.log(result);
console.log(`current`);
console.log(getDate(worksheet, 6));
console.log(`after change`);
console.log(setDate(worksheet, 6, 22));
