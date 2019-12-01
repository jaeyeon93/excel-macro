const xlsl = require('xlsx');
const test_file = xlsl.readFile('./test.xlsx');
const sheets = test_file.SheetNames;
const dec = sheets[4];

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
    if (checkUndefined(worksheet, value))
        return worksheet[value];
};

const getDescription = (worksheet, num) => {
    const value = `C${num}`;
    if (checkUndefined(worksheet, value))
        return worksheet[value];
};

const getPerson = (worksheet, num) => {
    const value = `D${num}`;
    if (checkUndefined(worksheet, value))
        return worksheet[value];
};

const getIncome = (worksheet, num) => {
    const value = `E${num}`;
    if (checkUndefined(worksheet, value))
        return worksheet[value];
};

const getOutcome = (worksheet, num) => {
    const value = `F${num}`;
    if (checkUndefined(worksheet, value))
        return worksheet[value];
};

const getReceiptNumber = (worksheet, num) => {
    const value = `J${num}`;
    if (checkUndefined(worksheet, value))
        return worksheet[value];
};

const getInputOutput = (worksheet, num) => {
    const value = `L${num}`;
    if (checkUndefined(worksheet, value))
        return worksheet[value];
};

const getPurpose = (worksheet, num) => {
    const value = `M${num}`;
    if (checkUndefined(worksheet, value))
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

const checkUndefined = (worksheet, value) => {
    if (typeof worksheet[value] != 'undefined')
        return true;
    return false;
};

const checkValidField = (worksheet, num) => {
    if (getDate(worksheet, num) == undefined && getDescription(worksheet, num) == undefined && getPerson(worksheet, num) == undefined && getIncome(worksheet, num) == undefined && getOutcome(worksheet, num) == undefined && getInputOutput(worksheet, num) == undefined && getPurpose(worksheet, num) == undefined && getReceiptNumber(worksheet, num) == undefined) {
        return true;
    }
    return false;
};

const setDate = (worksheet, num, data) => {
    worksheet[value].w = data;
    worksheet[value].v = data;
    worksheet[value].h = data;
    xlsl.writeFile(worksheet, 'hello.xlsx');
    return worksheet[value];
};

for (let i = 1; i < 91; i++)
    if (checkValidField(worksheet, i)) {
        console.log(`${i} 에서 참`);
    }

// console.log(getTotalData(worksheet, 90));
// console.log(checkValidField(worksheet, 90));
