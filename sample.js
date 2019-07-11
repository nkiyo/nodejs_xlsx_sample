// https://qiita.com/Kazunori-Kimura/items/29038632361fba69de5e
// https://stackoverflow.com/a/51442854

const xlsx = require('xlsx');
const utils = xlsx.utils;

const INPUT_XLSX = 'sample.xlsx';
const OUTPUT_XLSX = 'sample2.xlsx';

// read an excel file
const book = xlsx.readFile(INPUT_XLSX);
const firstSheetName = book.SheetNames[0];
const sheet = book.Sheets[firstSheetName];
const cell = sheet['C3'];
//const cell = sheet['c3']; // => failure セル指定時のアルファベットは大文字を使用
console.log(`read cell is ${cell.v}`);

// write an excel file
sheet['C4'].v = `てすと試験 ${Date.now()}`;
//sheet['G5'].v = 'testtest'; => failure 値未設定のセルにいきなりv属性を設定できない
console.log(`write ${sheet['C4'].v} to xlsx`);
xlsx.writeFile(book, OUTPUT_XLSX);

