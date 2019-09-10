// https://qiita.com/Kazunori-Kimura/items/29038632361fba69de5e
// https://stackoverflow.com/a/51442854

// TODO
// - 全行を取得
// - ヘッダ行, データ行, 空行を区別して取得

const xlsx = require('xlsx');
const utils = xlsx.utils;

const INPUT_XLSX = 'sample.xlsx';

// 全シートをループで回す
const book = xlsx.readFile(INPUT_XLSX);
const targetSheet = "s2";
for(const name of book.SheetNames) {
  console.log(`sheet name is ${name}`);
}

//const firstSheetName = book.SheetNames[0];

//const sheet = book.Sheets[firstSheetName];
//const cell = sheet['C3'];
////const cell = sheet['c3']; // => failure セル指定時のアルファベットは大文字を使用
//console.log(`read cell is ${cell.v}`);
//
//// write an excel file
//sheet['C4'].v = `てすと試験 ${Date.now()}`;
////sheet['G5'].v = 'testtest'; => failure 値未設定のセルにいきなりv属性を設定できない
//console.log(`write ${sheet['C4'].v} to xlsx`);
//xlsx.writeFile(book, OUTPUT_XLSX);

