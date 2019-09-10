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
const targetSheet = "hogesheet";
for(const name of book.SheetNames) {
  if(name !== targetSheet) {
    continue;
  }
  console.log(`${name} was found.`);

  // TODO ラベル行の情報を取得
  const sheet = book.Sheets[name];
  const keys;
  console.log(`label ${sheet['B2'].v}`);

  // TODO データ行の情報を取得
  const vals;
  console.log(`label ${sheet['B3'].v}`);
}

