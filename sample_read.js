// https://qiita.com/Kazunori-Kimura/items/29038632361fba69de5e
// https://stackoverflow.com/a/51442854
// https://qiita.com/indometacin/items/020513f7801a040dab33

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
  console.log(`Reading ${name}.`);

  // データが存在する範囲のセルをループで回す
  const sheet = book.Sheets[name];
  let range = sheet['!ref'];
  console.log(`${range}`);
  const rangeVal = utils.decode_range(range);
  console.log(`${rangeVal}`);
  for(let r = rangeVal.s.r; r <= rangeVal.e.r; r++) {
    for(let c = rangeVal.s.c; c <= rangeVal.e.c; c++) {
      const adr = utils.encode_cell({c:c, r:r});
      const cell = sheet[adr];
      if(typeof cell !== "undefined") {
        console.log(`${adr} value:${cell.v}`);
      }
    }
  }
}

