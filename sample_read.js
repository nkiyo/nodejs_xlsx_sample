// Excelファイル内の表のラベル行と特定のデータ行を取得するサンプル
// https://qiita.com/Kazunori-Kimura/items/29038632361fba69de5e
// https://stackoverflow.com/a/51442854
// https://qiita.com/indometacin/items/020513f7801a040dab33

const xlsx = require('xlsx');
const utils = xlsx.utils;


// Excelファイル内の全シートをループで回す
const xlsxFile = 'sample.xlsx';
const book = xlsx.readFile(xlsxFile);
const targetSheet = "hogesheet";
for(const name of book.SheetNames) {
  if(name !== targetSheet) {
    continue;
  }
  console.log(`Reading a sheet "${name}" in a book "${xlsxFile}".`);

  // データが存在する範囲のセルをループで回す
  const sheet = book.Sheets[name];
  let range = sheet['!ref'];
  const rangeVal = utils.decode_range(range);
  const keys = [];
  const vals = [];
  const targetLineId = "3";
  for(let r = rangeVal.s.r; r <= rangeVal.e.r; r++) {
    for(let c = rangeVal.s.c; c <= rangeVal.e.c; c++) {
      const adr = utils.encode_cell({c:c, r:r});
      const cell = sheet[adr];

      // キー文字列(2行目)のセルを保存
      if(/^[A-Z]+2$/.test(adr)) {
        keys.push(cell.v);
      }

      // 対象データセルの値を保存
      const adr2 = utils.encode_cell({c:1, r:r});
      const cell2 = sheet[adr2];
      let lineTitle = "";
      if(typeof cell2 !== "undefined") {
        lineTitle = cell2.v.toString();
      }
      if(lineTitle === targetLineId) {
        vals.push(cell.v);
      }

      // for debug
      //if(typeof cell !== "undefined") {
      //  console.log(`${adr} value:${cell.v}`);
      //}
    }
  }
  console.log(`keys are ${keys.map(k => `"${k}"`)}.`);
  console.log(`target val are ${vals.map(v => `"${v}"`)}.`);
}

