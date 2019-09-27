// Excelファイル内の表のラベル行と特定のデータ行を抽出するサンプル
// https://qiita.com/Kazunori-Kimura/items/29038632361fba69de5e
// https://stackoverflow.com/a/51442854
// https://qiita.com/indometacin/items/020513f7801a040dab33

const xlsx = require("xlsx");

const { utils } = xlsx;

// 抽出条件の指定
// -  Excelファイルパス
const xlsxFile = "sample.xlsx";
// -  Excelシート名
const targetSheet = "hogesheet";
// - 抽出したいデータ行のID(ラベル文字列)
const targetLineId = "3";

// Excelファイル内の全シートをループで回す
const book = xlsx.readFile(xlsxFile);
for (const name of book.SheetNames) {
  if (name !== targetSheet) {
    continue;
  }
  console.log(`Reading a sheet "${name}" in a book "${xlsxFile}".`);

  // データが存在する範囲のセルをループで回す
  const sheet = book.Sheets[name];
  const range = sheet["!ref"];
  const rangeVal = utils.decode_range(range);
  const keys = [];
  const vals = [];
  for (let { r } = rangeVal.s; r <= rangeVal.e.r; r++) {
    for (let { c } = rangeVal.s; c <= rangeVal.e.c; c++) {
      const adr = utils.encode_cell({ c, r });
      const cell = sheet[adr];

      // ラベル行(2行目)の値を保存
      if (/^[A-Z]+2$/.test(adr)) {
        keys.push(cell.v);
      }

      // 対象データ行の値を保存
      const adr2 = utils.encode_cell({ c: 1, r });
      const cell2 = sheet[adr2];
      let lineId = "";
      if (typeof cell2 !== "undefined") {
        lineId = cell2.v.toString();
      }
      if (lineId === targetLineId) {
        vals.push(cell.v);
      }

      // for debug
      // if(typeof cell !== "undefined") {
      //  console.log(`${adr} value:${cell.v}`);
      // }
    }
  }
  console.log(`keys are ${keys.map(k => `"${k}"`)}.`);
  console.log(`target val are ${vals.map(v => `"${v}"`)}.`);
}
