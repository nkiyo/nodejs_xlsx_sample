const xlsx = require('xlsx');
const utils = xlsx.utils;

// read an excel file
const book = xlsx.readFile('sample.xlsx');
const sheet = book.Sheets['hogesheet'];
const cell = sheet['C3'];
//const cell = sheet['c3']; // => failure
console.log(`cell is ${cell.v}`); // => "cell is aaaa"

// TODO write an excel file

