const fs = require("fs").promises;
const XLSX = require("./xlsx.full.min.js");

const filename_input = "Excel_Input";
const filename_ouput = "Contact_vCard";

try {
  var workbook = XLSX.readFile(`./input_fileExcel_In_Here/${filename_input}.xlsx`);
} catch(err1) {
  try {
    var workbook = XLSX.readFile(`./input_fileExcel_In_Here/${filename_input}.xls`);
  } catch(err2){
    console.log(`${err1}\nOR ${err2}`);
  }
}
//console.log(workbook);

var first_sheet_name = workbook.SheetNames[0];//Contact Sheet must first_sheet_name
var worksheet = workbook.Sheets[first_sheet_name];

//var desired_cell = worksheet['A1'];
//var desired_value = (desired_cell ? desired_cell.v : undefined);

var cell_array = [];//0: name, 1: tel1, 2: tel2
var data_array_object = []; //array object [{name:Nguyen Van Demo, tel1: xxxxxx, tel2: xxxxx},...]

var range = XLSX.utils.decode_range(worksheet['!ref']); // get the range
for(var R = range.s.r; R <= range.e.r; ++R) {
  for(var C = range.s.c; C <= range.e.c; ++C) {
    var cell_address = {c:C, r:R};
    /* if an A1-style address is needed, encode the address */
    var cell_ref = XLSX.utils.encode_cell(cell_address);
    cell_array.push(worksheet[cell_ref].v);
  }
  data_array_object.push({name:cell_array[0], tel1: cell_array[1], tel2: cell_array[2]});
  cell_array = [];
}
console.log(data_array_object);






// BEGIN:VCARD
// VERSION:3.0
// FN:bang
// N:;bang;;;
// TEL;TYPE=CELL:0914577757
// END:VCARD

//TH: 2 sdt
// TEL;TYPE=WORK,VOICE:(111) 555-1212
// TEL;TYPE=HOME,VOICE:(404) 555-1212

// var sheet2arr = function(sheet){
//   var result = [];
//   var row;
//   var rowNum;
//   var colNum;
//   for(rowNum = sheet['!range'].s.r; rowNum <= sheet['!range'].e.r; rowNum++){
//      row = [];
//       for(colNum=sheet['!range'].s.c; colNum<=sheet['!range'].e.c; colNum++){
//          var nextCell = sheet[
//             xlsx.utils.encode_cell({r: rowNum, c: colNum})
//          ];
//          if( typeof nextCell === 'undefined' ){
//             row.push(void 0);
//          } else row.push(nextCell.w);
//       }
//       result.push(row);
//   }
//   return result;
// };
