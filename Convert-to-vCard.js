const fs = require("fs").promises;
const XLSX = require("./xlsx.full.min.js");

const filename_input = "Excel_Input";
const folder_inputFileExcel = "./input_fileExcel_In_Here";
const filename_output = "Contact_vCard";

create_file_vCard(converdata2vCard(converdata2array(get_workbook_Excel(folder_inputFileExcel))));//Start->

function get_workbook_Excel(_pathFolder) {//return fist_workbook in Excel file
  try {
    var workbook = XLSX.readFile(`${_pathFolder}/${filename_input}.xlsx`);
  } catch(err1) {
    try {
      var workbook = XLSX.readFile(`${_pathFolder}/${filename_input}.xls`);
    } catch(err2){
      var err = `${err1}\nOR ${err2}\nPlease check file input and path.`;
      throw err;//Stop->outall !
    }
  }
  return workbook;
}
function converdata2array(workbook) {//return array object [{name:Nguyen Van Demo, tel1: xxxxxx, tel2: xxxxx},...]
  var first_sheet_name = workbook.SheetNames[0];//Contact Sheet must first_sheet_name
  var worksheet = workbook.Sheets[first_sheet_name];//get worksheet "Contact Sheet"

  var cell_array = [];//[0: name, 1: tel1, 2: tel2]
  var data_array_object = []; //array object [{name:Nguyen Van Demo, tel1: xxxxxx, tel2: xxxxx},...]

  var range = XLSX.utils.decode_range(worksheet['!ref']); // get the range

  for(var R = range.s.r; R <= range.e.r; ++R) {
    for(var C = range.s.c; C <= range.e.c; ++C) {
      var cell_address = {c:C, r:R};
      /* if an A1-style address is needed, encode the address */
      var cell_ref = XLSX.utils.encode_cell(cell_address);
      cell_array.push(worksheet[cell_ref].v);
    }
    //Excel maybe "number" or "string"
    data_array_object.push({name:cell_array[0].toString(), tel1: cell_array[1].toString(), tel2: cell_array[2].toString()});
    cell_array = [];
  }
  return data_array_object;
}
function converdata2vCard(data){//return string_vCard format!
  var string_vCard ='';
  data.forEach(function(item, index){
    //tel1, tel2 ->must string!
    //item: {name:Nguyen Van Demo, tel1: xxxxxx, tel2: xxxxx}
    // if(typeof item.tel1 !== "string" || typeof item.tel2 !== "string"){
    //   console.log('Not String:'+item.tel1 +'-'+ typeof item.tel1 +':'+item.name);
    //   return;
    // }
    if (item.tel1 === '' && item.tel2 === ''){
      console.log('Khong ton tai so dien thoai '+item.name);//add to data error
      return;
    }
    var letters = /^[ ()0-9+-]+$/;//✂ match(Space | ( | ) | 0->9 | + | -)
    //check tel1 tel2 is? (Space | ( | ) | 0->9 | + | -)
    if(item.tel1.match(letters)) //🧨match just match on typeof string
      if (item.tel2.match(letters) || item.tel2 === '')
        string_vCard += add_data2string(item);
      else{
        //console.log(item.name);
        return;
      }
    else{
      //console.log(item.name);
      return;
    }
  });
  return string_vCard;
}
function add_data2string(item) {//return string -vCard format
  if (item.tel2 === ''){////TH: 1 tel
    var string_data = 
`BEGIN:VCARD
VERSION:3.0
FN:${item.name}
N:;${item.name};;;
TEL;TYPE=CELL:${item.tel1}
END:VCARD`;
  }else{//TH: 2 tel
    var string_data = 
`BEGIN:VCARD
VERSION:3.0
FN:${item.name}
N:;${item.name};;;
TEL;TYPE=WORK,VOICE:${item.tel1}
TEL;TYPE=HOME,VOICE:${item.tel2}
END:VCARD`;
  }
  return string_data;
}
async function create_file_vCard(string_vCard){//write to file vCard.vcf
  try {
    await fs.writeFile(`./Output/${filename_output}.vcf`, string_vCard);
  } catch (err) {
    console.log(err);
  }
  console.log('(*^＠^*) Convert successful, please check folder Output: ' + `"${filename_output}.vcf"`);
}