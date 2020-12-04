const fs = require('fs').promises;
const XLSX = require('./xlsx.full.min.js');

readFile("input_filevCard_In_Here/vCard_Input.vcf");     //🎈vCardi_Input.vcf - 🧪file input  //🧶🛒vCrad v2.1 v3.0 v4.0
const filename = "Contact_Output";//🎇Contact_Output.xls - file output

async function readFile(filePath) {
  try {
    const data = await fs.readFile(filePath);
    write2Excel(handle_data(data.toString()));
  } catch (error) {
    console.error(`Got an error trying to read the file: ${error.message}`);
  }
}

function handle_data(data) {//return array object [{name:xxxx, telephone:xxx},...]
  var string_data = [];
  //console.log(data);
  var begin = data.indexOf("BEGIN:VCARD");
  while (begin !== -1) {
    var FN = data.indexOf("FN:", begin);
    var be_fn = FN + 3;
    var end_fn = data.indexOf("\n", be_fn);
    var fullname = data.substring(be_fn, end_fn - 1)
    //console.log(fullname);

    var TEL = data.indexOf("TEL", end_fn);
    var be_dauhaicham = data.indexOf(":", TEL);
    var end_tel = data.indexOf("\n", be_dauhaicham);
    var telephone = data.substring(be_dauhaicham + 1, end_tel - 1);
    if (data[end_tel + 1] === 'T' && data[end_tel + 2] === 'E' && data[end_tel + 3] === 'L' ) {//TH 2 TEL
      var telephone2_be = data.indexOf(":", end_tel + 1);
      var telephone2_end = data.indexOf("\n", telephone2_be);
      var telephone2 = data.substring(telephone2_be + 1, telephone2_end - 1);
      telephone += ' | ' + telephone2;
    }
    //console.log(telephone);
    //console.log(fullname +': '+ telephone);
    string_data.push(
        {
          name: fullname,
          tel: telephone
        }
    );
    begin = data.indexOf("BEGIN:VCARD", end_tel);
  }
  return string_data;
}

function write2Excel(data_array) {
  const data = [];
  data.push(['Full Name', 'Telephone']);
  data_array.forEach(
    element => data.push([element.name, element.tel])
  );
  const book = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet(data);
  XLSX.utils.book_append_sheet(book, sheet, 'Contact_Sheet');
  XLSX.writeFile(book, `Output/${filename}.xls`);
}