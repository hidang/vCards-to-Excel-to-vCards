const fs = require('fs').promises;
const XLSX = require('./xlsx.full.min.js');

readFile("input_filevCard_In_Here/vCard_Input.vcf");//🎈vCardi_Input.vcf - 🧪file input  //🧶🛒vCrad v2.1 v3.0 v4.0
const filename_output = "Contact_Excel_Output";//🎇Contact_Output.xls - file output

async function readFile(filePath) {// *vcf format: vCard v2.1 v3.0 v4.0
  try {
    const data = await fs.readFile(filePath);
    write2Excel(handle_data(data.toString()));
  } catch (error) {
    console.error(`Got an error trying to read the file: ${error.message}`);
  }
}

function handle_data(data) {//return array object [{name:xxxx, tel1:xxx, tel2:xxx },...]
  var string_data = [];
  //console.log(data);
  var begin = data.indexOf("BEGIN:VCARD");
  while (begin !== -1) {
    var FN = data.indexOf("FN:", begin);
    var be_fn = FN + 3;
    var end_fn = data.indexOf("\n", be_fn);
    var fullname = data.substring(be_fn, end_fn)
    //console.log(fullname);

    var TEL = data.indexOf("TEL", end_fn);
    var be_dauhaicham = data.indexOf(":", TEL);
    if (data[be_dauhaicham + 1] === 't' && data[be_dauhaicham + 2] === 'e' && data[be_dauhaicham + 3] === 'l'){//for vCard 4.0
      be_dauhaicham += 4;//TEL;TYPE=home,voice;VALUE=uri:tel:+1-404-555-1212
    }
    var end_tel = data.indexOf("\n", be_dauhaicham);
    var telephone = data.substring(be_dauhaicham + 1, end_tel);

    if (data[end_tel + 1] === 'T' && data[end_tel + 2] === 'E' && data[end_tel + 3] === 'L' ) {//TH 2 TEL
      var telephone2_be = data.indexOf(":", end_tel + 1);
      if (data[telephone2_be + 1] === 't' && data[telephone2_be + 2] === 'e' && data[telephone2_be + 3] === 'l'){//for vCard 4.0
        telephone2_be += 4;//TEL;TYPE=home,voice;VALUE=uri:tel:+1-404-555-1212
      }
      var telephone2_end = data.indexOf("\n", telephone2_be);
      var telephone2 = data.substring(telephone2_be + 1, telephone2_end);

    } else telephone2 = '';
    //console.log(telephone);
    //console.log(fullname +': '+ telephone);
    string_data.push(
        {
          name: fullname,
          tel1: telephone,
          tel2: telephone2
        }
    );
    begin = data.indexOf("BEGIN:VCARD", end_tel);
  }
  return string_data;
}

function write2Excel(data_array) {//input: array object [{name:xxxx, tel1:xxx, tel2:xxx },...]
  const data = [];
  data.push(['Full Name', 'Telephone 1 (Work)', 'Telephone 2 (Home)']);
  data_array.forEach(
    element => data.push([element.name, element.tel1, element.tel2])
  );
  const book = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet(data);
  XLSX.utils.book_append_sheet(book, sheet, 'Contact_Sheet');
  XLSX.writeFile(book, `Output/${filename_output}.xls`);
  console.log('(*^＠^*) Convert successful, please check folder Output: ' + `"${filename_output}.xls|.xlsx"`);
}