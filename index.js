const XLSX = require('xlsx')
const Axios = require('axios').default
const fs = require('fs');
const path = require('path');
const json2csv = require('json2csv').parse;
const FormData = require('form-data')

/*--------------------Request URL--------------------*/
const BASE_URL = "https://israelpost.co.il/umbraco/Surface/Zip/FindZip"
let addressDataInfo = [];
const fields = ['No', 'City', 'Street', 'Home Number', 'Zip Code']

/*---------------------Read Address file-------------------*/
const workbook = XLSX.readFile('addresses_data.xlsx');
const sheet_name_list = workbook.SheetNames;
const xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
// console.log(xlData.length);

/*---------------------Start Row and End Row for addresss-------------------*/
const startRow = 1121;
const endRow = 1680;

const requestZipCode = (city,street,homeNumber) =>{
  var bodyData = new FormData();
  bodyData.append('StreetID', "");
  bodyData.append('House', homeNumber);
  bodyData.append('Entry', "");
  bodyData.append('City', city);
  bodyData.append('Street', street);
  // bodyData.append('ByMaanimID', true);
  bodyData.append('__RequestVerificationToken', '8pFf-3UA0Ezy4jE1fiWq9LX9Nb6kV99KHwjcc6uMQrAfULn2KskclwybnHq4aLyn9eC6J7MLoHrix9TCn2XixKlI59C8-AEIhppmCP_1Nsk1');
  Axios({
    method  : 'POST',
    url     : BASE_URL,
    headers : bodyData.getHeaders(),
    data    : bodyData
  }).then((resolve) => {
      console.log(resolve.data);
  }).catch((error) => console.log(error));
}

const write = async (fileName, fields, data) => {
  // output file in the same folder
  const filename = path.join(__dirname,`${fileName}`);
  let rows;
  // If file doesn't exist, we will create new file and add rows with headers.    
  if (!fs.existsSync(filename)) {
      rows = json2csv(data, { header: true });
  } else {
      // Rows without headers.
      rows = json2csv(data, { header: false });
  }

  // Append file function can create new file too.
  fs.appendFileSync(filename, rows);
  // Always add new line if file already exists.
  fs.appendFileSync(filename, "\r\n");
}



const getAddressData = (xlsxData) =>{
  for (let index = 2; index < xlsxData.length; index++) {
    addressDataInfo.push({
      city:xlsxData[index].__EMPTY,
      address:xlsxData[index].__EMPTY_1
    })
  }
  return addressDataInfo;
}

getAddressData(xlData);

/*---------------------Main run part-------------------*/
const scrap = async (startRow, endRow) => {
  let scrapData = []
  let count = 0;
  for (let addressIndex = startRow; addressIndex < endRow; addressIndex++) {
    for (let homeNumber = 1; homeNumber <=500;) {
      let bodyData = new FormData();
      bodyData.append('House', homeNumber);
      bodyData.append('Entry', "");
      bodyData.append('City', xlData[addressIndex].__EMPTY);
      bodyData.append('Street', xlData[addressIndex].__EMPTY_1);
      // bodyData.append('ByMaanimID', true);
      bodyData.append('__RequestVerificationToken', '8pFf-3UA0Ezy4jE1fiWq9LX9Nb6kV99KHwjcc6uMQrAfULn2KskclwybnHq4aLyn9eC6J7MLoHrix9TCn2XixKlI59C8-AEIhppmCP_1Nsk1');
      try {
        const response = await Axios({
          method  : 'POST',
          url     : BASE_URL,
          headers : bodyData.getHeaders(),
          data    : bodyData
        });

        if(response.data.zip === '') {
          console.log("never")
          break;
        } else {
          console.log('success',response.data.zip)
          count ++;
          console.log([count, xlData[addressIndex].__EMPTY, xlData[addressIndex].__EMPTY_1, homeNumber, response.data.zip]);
          scrapData.push({
            'No' : count,
            'City' : xlData[addressIndex].__EMPTY,
            'Street' : xlData[addressIndex].__EMPTY_1,
            'Home Number' : homeNumber,
            'Zip Code': response.data.zip
          })
          homeNumber += 1;
        }
      } catch (e) {
        console.log("err")
        break;
      }
    }  
  }
  write(`result${startRow} ~ ${endRow}.csv`, fields, data);
}


scrap(1,560);
scrap(560,1120);