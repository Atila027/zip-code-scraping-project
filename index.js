const XLSX = require('xlsx')
const Axios = require('axios').default
const FormData = require('form-data')

const BASE_URL = "https://israelpost.co.il/umbraco/Surface/Zip/FindZip"
let addressDataInfo = [];
const workbook = XLSX.readFile('addresses_data.xlsx');
const sheet_name_list = workbook.SheetNames;
const xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
// console.log(xlData.length);

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

for (let index = 1; index < 501; index++) {
  requestZipCode(xlData[3].__EMPTY,xlData[3].__EMPTY_1,index)
}