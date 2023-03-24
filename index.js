const XLSX = require('xlsx')
const Axios = require('axios').default
const fs = require('fs');
const path = require('path');
const json2csv = require('json2csv').parse;
const FormData = require('form-data')

/*--------------------Request URL--------------------*/
const BASE_URL = "https://israelpost.co.il/umbraco/Surface/Zip/FindZip"
let addressDataInfo = [];
var fields = ['No', 'City', 'Street', 'Home Number', 'Zip Code']

/*---------------------Read Address file-------------------*/
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
  console.log(`--------------------------${fileName} is generated successfully--------------`)
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
  console.log(`-------------------------scraping(${startRow} ~ ${endRow}) is started--------------------------`)
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
            "No" : count,
            "City" : xlData[addressIndex].__EMPTY,
            "Street" : xlData[addressIndex].__EMPTY_1,
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

  console.log("------------------Successfully scraped all data---------------------")
  console.log(scrapData);
  write(`result${startRow} ~ ${endRow}.csv`, fields, scrapData);
  console.log(`--------------scraping(${startRow} ~ ${endRow}) is finished successfully----------`)
}


// scrap(1,5)
scrap(23456,23460)
// scrap(1,560);
// scrap(560,1120);
// scrap(1120,1680);
// scrap(1680,2240);
// scrap(2240,2800);
// scrap(2800,3360);
// scrap(3360,3920);
// scrap(3920,4480);
// scrap(4480,5040);
// scrap(5040,5600);
// scrap(6160,6720);
// scrap(6720,7280);
// scrap(7280,7840);
// scrap(7840,8400);
// scrap(8400,8960);
// scrap(8960,9520);
// scrap(9520,10080);
// scrap(10080,10640);
// scrap(10640,11200);
// scrap(11760,12320);
// scrap(12320,12880);
// scrap(12880,13440);
// scrap(13440,14000);
// scrap(14000,14560);
// scrap(14560,15120);
// scrap(15120,15680);
// scrap(15680,16240);
// scrap(16240,16800);
// scrap(16800,17360);
// scrap(17360,17920);
// scrap(17920,18480);
// scrap(18480,19040);
// scrap(19040,19600);
// scrap(19600,20160);
// scrap(20160,20720);
// scrap(20720,21280);
// scrap(21280,21840);
// scrap(21840,22400);
// scrap(22400,22960);
// scrap(22960,23520);
// scrap(23520,24080);
// scrap(24080,24640);
// scrap(24640,25200);
// scrap(25200,25760);
// scrap(25760,26320);
// scrap(26320,26880);
// scrap(26880,27440);
// scrap(27440,28000);
// scrap(28000,28560);
// scrap(28560,29120);
// scrap(29120,29680);
// scrap(29680,30240);
// scrap(30240,30800);
// scrap(30800,31360);
// scrap(31360,31920);
// scrap(31920,32480);
// scrap(32480,33040);
// scrap(33040,33600);
// scrap(33600,34160);
// scrap(34160,34720);
// scrap(34720,35280);
// scrap(35280,35840);
// scrap(35840,36400);
// scrap(36400,36960);
// scrap(36960,37520);
// scrap(37520,38080);
// scrap(38080,38640);
// scrap(38640,39200);
// scrap(39200,39760);
// scrap(39760,40320);
// scrap(40320,40880);
// scrap(40880,41440);
// scrap(41440,42000);
// scrap(42000,42560);
// scrap(42560,43120);
// scrap(43120,43680);
// scrap(43680,44240);
// scrap(44240,44800);
// scrap(44800,45360);
// scrap(45360,45920);
// scrap(45920,46480);
// scrap(46480,47040);
// scrap(47040,47600);
// scrap(47600,48160);
// scrap(48160,48720);
// scrap(48720,49280);
// scrap(49280,49840);
// scrap(49840,50400);
// scrap(50400,50960);
// scrap(50960,51520);
// scrap(51520,52080);
// scrap(52080,52640);
// scrap(52640,53200);
// scrap(53200,53760);
// scrap(53760,54320);
// scrap(54320,54880);
// scrap(54880,55440);
// scrap(55440,55908);

