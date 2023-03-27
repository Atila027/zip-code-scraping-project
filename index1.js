const XLSX = require('xlsx')
const Axios = require('axios').default
const fs = require('fs');
const path = require('path');
const json2csv = require('json2csv').parse;
const FormData = require('form-data')

/*--------------------Request URL--------------------*/
const BASE_URL = "https://israelpost.co.il/umbraco/Surface/Zip/FindZip"
var fields = ['No', 'City', 'Street', 'Home Number', 'Zip Code']

/*---------------------Read Address file-------------------*/
const workbook = XLSX.readFile('addresses_data.xlsx');
const sheet_name_list = workbook.SheetNames;
const xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);


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
  fs.appendFileSync(filename, rows,{ encoding: "utf8"});
  // Always add new line if file already exists.
  fs.appendFileSync(filename, "\r\n",{ encoding: "utf8"});
  console.log(`--------------------------${fileName} is generated successfully--------------`)
}

/*---------------------Main run part-------------------*/
const scrap = async (startRow, endRow) => {
  let scrapData = []
  let count = 0;
  console.log(`-------------------------scraping(${startRow} ~ ${endRow}) is started--------------------------`)
  for (let addressIndex = startRow; addressIndex < endRow; addressIndex++) {
    if(addressIndex%2 === 0 || addressIndex===endRow){
      await write(`scraping(${addressIndex-2} ~ ${addressIndex-1}).csv`, fields, scrapData);
      console.log(`--------------scraping(${addressIndex-2} ~ ${addressIndex-1}) is finished successfully----------`)
      scrapData=[]
      count = 0
    }
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
          // homeNumber += 1;
          break;
        } else {
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
        // homeNumber += 1;
        break;
      }
    }
  }
}

scrap(1,xlData.length);
