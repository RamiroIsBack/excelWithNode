var Excel = require("exceljs");
var fs = require("fs");

var helpingFunctions = require("./HelpingFunctions");

const getData = (workbookChild, formulaGroupObject) => {
  var inventoryWorksheet = workbookChild.getWorksheet("Inventory Master");
  var rowToFindExp = inventoryWorksheet.getRow(5);
  var expColNumber = helpingFunctions.getExpeditionsColumn(rowToFindExp);
  if (expColNumber === 0) {
    console.log("expedition column not found");
  }
  var arrayOfDataFromChildObjects = []; // will contain all data from this child
  var row = inventoryWorksheet.getRow(3);
  row.values.forEach((itemNum, rowIndex) => {
    if (itemNum) {
      var itemNumber = null;
      formulaGroupObject.itemNumbersArray.forEach(
        (itemNumFromMother, colFromMotherIdex) => {
          if (
            itemNumFromMother.toString().trim() === itemNum.toString().trim()
          ) {
            itemNumber = itemNum.toString().trim();
          }
        }
      );
      if (itemNumber !== null) {
        // there is a coincidence in child
        var matchingElementCol = inventoryWorksheet.getColumn(rowIndex); //getting col for quantity
        var lotCol = inventoryWorksheet.getColumn("A"); //getting col for lot
        var expCol = inventoryWorksheet.getColumn(expColNumber); //getting col for exp date
        var typeForBin = matchingElementCol.values[4];
        for (let i = 5; i < matchingElementCol.values.length; i++) {
          if (matchingElementCol.values[i]) {
            let unitNumberInStock = matchingElementCol.values[i].result;
            if (unitNumberInStock == !0) {
              //exclude last row with total amount
              let lotNotTotals = lotCol.values[i].result
                ? lotCol.values[i].result.toString()
                : lotCol.values[i].toString();
              if (lotNotTotals.toLowerCase() !== "totals") {
                //this is the line to get the data from
                let binWorkSheet = workbookChild.getWorksheet(
                  lotCol.values[i].result.toString()
                );
                let binLocation = helpingFunctions.getBinLocation(
                  binWorkSheet,
                  typeForBin
                );
                if (binLocation === "") {
                  console.log("bin location not found, emptyspace instead");
                }
                arrayOfDataFromChildObjects.push({
                  itemNumber, //to write data in the right place
                  stockQuantity: matchingElementCol.values[i].result,
                  lot: lotCol.values[i].result,
                  expirationDate: expCol.values[i].result,
                  binLocation
                });
              }
            }
          }
        }
      }
    }
  });
  console.log(arrayOfDataFromChildObjects);
  return arrayOfDataFromChildObjects;
};
module.exports.getDataFromFileChild = (documentList, formulaGroupObject) => {
  //open child document based on formula
  console.log("pass child");
  var dirPath = "./files/"; //directory path
  var file = "";
  //will return an array of objects with all the data required
  var dataToSendBack = [];

  for (var i = 0; i < documentList.length; i++) {
    var formulaPartOfName = documentList[i].substr(
      0,
      documentList[i].indexOf(" ")
    );
    if (formulaPartOfName === formulaGroupObject.formula) {
      file = `./files/${documentList[i].trim()}`;
      break;
    }
  }

  var workbookChild = new Excel.Workbook();
  //return workbookChild;
  return workbookChild.xlsx
    .readFile(file)
    .then(() => {
      dataToSendBack = getData(workbookChild, formulaGroupObject);
      return dataToSendBack;
    })
    .catch(error => {
      console.log(error);
      if (file === "") {
        dataToSendBack = [
          {
            error: `no se ha encontrado archivo de donde sacar datos para la formula: ${
              formulaGroupObject.formula
            }`
          }
        ];
        return dataToSendBack;
      } else {
        return error;
      }
    });
};
