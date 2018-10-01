var Excel = require("exceljs");
var fs = require("fs");

var helpingFunctions = require("./HelpingFunctions");

module.exports.getDataFromFileChild = (documentList, formulaGroupObject) => {
  //open child document based on formula
  var dirPath = helpingFunctions.getPath(); //directory path
  var file = "";
  //will return an array of objects with all the data required
  var dataToSendBack = [];

  for (var i = 0; i < documentList.length; i++) {
    var formulaPartOfName = documentList[i].substr(
      0,
      documentList[i].indexOf(" ")
    );
    if (formulaPartOfName === formulaGroupObject.formula) {
      file = `${dirPath}${documentList[i].trim()}`;
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
        dataToSendBack = [
          {
            error: `error: ${error.message} ${
              formulaGroupObject.formula
            }`
          }
        ];
        return dataToSendBack;
      }
    });
};

const getData = (workbookChild, formulaGroupObject) => {
  var inventoryWorksheet = workbookChild.getWorksheet("Inventory Master");
  var rowToFindExp = inventoryWorksheet.getRow(5);
  var expColNumber = helpingFunctions.getExpeditionsColumn(rowToFindExp);

  var arrayOfDataFromChildObjects = []; // will contain all data from this child
  var row = inventoryWorksheet.getRow(3);

  for (let rowIndex = 1 ; rowIndex<100 ;rowIndex++){
    var itemNum = row.values[rowIndex];
    if (itemNum) {
      var itemNumber = null;
      var identification = null;
      var formula = null;
      if (itemNum.toString().trim() === "?") {
        formula = formulaGroupObject.formula;
      } else {
        formulaGroupObject.itemsArray.forEach(
          (itemFromMother, colFromMotherIdex) => {
            if (
              itemFromMother.itemNumber.toString().trim() ===
              itemNum.toString().trim()
            ) {
              itemNumber = itemNum.toString().trim();
              identification = itemFromMother.identification;
              formula = itemFromMother.formula;
            }
          }
        );
      }

      if (itemNumber !== null || itemNum === "?") {
        // there is a coincidence in child
        var matchingElementCol = inventoryWorksheet.getColumn(rowIndex); //getting col for quantity
        var lotCol = inventoryWorksheet.getColumn("A"); //getting col for lot
        var expCol =
          expColNumber === 0
            ? "expedition-date column not found"
            : inventoryWorksheet.getColumn(expColNumber);
        var typeForBin = matchingElementCol.values[4];
        var lotNotTotals = '';
        for (let i = 5; i < matchingElementCol.values.length; i++) {
          if (matchingElementCol.values[i]) {
            if(lotCol.values[i]){
              lotNotTotals = lotCol.values[i].result
              ? lotCol.values[i].result.toString()
              : lotCol.values[i].toString();
              //exclude last row with total amount
              if (lotNotTotals.toLowerCase() === "totals") {
                break; // there is no more usefull data
              }
            }
            let unitNumberInStock = matchingElementCol.values[i].result;
            if (unitNumberInStock !== 0 && unitNumberInStock !== undefined) {
              
              //this is the line to get the data from
              let binLocation = "?";
              if (itemNum === "?") {
                itemNumber = "?";
                identification = "?";
              } else {
                let binWorkSheet = workbookChild.getWorksheet(
                  lotNotTotals
                );
                binLocation = helpingFunctions.getBinLocation(
                  binWorkSheet,
                  typeForBin
                );
              }

              arrayOfDataFromChildObjects.push({
                formula,
                identification,
                itemNumber,
                stockQuantity: unitNumberInStock,
                lot: lotNotTotals,
                expirationDate: expCol.values[i].result,
                binLocation,
                typeForBin
              });
              
            }
          }
        }
      }
    }
  }
  return arrayOfDataFromChildObjects;
};
