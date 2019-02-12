var Excel = require("exceljs");
var fs = require("fs");

var helpingFunctions = require("./HelpingFunctions");

module.exports.getDataFromFileChild = (arrayOfGroupedObjects, document) => {
  //open child document based on formula
  var dirPath = helpingFunctions.getPath(); //directory path
  var file = `${dirPath}${document.trim()}`;
  var formulaPartOfName = document.substr(0, document.indexOf(" "));
  //will return an array of objects with all the data required
  var dataToSendBack = [];
  var formulaGroupObject = {};
  for (var i = 0; i < arrayOfGroupedObjects.length; i++) {
    formulaGroupObject = arrayOfGroupedObjects[i];
    if (formulaPartOfName === formulaGroupObject.formula) {
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
            error: ` there is no file for this formula to get the data from: ${
              formulaPartOfName
            }`
          }
        ];
        return dataToSendBack;
      } else {
        dataToSendBack = [
          {
            error: `error: ${error.message} ${formulaGroupObject.formula}`
          }
        ];
        return dataToSendBack;
      }
    });
};

const getData = (workbookChild, formulaGroupObject) => {
  var inventoryWorksheet = workbookChild.getWorksheet("Inventory Master");

  if (!inventoryWorksheet) {
    throw new Error(
      `${
        formulaGroupObject.formula
      } there is no Inventory Master sheet with this formula in the .xlsx file`
    );
  }

  var arrayOfDataFromChildObjects = []; // will contain all data from this child
  var row = inventoryWorksheet.getRow(3);

  for (let rowIndex = 1; rowIndex < 100; rowIndex++) {
    var itemNum = row.values[rowIndex];
    if (itemNum) {
      var itemNumber = null;
      var identification = null;
      var formula = null;
      if (itemNum.toString().trim() === "?") {
        formula = formulaGroupObject.formula;
        itemNumber = "?";
        identification = "?";
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
        var typeForBin = matchingElementCol.values[4];
        var lotNotTotals = "";
        for (let i = 5; i < matchingElementCol.values.length; i++) {
          if (matchingElementCol.values[i]) {
            if (lotCol.values[i]) {
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

              var rowToFindExp = inventoryWorksheet.getRow(i);
              var expColNumber = helpingFunctions.getExpeditionsColumn(
                rowToFindExp
              );

              let binLocation = "";
              //lotNotTotals contained in the worksheet name => open worksheet not force error
              let sheetNameForBin = "";
              for (let i = 0; i < workbookChild._worksheets.length; i++) {
                if (workbookChild._worksheets[i]) {
                  let sheet = workbookChild._worksheets[i];

                  if (
                    sheet.name.trim() === lotNotTotals.trim()
   // this was catching more than expected        || sheet.name.includes(lotNotTotals.trim())
                  ) {
                    sheetNameForBin = sheet.name;
                    break;
                  }
                }
              }
              if (sheetNameForBin === "") {
                for (let i = 0; i < workbookChild._worksheets.length; i++) {
                  if (workbookChild._worksheets[i]) {
                    let sheet = workbookChild._worksheets[i];
                    let firstPartOfLotNotTotals = lotNotTotals
                      .trim()
                      .substring(0, lotNotTotals.indexOf(" "));
                    if (
                      firstPartOfLotNotTotals.length > 4 &&
                      sheet.name.includes(firstPartOfLotNotTotals)
                    ) {
                      sheetNameForBin = sheet.name;
                      break;
                    }
                  }
                }
              }
              let binWorkSheet = workbookChild.getWorksheet(sheetNameForBin);
              if (!binWorkSheet) {
                binLocation = `${lotNotTotals} there is no corresponding worksheet for this lotNumber`;
              } else {
                binLocation = helpingFunctions.getBinLocation(
                  binWorkSheet,
                  typeForBin
                );
              }

              let expirationDate = "";
              if (expColNumber === 0) {
                expirationDate = "expedition-date column not found";
              } else {
                var expCol = inventoryWorksheet.getColumn(expColNumber);
                if (expCol.values[i]) {
                  expirationDate = expCol.values[i].result
                    ? expCol.values[i].result
                    : expCol.values[i];
                } else {
                  expirationDate = "expiration-date doesnt exist for this one";
                }
              }

              arrayOfDataFromChildObjects.push({
                formula,
                identification,
                itemNumber,
                stockQuantity: unitNumberInStock,
                lot: lotNotTotals,
                expirationDate,
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
