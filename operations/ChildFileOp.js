var Excel = require("exceljs");
var fs = require("fs");

const getData = (workbookChild, formulaGroupObject) => {
  var Inventoryworksheet = workbookChild.getWorksheet("Inventory Master");
  var row = Inventoryworksheet.getRow(3);

  for (ItemNumber in formulaGroupObject.itemNumbersArray) {
    console.log(ItemNumber);
  }

  return {
    caca: {
      delavaca: {
        fresca: true
      }
    }
  };
  // dentro del fichero en la tab Inventory Master
  // for(let i = 0; i<itemsNumbers.length; i++){
  // if(row[3]=== itemNumbersArray[i]
  //  devolvera un objeto con todo lo relevante a ese fichero para luego guardarlo en la madre
  //}
};
module.exports.getDataFromFileChild = (documentList, formulaGroupObject) => {
  //open child document based on formula
  console.log("caca de child");
  var dirPath = "./files/"; //directory path
  var file = "";

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
  return workbookChild.xlsx
    .readFile(file)
    .then(() => {
      //will return an object with all the data required
      dataToSendBack = {};
      dataToSendBack = getData(workbookChild, formulaGroupObject);
      return dataToSendBack;
    })
    .catch(error => {
      console.log(error);
      throw error;
    });
};
