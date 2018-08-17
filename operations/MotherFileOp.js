var Excel = require("exceljs");
var mapSeries = require("async/mapSeries");
var getDataFromFileChild = require("./ChildFileOp.js").getDataFromFileChild;

const writeDataInMother = (results, worksheet) => {
  //write dadta into Mother
  console.log("finikitaun", dataObject);
};
module.exports.readingMotherFile = documentList => {
  var workbook = new Excel.Workbook();
  workbook.xlsx
    .readFile("./files/Stock Loading-Inter and FG.xlsx")
    .then(function() {
      console.log("caca de mother");
      var worksheet = workbook.getWorksheet("Stock Detailed");
      var formulaCol = worksheet.getColumn("A");
      var itemNumberCol = worksheet.getColumn("C");
      var formulaSelected = formulaCol.values[2].result;
      var itemNumbersArray = [];
      var arrayOfGroupedObjects = [];
      var dataFromChildFile = {};
      for (let i = 2; i < formulaCol.values.length; i++) {
        if (formulaCol.values[i].result !== formulaSelected) {
          //meter este grup en el array de grupos
          arrayOfGroupedObjects.push({
            formula: formulaSelected,
            itemNumbersArray
          });
          //start with the next formula
          formulaSelected = formulaCol.values[i].result;
          itemNumbersArray = [];
          itemNumbersArray.push(itemNumberCol.values[i]);
        } else {
          itemNumbersArray.push(itemNumberCol.values[i]);
        }
      }
      //hacer un mapeado async de cada uno de los blokes
      mapSeries(
        arrayOfGroupedObjects,
        (formulaGroupObject, callback) => {
          getDataFromFileChild(documentList, formulaGroupObject)
            .then(res => {
              callback(null, res);
            })
            .catch(err => {
              console.log(err);
              callback(err);
              throw err;
            });
        },
        function(err, results) {
          if (err) console.log(err); //TODO::: somenthing more??
          //results will be an array of objects
          writeDataInMother(results, worksheet);
        }
      );
    });
};
