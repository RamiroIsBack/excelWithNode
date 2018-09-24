var Excel = require("exceljs");
var mapSeries = require("async/mapSeries");
var getDataFromFileChild = require("./ChildFileOp.js").getDataFromFileChild;
var helpingFunctions = require("./HelpingFunctions.js");

module.exports.readingMotherFile = documentList => {
  var workbookRead = new Excel.Workbook();
  console.log("... start data processing ...");
  var dirPath = helpingFunctions.getPath(); //directory path

  workbookRead.xlsx
    .readFile(`${dirPath}Stock Loading-Inter and FG.xlsx`)
    .then(function() {
      var worksheetRead = workbookRead.getWorksheet("Stock Detailed");
      var arrayOfGroupedObjects = helpingFunctions.groupItemNumbersByFormula(
        worksheetRead
      );

      console.log("... creating document ...");
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
          writeData(results, dirPath);
        }
      );
    })
    .catch(error => {
      console.log(error);
    });
};

var writeData = (results, dirPath) => {
  var workbookWrite = new Excel.Workbook();
  // create new sheet with pageSetup settings for A4 - landscape
  var worksheetWrite = workbookWrite.addWorksheet("sheet", {
    pageSetup: { paperSize: 9, orientation: "landscape" }
  });
  helpingFunctions.writeDataInMother(results, worksheetWrite);
  let resultFile = `./results Stock L-I and FG.xlsx`;
  workbookWrite.xlsx.writeFile(resultFile).then(function() {
    console.log(`... done ...
      result file: ${resultFile}`);
  });
};
