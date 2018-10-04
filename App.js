var fs = require("fs");
const MotherFileOp = require("./operations/MotherFileOp");
var helpingFunctions = require("./operations/HelpingFunctions.js");

var dirPath = helpingFunctions.getPath(); //directory path

fs.readdir(dirPath, function(err, documentList) {
  if (err) {
    throw err;
  }
  MotherFileOp.readingMotherFile(documentList);
});
