var fs = require("fs");

const MotherFileOp = require("./operations/MotherFileOp");

var dirPath = "./files/"; //directory path

fs.readdir(dirPath, function(err, documentList) {
  if (err) throw err;
  MotherFileOp.readingMotherFile(documentList);
});
