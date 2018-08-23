var Excel = require("exceljs");

module.exports.getExpeditionsColumn = rowToFindExp => {
  var found = 0;
  rowToFindExp.values.forEach((val, i) => {
    var expPart = val.result.toString();
    expPart = expPart.substr(0, 3).toLowerCase();
    if (expPart === "exp" || expPart === "inv") {
      found = i;
    }
  });

  return found;
};

module.exports.getBinLocation = (binWorkSheet, typeForBin) => {
  var found = "";
  rowFindTypeForBin = binWorkSheet.getRow(5);
  rowForBinLocation = binWorkSheet.getRow(4);
  rowFindTypeForBin.values.forEach((val, i) => {
    if (val) {
      if (val.toString().trim() === typeForBin.toString().trim()) {
        found = rowForBinLocation.values[i];
      }
    }
  });
  return found;
};

module.exports.writeDataInMother = (results, motherWorksheet) => {
  // E stokquantity F lot G exp H bin
  var stockQuantityCol = motherWorksheet.getColumn("E");
  var lotCol  = motherWorksheet.getColumn('F');
  var expCol = motherWorksheet.getColumn('G');
  var binCol = motherWorksheet.getColumn('H');
  var itemNumberCol = motherWorksheet.getColumn('C');
  
  let done = results.map((resultado, index) => {
    if (resultado.length > 0) {
      //hay datos q meter
      rowNumberToWrite=0;
      resultado.forEach(element => {
        itemNumberCol.values.forEach((ItemNumVal,itemNumIndex)=>{
          if(ItemNumVal ===element.itemNumber){
            rowNumberToWrite = itemNumIndex;
          }
        });
        if (resultado[0].error) {
          //ponemos el error en todas las columnas
        } else {
          //escribimos los datos
          //MIRAR SI VA ESTO Q NO PARECE
          // ALOMEJOR hacerlo con la row directo q x columnas no parece q marche el getCelL
          stockQuantityCol.getCell(rowNumberToWrite) = element.stockQuantity
          // lot: lotCol.values[i].result,
          // expirationDate: expCol.values[i].result,
          // binLocation
  
        }
      });
    } else {
      //no habia datos diferentes a 0
    }
  });
  return 1;
};
