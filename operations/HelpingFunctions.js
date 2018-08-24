////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////// FUNCTIONS FOR CHILD //////////////////////////////////////////////////

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
////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////// FUNCTIONS FOR MOTHER /////////////////////////////////////////////////

module.exports.groupItemNumbersByFormula = worksheetRead => {
  let formulaCol = worksheetRead.getColumn("A");
  let itemNumberCol = worksheetRead.getColumn("C");
  let formulaSelected = formulaCol.values[2].result;
  let itemNumbersArray = [];
  let arrayOfGroupedObjects = [];

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
  return arrayOfGroupedObjects;
};

module.exports.writeDataInMother = (results, worksheetWrite) => {
  // E stokquantity F lot G exp H bin
  worksheetWrite.columns = [
    { header: "Item Number", key: "itemNumber", width: 20 },
    { header: "Stock Quantity", key: "stockQuantity", width: 20 },
    { header: "Lot", key: "lot", width: 20 },
    { header: "Exp", key: "exp", width: 20 },
    { header: "Bin", key: "bin", width: 20 }
  ];

  results.map((resultado, index) => {
    if (resultado.length > 0) {
      //hay datos q meter
      rowNumberToWrite = 0;
      resultado.forEach(element => {
        // Add a couple of Rows by key-value, after the last current row, using the column keys

        if (element.error) {
          //ponemos el error
          let rowToWrite = worksheetWrite.addRow({ itemNumber: element.error });
          rowToWrite.commit();
        } else {
          //escribimos los datos
          let rowToWrite = worksheetWrite.addRow({
            itemNumber: element.itemNumber,
            stockQuantity: element.stockQuantity,
            lot: element.lot,
            exp: element.expirationDate,
            bin: element.binLocation
          });
          rowToWrite.commit();
        }
      });
    } else {
      //no habia datos diferentes a 0
    }
  });
};
