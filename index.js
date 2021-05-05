const excelToJSON = require("convert-excel-to-json");

// Get  the workbook
const workbook = excelToJSON({
  sourceFile: "sheets.xlsx",
});

/**
 * In the workbook we have worksheets which are arrays of objects.
 * Each worksheet has an object which has destination information.
 * In each worksheet from the 5th column,  there's actual trip data.
 *
 */

console.log(`workbook`, Object.values(workbook).length);

const worksheets = [];

for (let sheet of Object.values(workbook)) {
  if (sheet.length != 0) {
    worksheets.push(sheet);
  }
}

// console.log(`worksheets`, worksheets);

worksheets.forEach((sheet) => {
  // console.log(`sheet`, sheet);
  organizeSheet(sheet, getDistanceFromSheet(sheet));
});

function getDistanceFromSheet(sheet) {
  let headers = sheet.filter(
    (row) => Object.keys(row).length < 16 && Object.keys(row).length > 2 // Object.keys(row).length == 0
  );
  let data = Object.entries(headers[0]).filter((entry) =>
    entry[1].includes("Destination:")
  );
  return data[0][1];
}

function reassignKeys(tuple) {
  let propKeys = {
    A: "serialNumber",
    B: "dalbitOrderNo",
    C: "dateOfLoading",
    D: "truck_trailerNo",
    E: "qtyLoaded",
    F: "arrivalDate",
    G: "dateOffloaded",
    H: "qtyOffloaded",
    I: "product",
    J: "transitLoss",
    K: "allowedLoss",
    L: "changeableLoss",
    M: "transportRate",
    N: "grossEarning",
    O: "lossValue",
    P: "netEarning",
  };
  let head = propKeys[tuple[0]];
  tuple[0] = head;
  return tuple;
}

function organizeSheet(sheet, destination) {
  let data = [];
  data = sheet.filter((row) => Object.keys(row).length == 16);
  // console.log(`data`, data);
  data = data.map((row) => {
    return Object.entries(row);
  });
  console.log(`${destination}`, data);
  data = data.map((entries) => {
    // console.log(`destination`, destination);
    // console.log(`entries`, entries);
    let e = entries.map((cell) => {
      // let data = [];
      // console.log(`cell`, reassignKeys(cell));
      // data.push(reassignKeys(cell));
      // console.log(`data`, Object.fromEntries(data));
      // return Object.fromEntries(data);
      // console.log(`cell`, cell);
    });
    // return Object.fromEntries(e);
    // console.log(`e`, e);
  });
  // console.log(`${destination}`, data);
}
