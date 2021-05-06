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

// At the moment i'm delete this because it's the only sheet different from the rest
// should create a separate function for this sheet
delete workbook["17.12.2020"];
// console.log(`workbook`, workbook);

// Worksheets
let worksheets = [];

// This will hold the final output
let output = [];

for (let sheet of Object.values(workbook)) {
  // Leave out empty sheets
  if (sheet.length != 0) {
    worksheets.push(sheet);
  }
}

// console.log(`worksheets`, worksheets);

worksheets.forEach((sheet) => {
  // Organize every row in every sheet so they are in a good structure...
  let data = organizeSheet(sheet, getDistanceFromSheet(sheet));
  output.push(...data);
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
  data = sheet.filter((row) => {
    if (Object.keys(row).length == 16) {
      return row;
    }
  });

  data = data.map((row) => {
    return Object.entries(row);
  });

  data = data.map((entries) => {
    let e = entries.map((cell) => {
      return reassignKeys(cell);
    });
    e = Object.fromEntries(e);
    e.dateOfLoading = new Date(e.dateOfLoading.split(".").reverse().join("-"));
    e.arrivalDate = new Date(e.arrivalDate.split(".").reverse().join("-"));
    e.dateOffloaded = new Date(e.dateOffloaded.split(".").reverse().join("-"));
    e.vehicle = e.truck_trailerNo.split(" ").join("").split("/")[0];
    e.trailer = e.truck_trailerNo.split(" ").join("").split("/")[1];
    e.destination = destination;
    return e;
  });
  // console.log(`data`, data);
  return data;
}

console.log(`output`, output);
