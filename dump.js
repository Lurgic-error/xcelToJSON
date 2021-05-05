const excelToJSON = require("convert-excel-to-json");

// Get  the workbook
const workbook = excelToJSON({
  sourceFile: "sheets.xlsx",
  // header: {
  //     rows: 5
  // },
  // columnToKey: {
  //     '*': '{{columnHeader}}'
  // }
});

// console.log(`workbook`, workbook)

let entries = [];
for (let sheet in workbook) {
  if (workbook[sheet].length != 0) {
    createSheets(workbook[sheet]);
  }

  if (typeof workbook[sheet][1] === "object") {
  }

  for (let row of workbook[sheet]) {
    if (Object.keys(row).length == 16) {
      entries.push(row);
    }
  }
}

function createSheets(sheet) {
  let destination = sheet[1].E;
  let sheets = [];
  let dataRows = [];

  // console.log(`sheet`, sheet);
  // console.log(`destination`, destination, sheet.length);
}
let rows = entries.map((entry) => {
  return Object.entries(entry);
});

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

function reassignKeys(tuple) {
  let head = propKeys[tuple[0]];
  tuple[0] = head;
  return tuple;
}

let newRows = rows.map((row) => {
  let roww = row.map((item) => {
    return reassignKeys(item);
  });
  return Object.fromEntries(roww);
});

// console.log(`newRows`, newRows);

let date = new Date("20.12.2021".split(".").reverse().join("-"));

// console.log(`date`, date)

newRows.forEach((row) => {
  row.dateOfLoading = new Date(
    row.dateOfLoading.split(".").reverse().join("-")
  );
  row.arrivalDate = new Date(row.arrivalDate.split(".").reverse().join("-"));
  row.dateOffloaded = new Date(
    row.dateOffloaded.split(".").reverse().join("-")
  );
  row.vehicle = row.truck_trailerNo.split(" ").join("").split("/")[0];
  row.trailer = row.truck_trailerNo.split(" ").join("").split("/")[1];
});

// console.log(`newRows`, newRows);
console.log(`rows`, rows);
