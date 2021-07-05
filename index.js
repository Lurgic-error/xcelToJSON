const { writeFile } = require("fs");
const excelToJSON = require("convert-excel-to-json");
const tripCollection = require("./database");
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
  let data = organizeSheet(sheet, getDestination(sheet));
  output.push(...data);
});

// A function to get the destination from the sheet header row 2
function getDestination(sheet) {
  let headers = sheet.filter(
    (row) => Object.keys(row).length < 16 && Object.keys(row).length > 2
  );
  let data = Object.entries(headers[0]).filter((entry) =>
    entry[1].includes("Destination:")
  );
  // Return destination
  return data[0][1];
}

function reassignKeys(tuple) {
  /**
   * This function reassigns keys
   * TODO: Tony explain it further please...
   */
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
  /***
   * A function to organize sheet by getting the data
   * from each row from the sheet and creating an object to represent
   * that row with keys matching those in the sheet's column
   */
  let rows = [];
  rows = sheet.filter((row) => {
    if (Object.keys(row).length == 16) {
      return row;
    }
  });

  rows = rows.map((dataRow) => {
    return Object.entries(dataRow);
  });

  rows = rows.map((entries) => {
    let row = entries.map((cell) => {
      return reassignKeys(cell);
    });
    row = Object.fromEntries(row);
    row.dateOfLoading = new Date(
      row.dateOfLoading.split(".").reverse().join("-")
    ); // Change date of loading to a date object.
    row.arrivalDate = new Date(row.arrivalDate.split(".").reverse().join("-")); // Change arrival date to a date object.
    row.dateOffloaded = new Date(
      row.dateOffloaded.split(".").reverse().join("-")
    ); // Change date offloaded to a date object.
    row.vehicle = row.truck_trailerNo.split(" ").join("").split("/")[0]; // Create a vehicle property
    row.trailer = row.truck_trailerNo.split(" ").join("").split("/")[1]; // Create a trailer property
    row.destination = destination.slice(13); // Sliced "Destination: " from destination
    return row;
  });
  return rows;
}

// Remove "." from some destinations
output.map((data) => {
  if (data.destination.indexOf(".") > 0) {
    let word = data.destination.split("");
    word = word.filter((letter) => {
      if (letter != ".") {
        return letter;
      }
    });
    data.destination = word.join("");
  }
});

let destinations = Array.from(
  new Set([...output.map((row) => row.destination)])
);

let products = Array.from(new Set([...output.map((row) => row.product)]));

let vehicles = Array.from(
  new Set([
    ...output.map((row) => {
      if ((row.vehicle.length = 7)) return row.vehicle;
    }),
  ])
);

let trailers = Array.from(new Set([...output.map((row) => row.trailer)]));

console.log(`destinations`, destinations);
console.log(`vehicles`, vehicles.length);
console.log(`trailers`, trailers);
console.log(`trailers`, trailers.length);

let totalNetEarnings = () => {
  let total = 0;
  output.forEach((row) => (total += row.netEarning));
  return total;
};

let totalGrossEarnings = () => {
  let total = 0;
  output.forEach((row) => (total += row.grossEarning));
  return total;
};

let totalQtyLoaded = () => {
  let total = 0;
  output.forEach((row) => (total += row.qtyLoaded));
  return total;
};

let totalQtyOffloaded = () => {
  let total = 0;
  output.forEach((row) => (total += row.qtyOffloaded));
  return total;
};

let totalLossValue = () => {
  let total = 0;
  output.forEach((row) => (total += row.lossValue));
  return total;
};

let totalAllowedLoss = () => {
  let total = 0;
  output.forEach((row) => (total += row.allowedLoss));
  return total;
};

let totalChangeableLoss = () => {
  let total = 0;
  output.forEach((row) => (total += row.changeableLoss));
  return total;
};

let totalTransitLoss = () => {
  let total = 0;
  output.forEach((row) => (total += row.transitLoss));
  return total;
};

console.log(`totalQtyLoaded`, totalQtyLoaded());

console.log(`totalQtyOffloaded`, totalQtyOffloaded());

console.log(`totalLossValue`, totalLossValue());

console.log(`totalAllowedLoss`, totalAllowedLoss());

console.log(`totalChangeableLoss`, totalChangeableLoss());

console.log(`totalTransitLoss`, totalTransitLoss());

console.log(`totalNetEarnings`, totalNetEarnings());

console.log(`totalGrossEarnings`, totalGrossEarnings() * 2300);
// console.log(`output`, output);

console.log(15 * 365 * 10);

// writeFile("data.json", JSON.stringify(output), (err) => {
//   if (err) {
//     throw new Error(err);
//   } else {
//     console.log("kimepop");
//   }
// });

// Save each row in the database
// output.forEach(async (row) => {
//   await tripCollection.create(row);
// });

tripCollection.find({}).then((rows) => console.log("rows", rows));
