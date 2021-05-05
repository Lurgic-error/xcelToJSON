const excelToJSON = require('convert-excel-to-json')

const workbook = excelToJSON({
    sourceFile: 'sheets.xlsx',
    header: {
        rows: 5
    },
    // columnToKey: {
    //     '*': '{{columnHeader}}'
    // }
})

// console.log(`workbook`, workbook)
let rows =[]
for(let worksheet in workbook){
// console.log(`rows ==>`, workbook[worksheet]
//                                 .filter(element => Object.keys(element)
//                                         .length == 16))
   
}

// let sheets = Object.keys(workbook)
// // console.log(`sheets`, sheets)
// let sample = workbook['31.3.2021B']

// // console.log(`sample`, sample)
let propKeys = [
    'serial number', 'Dalbit Order No', 'date of loading',
    'truck/trailer no', 'qty loaded', 'arrival date',
    'date offloaded', 'qty offloaded', 'product', 'transit loss',
    'allowed loss', 'changeable loss', 'transport rate',
    'gross earning', 'loss value', 'net earning'
]

// let data = []
// for (let row of sample) {
    
//     if (Object.keys(row).length == 16) {
//         data.push(row)
//         // console.log(`row`, row)
//         // for(let prop in row){
//         //     // console.log(`${prop}:${row[prop]}`)
//         //     for(let key in propKeys){
//         //         console.log(`${propKeys[key]}:${row[prop]}`)
//         //     }
//         // }
//     }
// }

// let res = data.map(item => {
//     console.log(`item`, item)
//     propKeys.forEach(element => {
//         console.log(`element`, element)
//     });
// })
// console.log(`data`, data)
// let company = sample