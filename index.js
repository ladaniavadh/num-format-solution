console.clear();

const exceljs = require("exceljs");
const fs = require("fs");

const data = [
    {test: "00:20"},
    {test: "01:10"},
    {test: "00:45"}
]

const workbook = new exceljs.Workbook();
const worksheet = workbook.addWorksheet("test worksheet");

worksheet.columns = [
    {header: 'Data', key: 'test'}
  ]
  worksheet.getColumn(1).numFmt = '[hh]:mm:ss'

  data.forEach((e, index) => {
    // row 1 is the header.
    const rowIndex = index + 2
  
    // By using destructuring we can easily dump all of the data into the row without doing much
    // We can add formulas pretty easily by providing the formula property.
    worksheet.addRow({
      ...e,
    })
  })

  let overTimeFormula = ''
  worksheet.eachRow(function (row, rowNumber) {
    if (rowNumber != 1) {
      worksheet.getRow(rowNumber).eachCell((cell, cellNumber) => {
        if (cellNumber == 1) {
          if (overTimeFormula == '') {
            overTimeFormula = `A${rowNumber}`
          } else {
            overTimeFormula = `${overTimeFormula}+A${rowNumber}`
          }
        }
      })
    }
  });
// //   console.log('overTimeFormula', overTimeFormula);
//   // Add the total Rows
  worksheet.addRow([
    {
      formula: `=${overTimeFormula}`
    }
  ])


//   worksheet.getColumn(1).values = data;
workbook.xlsx
    .writeFile("output.xlsx")
    .then(
        () => {
            console.log("workbook saved!");
        }
    )
    .catch(
        (error) => {
            console.log("something went wrong!");
            console.log(error.message);
        }
    );