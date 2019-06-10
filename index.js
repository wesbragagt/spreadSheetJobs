var xl = require("excel4node");

var wb = new xl.Workbook();

var ws = wb.addWorksheet("Sheet 1");

var style = {
    font: {
        color: "#000000",
        size: 12
    },
    numberFormat: "$#, ##0.00; ($#,##0.00); -"
};

var arr = ["aol", "google", "apple"];

function genSpreadsheet(arr, column) {
    arr.forEach(function(element, index) {
        ws.cell(2 + index, column)
            .string(element)
            .style(style);
    });
}

genSpreadsheet(arr, 4);

// Set value of cell A1 to 100 as a number type styled with paramaters of style
ws.cell(1, 1)
    .string("Employer")
    .style(style);

ws.column().setWidth(25);
ws.row(1).setHeight(25);

// Set value of cell B1 to 200 as a number type styled with paramaters of style
ws.cell(1, 2)
    .string("Location")
    .style(style);

// Set value of cell C1 to a formula styled with paramaters of style
ws.cell(1, 3)
    .string("Applied")
    .style(style);

ws.cell(1, 4)
    .string("Interviews")
    .style(style);

wb.write("Excel.xlsx");
