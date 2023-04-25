// Require library
const express = require('express');
const bodyParser = require('body-parser');
const path = require('path');
const fs = require('fs');
const xl = require('excel4node');


let loopNumber;

// Check if the file exists
if (fs.existsSync('loop.txt')) {
  // If the file exists, read its contents into a variable
  loopNumber = fs.readFileSync('loop.txt', 'utf-8');
} else {
  // If the file doesn't exist, create it with a default value
  loopNumber = 1;
  fs.writeFileSync('loop.txt', loopNumber.toString(), 'utf-8');
  console.log('The file has been created with default value 1.');
}

// Increment the loop number and write it back to the file
loopNumber++;
console.log('The file exists and its contents are:', loopNumber);
fs.writeFileSync('loop.txt', loopNumber.toString(), 'utf-8');


// Create a new instance of a Workbook class
var wb = new xl.Workbook();

// Add Worksheets to the workbook
var ws = wb.addWorksheet('Sheet 1');
var ws2 = wb.addWorksheet('Hello');

// Create a reusable style
var style = wb.createStyle({
  font: {
    color: '#000000',
    size: 12,
  }
});


let topStyle = wb.createStyle({
    font: {
      color: '#000000',
      size: 12,
      bold: true,
    }
  });

let contentStyle = wb.createStyle({
    font: {
      color: '#000000',
      size: 12,
      bold: false,
    }
  });


//Linea
for (let x = 1; x < 11; x++) {
    ws.cell(x + 1, 1).number(x).style(topStyle);
}
//Columna
for (let y = 1; y < 11; y++) {
    ws.cell(1, y + 1).number(y).style(topStyle);
}


// Multiplicacion
for (let a = 1; a < 11; a++) {
    for (let b = 1; b < 11; b++) {
        const rowHeader = excelRowColToCell(a + 1, 1);
        const columnHeader = excelRowColToCell(1, b + 1);
        ws.cell(a + 1, b + 1).formula(`${rowHeader}*${columnHeader}`);
    }
}

// Utility function to convert row and column numbers to Excel cell references (e.g. A1, B2, etc.)
function excelRowColToCell(row, col) {
    const letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    let cell = '';
    while (col > 0) {
        const remainder = (col - 1) % 26;
        cell = letters[remainder] + cell;
        col = (col - 1 - remainder) / 26;
    }
    cell += row;
    return cell;
}

ws2.cell(1,1).string('Hello World');
/*
}*/
/*
// Set value of cell A1 to 100 as a number type styled with paramaters of style
ws.cell(1, 1)
  .number(100)
  .style(style);

// Set value of cell B1 to 200 as a number type styled with paramaters of style
ws.cell(1, 2)
  .number(200)
  .style(style);

// Set value of cell C1 to a formula styled with paramaters of style
ws.cell(1, 3)
    .formula('A1 + B1')
  .style(style);

// Set value of cell A2 to 'string' styled with paramaters of style
ws.cell(2, 1)
  .string('string')
  .style(style);

// Set value of cell A3 to true as a boolean type styled with paramaters of style but with an adjustment to the font size.
ws.cell(3, 1)
  .bool(true)
  .style(style)
  .style({font: {size: 14}});
*/

let currentDate = new Date().toISOString().split('T')[0]


wb.write(`Excel${currentDate}.xlsx`);



const app = express();

// Middleware to parse incoming request bodies
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());

// Serve static files from the "public" folder
app.use(express.static(path.join(__dirname, 'public')));

// Route to generate and download the Excel file
app.get('/download', (req, res) => {
  // Generate the Excel file here (use your existing code to create the workbook and worksheets)
  
  // Save the workbook to a buffer
  const currentDate = new Date().toISOString().split('T')[0];
  const fileName = `Excel${currentDate}.xlsx`;

  wb.writeToBuffer().then((buffer) => {
    // Set the appropriate headers and send the buffer as a response
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=${fileName}`);
    res.send(buffer);
  }).catch((err) => {
    console.error(err);
    res.status(500).send('Error generating the Excel file');
  });
});

// Start the server
const ipAddress = '0.0.0.0';
const port = process.env.PORT || 3000;

app.listen(port, ipAddress, () => {
    console.log(`Server is running on http://${ipAddress}:${port}`);
  });



