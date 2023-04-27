// Require library
const express = require('express');
const bodyParser = require('body-parser');
const path = require('path');
const fs = require('fs');
//const xl = require('excel4node');


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


const app = express();

// Middleware to parse incoming request bodies
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());

// Serve static files from the "public" folder
app.use(express.static(path.join(__dirname, 'public')));


// sendFile will go here
app.get('/', function(req, res) {
    res.sendFile(path.join(__dirname, '/index.html'));
  });


// Route to generate and download the Excel file
app.get('/download', (req, res) => {
  // Set the content of the text file
  const content = `Text file created in the Backend with NodeJs.
  Number of trys: ${loopNumber}`;

  // Set the file name
  const fileName = 'loop.txt';

  // Set the appropriate headers and send the content as a response
  res.setHeader('Content-Type', 'text/plain');
  res.setHeader('Content-Disposition', `attachment; filename=${fileName}`);
  res.send(content);
});


// Start the server
const port = process.env.PORT || 3000;
app.listen(port, () => { console.log(`Server is running on http://localhost:${port}`);});



/*
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
let currentDate = new Date().toISOString().split('T')[0]


wb.write(`Excel${currentDate}.xlsx`);

*/