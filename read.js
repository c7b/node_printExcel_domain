// Requiring the module
const reader = require('xlsx')
  
// Reading our test file

let currentDate = new Date().toISOString().split('T')[0]


const file = reader.readFile(`./Excel${currentDate}.xlsx`)


  
let data = []
  
const sheets = file.SheetNames
  
for(let i = 0; i < sheets.length; i++)
{
   const temp = reader.utils.sheet_to_json(
        file.Sheets[file.SheetNames[i]])
   temp.forEach((res) => {
      data.push(res)
   })
}
  
// Printing data
console.log(data)