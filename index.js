const express = require("express");
const app = express();
const PORT = process.env.PORT || 3030;

const xlsx = require('xlsx');
var fs = require('fs');

function convertExcelFileToJsonUsingXlsx() {

    // Read the file using pathname
    const file = xlsx.readFile('./Data.xlsx');

    // Grab the sheet info from the file
    const sheetNames = file.SheetNames;
    const totalSheets = sheetNames.length;

    // Variable to store our data
    let parsedData = [];

    // Loop through sheets
    for (let i = 0; i < totalSheets; i++) {

        // Convert to json using xlsx
        const tempData = xlsx.utils.sheet_to_json(file.Sheets[sheetNames[i]]);

        // Skip header row which is the colum names
        tempData.shift();

        // Add the sheet's json to our data array
        parsedData.push(...tempData);
    }

    return parsedData;
}

app.get('/',(req,res)=>{
    res.send(convertExcelFileToJsonUsingXlsx());
})

app.listen(PORT, () => {
  console.log(`server started on port ${PORT}`);
});