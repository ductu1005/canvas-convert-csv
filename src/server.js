const express = require('express');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const csvParser = require('csv-parser');

const app = express();
const port = 3001;

// Ensure the uploads directory exists
const uploadsDir = path.join(__dirname, '..', 'uploads');
if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir);
}

// Middleware for handling file uploads
const upload = multer({ dest: uploadsDir });

// Serve static files from the public directory
app.use(express.static(path.join(__dirname, '..', 'public')));


// Function to find column name closest to finalScoreKey
function findClosestColumnName(row, target) {
    let closestMatch = null;
    for (const columnName in row) {
        if (target.trim() === removeNumbersInParentheses(columnName).trim()) {
            closestMatch = columnName;
            break;
        }
    }
    return closestMatch;
}

function removeNumbersInParentheses(str) {
    return str.replace(/\s*\(\d+\)$/, '');
}

function extractClassInfo(input) {
    const regex = /^(.*)\s\((.*)\)$/;
    const match = input.match(regex);

    if (match) {
        const className = match[1].trim();
        const classCode = match[2].trim();
        return { className, classCode };
    } else {
        const className = input;
        const classCode = null;
        return { className, classCode };
    }
}


// Route for converting CSV to Excel
app.post('/convert', upload.array('csvfiles', 10), async (req, res) => {
    const templateFilePath = path.join(__dirname, '..', 'template.xlsx');
    const outputFilePath = path.join(uploadsDir, 'output.xlsx');

    // Load the template
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templateFilePath);
    const worksheet = workbook.getWorksheet(1);

    // Process each uploaded CSV file
    for (const file of req.files) {
        const csvFilePath = file.path;

        // Use csv-parser to parse CSV file
        const csvRows = [];
        let lineNumber = 0;
        fs.createReadStream(csvFilePath)
            .pipe(csvParser())
            .on('data', (row) => {
                lineNumber++;
                if (lineNumber > 1) { // Skip header row
                    csvRows.push(row);
                }
            })
            .on('end', () => {
                const componentScoreText = req.body.componentScore.trim();
                const finalScoreKeyText = req.body.finalScore.trim();
                console.log("-------- ", csvRows);
                let course = null;

                // Add data rows starting from row 10 (adjust as needed)
                let startRow = 12;

                const numberOfRowsToInsert = csvRows.length;
                worksheet.getCell('A13').value = worksheet.getCell('A13').value + numberOfRowsToInsert + " sinh viên";
                worksheet.spliceRows(startRow, 0, ...Array(numberOfRowsToInsert).fill([]));

                csvRows.forEach((row, index) => {
                    const rowNumber = startRow + index;
                    worksheet.getCell(`A${rowNumber}`).value = index + 1; 
                    worksheet.getCell(`B${rowNumber}`).value = row['SIS User ID'];
                    worksheet.getCell(`C${rowNumber}`).value = row['Student']; 
                    const componentScore = findClosestColumnName(row, componentScoreText);
                    const finalScoreKey = findClosestColumnName(row, finalScoreKeyText);

                    if (componentScore)
                        worksheet.getCell(`F${rowNumber}`).value = row[componentScore];

                    if (finalScoreKey) 
                        worksheet.getCell(`G${rowNumber}`).value = row[finalScoreKey]; 

                    worksheet.getCell(`H${rowNumber}`).value = worksheet.getCell(`F${rowNumber}`).value * worksheet.getCell('G7').value +
                                                                   worksheet.getCell(`G${rowNumber}`).value * worksheet.getCell('G8').value;

                    ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'].forEach(col => {
                        const cell = worksheet.getCell(`${col}${rowNumber}`);
                        cell.border = {
                            top: { style: 'thin' },
                            left: { style: 'thin' },
                            bottom: { style: 'thin' },
                            right: { style: 'thin' }
                        };
                    });

                    if (index < 2) {
                        course = row['Section'];
                    }
                });
                const classInfo = extractClassInfo(course);
                if (classInfo && classInfo.classCode) {
                    worksheet.getCell('C8').value = classInfo.classCode;
                }
                if (classInfo && classInfo.className) {
                    worksheet.getCell('C9').value = classInfo.className;
                }

                // Remove the temporary CSV file
                fs.unlinkSync(csvFilePath);

                // Save the populated template to the output file
                workbook.xlsx.writeFile(outputFilePath)
                    .then(() => {
                        res.download(outputFilePath, 'output.xlsx', (err) => {
                            if (err) console.error(err);
                            // Remove the Excel file after download
                            fs.unlinkSync(outputFilePath);
                        });
                    })
                    .catch((err) => {
                        console.error('Error saving Excel file:', err);
                        res.status(500).send('Error converting CSV to Excel');
                    });
            });
    }
});

app.listen(port, () => {
    console.log(`Server running on http://localhost:${port}`);
});
