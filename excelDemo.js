const ExcelJS = require('exceljs');

async function writeExcelTest(searchText, replaceText, change, filePath) {

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.getWorksheet('Sheet1');

    const output = readExcel(worksheet, searchText);
    if (!output) {
        return;
    }

    const cell = worksheet.getCell(
        output.row + change.rowChange,
        output.column + change.columnChange
    );
    cell.value = replaceText;

    await workbook.xlsx.writeFile(filePath);
    console.log("Planilha atualizada!");
}

function readExcel(worksheet, searchText) {

    let output = { row: -1, column: -1 };

    worksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
            if (cell.value === searchText) {
                output.row = rowNumber;
                output.column = colNumber;
            }
        });
    });

    if (output.row === -1) {
        console.log(`Texto '${searchText}' não encontrado na planilha`);
        return null;
    }
    return output;
}

writeExcelTest(
    "Mango",
    350,
    { rowChange: 0, columnChange: 2 },
    "C:\\Users\\natal\\Downloads\\download.xlsx"
);