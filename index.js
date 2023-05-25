const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const translationData = require('./src/data.json');
const worksheetName = 'Translation';
const exportedFile = `./exported/sheet-${(new Date()).getTime()}.xlsx`;

const writeExcelFile = async ({ fileName, data }) => {

  // Create the Workbook
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Cromwell Digital Team';
  workbook.created = new Date();
  workbook.modified = new Date();

  // Add the Worksheet
  const worksheet = workbook.addWorksheet(worksheetName, {});

  // Define column settings
  const columnSettings = [
    { header: "Domain id", key: "domain_config_id", width: 12 },
    { header: "Page id", key: "page_id", width: 12 },
    { header: "Page name", key: "page_name", width: 12 },
    { header: "Key", key: "key", width: 24 },
    { header: "English version", key: "english_version", width: 56 },
    { header: "For translation", key: "for_translation", width: 56 },
    { header: "For API", key: "is_api_translation", width: 12 },
    { header: "Notes", key: "notes", width: 56 },
  ];

  worksheet.columns = columnSettings;

  // Add the json data as new rows
  data.forEach((item) => {
    worksheet.addRow({ ...item });
  });

  // Style each cell
  columnSettings.forEach((col, index) => {
    worksheet.getColumn(index + 1).eachCell((cell) => {
      cell.alignment = {
        horizontal: 'left',
        vertical: 'top',
        wrapText: true
      };
      if (col.key === 'for_translation') {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'AAFFFFCC' }
        };
        cell.border = {
          bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
        };
      }
    });
  });

  // Make the to_translate the only editable column
  worksheet.getColumn('for_translation').protection = {
    locked: false,
    lockText: false
  };

  // Style the first row for headers
  const headerRow = worksheet.getRow(1);
  headerRow.protection = { locked: true, lockText: true };
  headerRow.eachCell({ includeEmpty: true }, (cell) => {
    cell.font = {
      color: { argb: 'FF000000' },
      family: 2,
      size: 10,
      bold: true
    };
  });

  // Lock the whole worksheet without a password
  // The to_translate column will remain editable
  const password = '';
  worksheet.protect(password, {
    selectLockedCells: false,
    selectUnlockedCells: true,
  });

  // Export the Excel file
  try {
    const destPath = path.join(__dirname, fileName);
    const buffer = await workbook.xlsx.writeBuffer();
    await fs.promises.writeFile(destPath, buffer);
    console.log(`### Exported Excel file: ${fileName}`);
  } catch (err) {
    console.log("### Error exporting Excel file", err);
  }
}

const readExcelFile = async ({ fileName }) => {
  const excelData = [];
  try {
    const headerRowIndex = 1;
    const newWorkbook = new ExcelJS.Workbook();
    const destPath = path.join(__dirname, fileName);
    await newWorkbook.xlsx.readFile(destPath);
    newWorkbook.getWorksheet(worksheetName).eachRow((row, rowIndex) => {
      if (rowIndex > headerRowIndex) {
        const [_,
          domain_config_id,
          page_id,
          page_name,
          key,
          english_version,
          for_translation,
          is_api_translation,
          notes
        ] = row.values;
        excelData.push({
          domain_config_id,
          page_id,
          page_name,
          key,
          english_version,
          for_translation,
          is_api_translation,
          notes
        });
      }
    });
    console.log(`### Read Excel file: ${fileName}`);
  } catch (err) {
    console.log("### Error reading Excel file", err);
  }
  return excelData;
}

// Main function to write and then read Excel file
const main = async () => {
  await writeExcelFile({ fileName: exportedFile, data: translationData });
  const excelData = await readExcelFile({ fileName: exportedFile });
  console.log(JSON.stringify(excelData, null, 2));
};

main();