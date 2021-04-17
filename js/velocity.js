const MATERIAL_CONVERSIONS = {
  0.013: 'RCP',
};

const velocity = (() => {
  async function handleVelocity(event) {
    // The default behavior is for the page to refresh on form submit.  This will prevent that.
    event.preventDefault();

    const files = getFiles("velocity");

    const workbook = await processFiles(files);
    const workbookName = getWorkbookName(files);

    downloadWorkbook(workbook, workbookName);
  }

  // Return the name of the file to be downloaded
  function getWorkbookName(files) {
    return 'pipe-velocity.xlsx';
  }

  // Convert the actual files to the desired workbook
  async function processFiles(files) {
    const workbook = new ExcelJS.Workbook();
    console.log('initialized workbook, yay', workbook, files);

    for (file of files) {
      const sheetName = file.name.replace('.txt', '').toUpperCase();
      console.log('sheetName: ', sheetName);

      const sheet = workbook.addWorksheet(sheetName);
      console.log('sheet', sheet);

      sheet.columns = [
        getColumnDef(),
        getColumnDef(),
        getColumnDef(),
        getColumnDef(),
        getColumnDef(),
        getColumnDef(),
        getColumnDef(),
        getColumnDef(),
        getColumnDef(),
        getColumnDef(),
        getColumnDef(10, "##0.000"),
        getColumnDef(),
      ];
      console.log('first sheet', sheet);

      const text = await file.text();      
      const dataTable = text.split('\n')
        .map(formatRow)
        .map(attemptToParse);

      console.log(dataTable);
      updateToCellName(dataTable);
        
      dataTable.forEach(r => sheet.addRow(r));

      sheet.insertRow(1, [`${sheetName} (2-YEAR ANALYSIS)`]);
      sheet.mergeCells(0, 0, 0, 12);
      sheet.mergeCells(2, 2, 2, 3);
    }

    return workbook;
  }

  function updateToCellName(dataTable) {
    for (let i = 2; i < dataTable.length; i++) {
      const row = dataTable[i];  
      const toCellValue = row[2];
      let cellValueAsInt;
      try {
        cellValueAsInt = parseInt(toCellValue, 10);
      } catch {}
      if (cellValueAsInt !== undefined && !Number.isNaN(cellValueAsInt)) {
        row[2] = dataTable[cellValueAsInt + 1][1];
      } else {
        row[2] = typeof row[2] === 'string' ? row[2].toUpperCase() : row[2];
      }
    }
  }

  function getColumnDef(width = 10, numFmt = "##0.00") {
    return {
      width,
      style: {
        numFmt,
        alignment: {
          vertical: 'middle',
          horizontal: 'center',
          wrapText: true
        },
      }
    }
  }

  function formatRow(rowText, index) {
    if (index === 0) {
      return [
        'INLET TYPE', 'STRUCTURE', '', 'A(TOTAL)', 'TC', 'I', 'Q', 'V', 'PIPE LENGTH', 'PIPE SIZE', 'MATERIAL', 'SLOPE'
      ];
    }

    if (index === 1) {
      return [
        '', 'FROM', 'TO', '(AC)', '(MIN)', '(IN/HR)', '(CFS)', '(FT/S)', '(FT)', '(IN)', '', '(%)'
      ];
    }

    const columns = rowText.split(',').filter((_, i) => [0].indexOf(i) === -1);
    if (index >= 2) {
      // if (columns[3].toLowerCase() === 'outfall' || columns[0].startsWith('DI')) {
      //   columns[3] = 'OUT';
      //   columns[8] = 'N/A';
      // }
      columns[0] = typeof columns[0] === 'string' ? columns[0].toUpperCase() : columns[0];

      columns[10] = MATERIAL_CONVERSIONS[columns[10]] || columns[10];
    }
    return columns;
  }

  return handleVelocity;
})();