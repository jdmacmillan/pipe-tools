
async function handleGutterSpread(event) {
  // The default behavior is for the page to refresh on form submit.  This will prevent that.
  event.preventDefault();

  const files = getFiles();

  const workbook = await processFiles(files);
  const workbookName = getWorkbookName(files);

  downloadWorkbook(workbook, workbookName);
}

// Return the name of the file to be downloaded
function getWorkbookName(files) {
  return 'gutter-spread.xlsx';
}

// Convert the actual files to the desired workbook
async function processFiles(files) {
  const workbook = new ExcelJS.Workbook();
  console.log('initialized workbook', workbook);

  for (file of files) {
    const sheetName = file.name.replace('.txt', '').toUpperCase();
    const sheet = workbook.addWorksheet(sheetName);

    sheet.columns = [
      getColumnDef(20),
      getColumnDef(20),
      getColumnDef(15),
      getColumnDef(),
      getColumnDef(),
      getColumnDef(),
      getColumnDef(),
      getColumnDef(),
      getColumnDef(),
      getColumnDef(),
    ];
    console.log('first sheet', sheet);

    const text = await file.text();
    const lines = text.split('\n')
      .map(formatRow)
      .map(attemptToParse)
      .forEach(r => sheet.addRow(r));

    sheet.insertRow(1, [`${sheetName} (GUTTER SPREAD)`]);
    sheet.mergeCells(0, 0, 0, 9);
  }

  return workbook;
}

function attemptToParse(row) {
  return row.map(column => {
    const number = parseFloat(column, 10);

    if (Number.isNaN(number)) return column;
    return number;
  });
}

function getColumnDef(width = 10) {
  return {
    width: width,
    style: {
      numFmt: '##0.00',
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
      'STRUCTURE', 'INLET TYPE', 'DRAINAGE AREA', 'TC', 'I', 'C', 'Q (INLET)', 'LONG SLOPE', 'GUTTER SPREAD'
    ];
  }

  if (index === 1) {
    return [
      '', '', '(AC)', '(MIN)', '(IN/HR)', '', '(CFS)', '(FT/FT)', '(FT)',
    ];
  }

  const columns = rowText.split(',').filter((_, i) => [0, 3, 9, 10, 11].indexOf(i) === -1);
  if (index >= 2) {
    if (columns[7].toLowerCase() === 'sag' || columns[0].startsWith('DI')) {
      columns[7] = 'N/A';
      columns[8] = 'N/A';
    }

    columns[1] = columns[1].toUpperCase();
  }
  return columns;
}