function attemptToParse(row) {
  return row.map(column => {
    const number = parseFloat(column, 10);

    if (Number.isNaN(number)) return column;
    return number;
  });
}

function getFiles(inputType) {
  const fileInput = document.getElementById(`${inputType}Files`);
  return fileInput.files;
}

// This won't need to change at all.
function downloadWorkbook(workbook, workbookName) {
  // Write the workbook that's been created into a buffer of content
  workbook.xlsx.writeBuffer({ base64: true })
    .then(binaryData => {
      // Basically this will create a new, hidden link that will download this file, then click the link.
      var a = document.createElement("a");
      document.body.appendChild(a);
      a.style = "display: none";

      // Create the actual URL from the buffer.
      var url = window.URL.createObjectURL(new Blob([binaryData], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));

      a.href = url;
      a.download = workbookName;

      a.click();

      // Cleanup
      window.URL.revokeObjectURL(url);
    })
    .catch(function (error) {
      console.error(error.message);
    });
}

function attemptToParse(row) {
  return row.map(column => {
    const number = parseFloat(column, 10);

    if (Number.isNaN(number)) return column;
    return number;
  });
}
