

function addPositive() {
  appendTimestampToColumnOnCurrentDay(1);
}


function addNegative() {
  appendTimestampToColumnOnCurrentDay(2);
}


function appendTimestampToColumnOnCurrentDay(column) {
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  // const spreadsheet = SpreadsheetApp.openById("1m1Mub3a-Tp0rPuChY3_PuPwMg8sGgZfVWUMr9mhSOF8");
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(today);
  if(!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(today, 2);
    sheet.getRange(1,1,1,3).setValues([["Positive", "Negative", "Attendance"]]);
    sheet.getRange('A2:A').setNumberFormat('HH:mm:ss');
    sheet.getRange('B2:B').setNumberFormat('HH:mm:ss');
    sheet.getRange('C2:C').setNumberFormat('HH:mm:ss');
  }
  
  // Get the current number of rows with data in Column A
  const lastRow = sheet.getRange(sheet.getMaxRows(), column).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  
  // Calculate the next row position
  const nextRow = lastRow + 1;
  
  // Get current time and append to Column A
  const currentTime = new Date();
  sheet.getRange(nextRow, column).setValue(currentTime);
}

function onEdit() {
  let activeCell = SpreadsheetApp.getActiveSpreadsheet().getActiveCell();
  let reference = activeCell.getA1Notation();
  let sheetName = activeCell.getSheet().getSheetName();
  let value = activeCell.getValue();
  if (sheetName == "Buttons" && value) {
    if (reference == "A2") {
      appendTimestampToColumnOnCurrentDay(1);
      activeCell.setValue(false);
    } else if (reference == "B2") {
      appendTimestampToColumnOnCurrentDay(2);
      activeCell.setValue(false);
    } else if (reference == "C2") {
      appendTimestampToColumnOnCurrentDay(3);
      activeCell.setValue(false);
    }
  }
}