function setCellAlign(sheet, range, alignment) {
   var cell = sheet.getRange(range);
   cell.setHorizontalAlignment(alignment);
}

function correctAddress(sheet, range) {
 var values = sheet.getRange(range).getValues();
 var addressParts = values[0].toString().split(',');
 var formattedAddress = '';
 for (var i = 0; i < addressParts.length; i++) {
    var part = addressParts[i].trim();
    if (i === addressParts.length - 1) {
      formattedAddress += part;
    } else {
      formattedAddress += part + ', ';
    }
 }
 values[0] = [formattedAddress]; // Change this line
 sheet.getRange(range).setValues(values);
}

function cleanName(sheet, range) {
  var values = sheet.getRange(range).getValues();

  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      var cellValue = values[i][j];
      if (typeof cellValue === 'string') {
        // Maintain the dot following specific prefixes (case-insensitive)
        // cellValue = cellValue.replace(/(Dr|Mrs|Ms|Mr)\.?(\s|$)/gi, '$1.$2');

        // Replace dot followed by space or at the end of the string with a single space.
        cellValue = cellValue.replace(/\.\s*|\s*\./g, ' ');

        // Remove any extra spaces and trim the result.
        cellValue = cellValue.replace(/\s+/g, ' ').trim();

        // Update the value in the 'values' array.
        values[i][j] = cellValue;
      }
    }
  }
  sheet.getRange(range).setValues(values);
}


// Function to convert a string to sentence case.
function toTitlecase(sheet,range) {
  // Get the values in the specified range.
  var values = sheet.getRange(range).getValues();

  // Loop through the values and convert to sentence case.
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      var cellValue = values[i][j];
      if (typeof cellValue === 'string' && cellValue.trim() !== '') {
        values[i][j] = cellValue.replace(/\w\S*/g, function (txt) {
          return txt.charAt(0).toUpperCase() + txt.substring(1).toLowerCase();
        });
      }
    }
  }
  // Set the updated values back to the range.
  sheet.getRange(range).setValues(values);
}

// Function to convert a string to sentence case.
function toSentenceCase(sheet,range) {

  // Get the values in the specified range.
  var values = sheet.getRange(range).getValues();

  // Loop through the values and convert to sentence case.
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      var cellValue = values[i][j];
      if (typeof cellValue === 'string' && cellValue.trim() !== '') {
        values[i][j] = cellValue.charAt(0).toUpperCase() + cellValue.substr(1).toLowerCase();
      }
    }
  }
  // Set the updated values back to the range.
  sheet.getRange(range).setValues(values);
}

function toLowercase(sheet,range) {
  // Get the values in the specified range.
  var values = sheet.getRange(range).getValues();

  // Loop through the values and convert to lowercase.
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      var cellValue = values[i][j];
      if (typeof cellValue === 'string' && cellValue.trim() !== '') {
        values[i][j] = cellValue.toLowerCase();
      }
    }
  }

  // Set the updated values back to the range.
  sheet.getRange(range).setValues(values);
}

function toUppercase(sheet,range) {
  // Get the values in the specified range.
  var values = sheet.getRange(range).getValues();

  // Loop through the values and convert to uppercase.
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      var cellValue = values[i][j];
      if (typeof cellValue === 'string' && cellValue.trim() !== '') {
        values[i][j] = cellValue.toUpperCase();
      }
    }
  }

  // Set the updated values back to the range.
  sheet.getRange(range).setValues(values);
}

function onFormSubmit() {
  var sheetName1 = "Responses";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName(sheetName1);

  setCellAlign(sheet1,"A2:D", "center");
  setCellAlign(sheet1,"E2:E", "left");
  setCellAlign(sheet1,"F2:I", "center");
  setCellAlign(sheet1,"J2:J", "left");
  setCellAlign(sheet1,"K2:L", "center");
  setCellAlign(sheet1,"M2:M", "left");
  setCellAlign(sheet1,"N2:N", "center");
  setCellAlign(sheet1,"O2:O", "left");
  setCellAlign(sheet1,"P2:P", "center");
  setCellAlign(sheet1,"Q2:Q", "left");
  setCellAlign(sheet1,"R2:S", "center");
  setCellAlign(sheet1,"T2:U", "left");
  setCellAlign(sheet1,"V2:V", "center");
  cleanName(sheet1,"E2:E");
  correctAddress(sheet1,"K2:K");
  cleanName(sheet1,"O2:O");
  cleanName(sheet1,"M2:M");
  cleanName(sheet1,"U2:U");
  toTitlecase(sheet1,"E2:E");
  toTitlecase(sheet1,"O2:O");
  toTitlecase(sheet1,"M2:M");
  toTitlecase(sheet1,"U2:U");
  toUppercase(sheet1,"C2:D");
  toTitlecase(sheet1,"J2:J");
  toUppercase(sheet1,"Q2:Q");
  toTitlecase(sheet1,"K2:K");
  toUppercase(sheet1,"S2:S");
  toTitlecase(sheet1,"T2:T");
}
