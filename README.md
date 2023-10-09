# Google Sheets Script Documentation

## Overview

This Google Sheets script provides a suite of utility functions for manipulating and formatting the data in a Google Sheet. These functions include aligning cells, cleaning and transforming names, fixing addresses, and converting text to different cases.

## Function Descriptions

### setCellAlign(sheet, range, alignment)

**Description:** This function sets the horizontal alignment of text in the specified range of a sheet.

**Parameters:**

- **sheet**: The sheet object where the range is located.
- **range**: A string that defines the range of cells to be aligned (e.g., "A2:D").
- **alignment**: A string that specifies the desired alignment, which can be "center", "left", "right", etc.

### correctAddress(sheet, range)

**Description:** This function cleans up and formats the address in a specified range of the sheet by trimming whitespaces and adding commas between address parts.

**Parameters:**

- **sheet**: The sheet object where the range is located.
- **range**: A string that defines the range of cells containing addresses to be formatted (e.g., "K2:K").

### cleanName(sheet, range)

**Description:** This function cleans up names in a specified range of the sheet by removing extra spaces and dots.

**Parameters:**

- **sheet**: The sheet object where the range is located.
- **range**: A string that defines the range of cells containing names to be cleaned up (e.g., "E2:E").

### toTitlecase(sheet, range)

**Description:** This function converts the text in a specified range of the sheet to title case, i.e., capitalizing the first letter of each word.

**Parameters:**

- **sheet**: The sheet object where the range is located.
- **range**: A string that defines the range of cells containing the text to be converted to title case (e.g., "E2:E").

### toSentenceCase(sheet, range)

**Description:** This function converts the text in a specified range of the sheet to sentence case, i.e., capitalizing only the first letter of the text.

**Parameters:**

- **sheet**: The sheet object where the range is located.
- **range**: A string that defines the range of cells containing the text to be converted to sentence case (e.g., "K2:K").

### toLowercase(sheet, range)

**Description:** This function converts the text in a specified range of the sheet to lowercase.

**Parameters:**

- **sheet**: The sheet object where the range is located.
- **range**: A string that defines the range of cells containing the text to be converted to lowercase (e.g., "C2:D").

### toUppercase(sheet, range)

**Description:** This function converts the text in a specified range of the sheet to uppercase.

**Parameters:**

- **sheet**: The sheet object where the range is located.
- **range**: A string that defines the range of cells containing the text to be converted to uppercase (e.g., "Q2:Q").

### onFormSubmit()

**Description:** This function demonstrates how to use the utility functions to format specific cells in a sheet named "Responses". This function is intended to be triggered when a form is submitted.

**Usage:** This function does not take any parameters and is intended to be used as an example of how the utility functions can be combined and used together to format a specific sheet. 

## Example Usage

To use any of the functions, you need to call them from the script editor in Google Sheets. Here's an example of how to use the `correctAddress` function:

```javascript
// Get the active spreadsheet
var ss = SpreadsheetApp.getActiveSpreadsheet();

// Get the sheet named "Responses"
var sheet = ss.getSheetByName("Responses");

// Correct the addresses in the range "K2:K"
correctAddress(sheet, "K2:K");
```

Similarly, you can use the other functions by replacing `correctAddress` with the desired function name and updating the range accordingly.
