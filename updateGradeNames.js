function updateGradeNames() {
  /**
   * Trims white space for each cell in the grade column &
   * Reduces the grade names to common groups as specified and returns these condensed subject names
   */

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Error Database [Aggregate]");
  const range = sheet.getRange("G2:G"); // Start from G2 to skip the header
  const values = range.getValues();

  // Create an array to hold updated values for the new column
  const updatedValues = values.map((row) => {
    const cellValue = row[0].trim();
    let newValue = cellValue; // Initialize with the original value

    if (cellValue === "Standard 1" || cellValue === "Grade 1") {
      newValue = "Primary 1";
    } else if (cellValue === "Standard 2" || cellValue === "Grade 2") {
      newValue = "Primary 2";
    } else if (cellValue === "Standard 3" || cellValue === "Grade 3") {
      newValue = "Primary 3";
    } else if (cellValue === "Standard 4" || cellValue === "Grade 4") {
      newValue = "Primary 4";
    } else if (cellValue === "Standard 5" || cellValue === "Grade 5") {
      newValue = "Primary 5";
    } else if (cellValue === "Grade 6" || cellValue === "Class 6") {
      newValue = "Primary 6";
    } else if (cellValue === "Class 7" || cellValue === "Grade 7" || cellValue === "Primary 7") {
      newValue = "JSS 1";
    } else if (cellValue === "Class 8") {
      newValue = "JSS 2";
    } else if (cellValue === "Class 9") {
      newValue = "JSS 3";
    } else if (cellValue === "Class 10" || cellValue === "Grade 10") {
      newValue = "Class 10";
    }

    return [newValue];
  });

  // Write the updated values back to the sheet in the same column (G), starting from G2 to match the original range
  sheet.getRange(2, 7, updatedValues.length, 1).setValues(updatedValues);
}
