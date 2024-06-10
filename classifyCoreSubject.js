function classifyCoreSubjects() {
  /**
   * Looks at Column I in the Error Database Tab (Lesson Code) and determines the appropriate level value based on the values in the column.
   * Trims leading and trailing whitespace from each cell value before comparison.
   * Compares the trimmed value to a list of predefined values and returns the corresponding level value to Column R (Core Subject / Level).
   *
   * @function classifyCore
   * @returns {void}
   */

  // Access the active spreadsheet and the 'Error Database [Aggregate]' sheet
  var currentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Error Database [Aggregate]");

  // Get the range of values in Column I (Lesson Code)
  var lessonCodeColumn = currentSheet.getRange("I2:I");

  // Get the values from the range and trim whitespace from each value
  var lessonCodeColValues = lessonCodeColumn.getValues().map((row) => row[0].trim());

  // Get the range where the output values will be placed (Column R)
  var destinationColumn = currentSheet.getRange("Q2:Q");

  // Initialize an array to store the output values
  var outputValues = [];

  // Iterate through each value in the Lesson Code column
  for (let i = 0; i < lessonCodeColValues.length; i++) {
    var cellValue = lessonCodeColValues[i];
    var level = "";

    // Determine the level based on the prefix of the Lesson Code
    if (cellValue.startsWith("LAL")) {
      level = "Reading Level A";
      if (cellValue.endsWith("_C")) level += " - Contingency";
    } else if (cellValue.startsWith("LBL")) {
      level = "Reading Level B";
      if (cellValue.endsWith("_C")) level += " - Contingency";
    } else if (cellValue.startsWith("LCL")) {
      level = "Reading Level C";
      if (cellValue.endsWith("_C")) level += " - Contingency";
    } else if (cellValue.startsWith("LDL")) {
      level = "Reading Level D";
      if (cellValue.endsWith("_C")) level += " - Contingency";
    } else if (cellValue.startsWith("LEL")) {
      level = "Reading Level E";
      if (cellValue.endsWith("_C")) level += " - Contingency";
    } else if (cellValue.startsWith("LALG")) {
      level = "Language Level A";
    } else if (cellValue.startsWith("LBLG")) {
      level = "Language Level B";
    } else if (cellValue.startsWith("LAN")) {
      level = "Mathematics Level A";
      if (cellValue.endsWith("_C")) level += " - Contingency";
    } else if (cellValue.startsWith("LBN")) {
      level = "Mathematics Level B";
      if (cellValue.endsWith("_C")) level += " - Contingency";
    } else if (cellValue.startsWith("LCN")) {
      level = "Mathematics Level C";
      if (cellValue.endsWith("_C")) level += " - Contingency";
    } else if (cellValue.startsWith("LDN")) {
      level = "Mathematics Level D";
      if (cellValue.endsWith("_C")) level += " - Contingency";
    } else if (cellValue.startsWith("LEN")) {
      level = "Mathematics Level E";
      if (cellValue.endsWith("_C")) level += " - Contingency";
    }

    // Add the determined level to the outputValues array
    outputValues.push([level]);
  }

  // Set the values in the destination column (Column R) to the outputValues array
  destinationColumn.setValues(outputValues);
}
