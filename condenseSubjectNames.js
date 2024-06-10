function condenseSubjectNames() {
  /**
   * Trims white space for each cell in the subject column &
   * Reduces the subject names to common groups as specified and returns these condensed subject names
   */

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Error Database [Aggregate]");
  let range = sheet.getRange("H:H");
  let gradeRange = sheet.getRange("G:G").getValues().flat();
  let values = range.getValues().map((row) => [row[0].trim()]);

  const gradesFourAndUp = ["Primary 4", "Primary 5", "Primary 6", "JSS 1", "JSS 2", "JSS 3", "Class 10"];

  for (let i = 0; i < values.length; i++) {
    let cellValue = values[i][0];
    let gradeValue = gradeRange[i];

    if (
      gradesFourAndUp.includes(gradeValue) &&
      (cellValue === "Basic Science and Technology" ||
        cellValue === "Basic Science and Technology - Basic Science" ||
        cellValue === "Science" ||
        cellValue === "Science & Technology")
    ) {
      sheet.getRange(i + 1, 8).setValue("Science (P4+)");
    } else if (
      cellValue === "BECE BST Prep" ||
      cellValue === "BECE English Prep" ||
      cellValue === "BECE Mathematics Prep" ||
      cellValue === "BECE National Values Prep" ||
      cellValue === "BECE Pre-Vocational Studies Prep"
    ) {
      sheet.getRange(i + 1, 8).setValue("BECE Prep");
    } else if (
      cellValue === "HSLC Prep - English" ||
      cellValue === "HSLC Prep - Mathematics" ||
      cellValue === "HSLC Prep - Science" ||
      cellValue === "HSLC Prep - Social Science"
    ) {
      sheet.getRange(i + 1, 8).setValue("HSLC Prep");
    } else if (
      cellValue === "KPSEA Prep Creative Arts" ||
      cellValue === "KPSEA Prep Creative Arts and Social Studies" ||
      cellValue === "KPSEA Prep English" ||
      cellValue === "KPSEA Prep Integrated Sciences" ||
      cellValue === "KPSEA Prep Kiswahili" ||
      cellValue === "KPSEA Prep Mathematics" ||
      cellValue === "KPSEA Prep Science & Technology" ||
      cellValue === "KPSEA Prep Social Studies"
    ) {
      sheet.getRange(i + 1, 8).setValue("KPSEA Prep");
    } else if (
      cellValue === "Co-curricular" ||
      cellValue === "Co-Curricular" ||
      cellValue === "Co Curricular" ||
      cellValue === "Cocurricular" ||
      cellValue === "Clubs"
    ) {
      sheet.getRange(i + 1, 8).setValue("Co-Curricular");
    } else if (cellValue.includes("English Studies") && cellValue.includes("Reading")) {
      sheet.getRange(i + 1, 8).setValue("English Studies - Reading");
    } else if (cellValue.includes("English Studies") && cellValue.includes("Language")) {
      sheet.getRange(i + 1, 8).setValue("English Studies - Language");
    } else if (
      cellValue === "Mathematics 1" ||
      cellValue === "Mathematics 2" ||
      cellValue === "Mathematics 3" ||
      cellValue.includes(" Mathematics") ||
      cellValue.includes("Mathematics 1 ") ||
      cellValue.includes("Mathematics 2 ")
    ) {
      sheet.getRange(i + 1, 8).setValue("Mathematics");
    } else if (cellValue === "Maths" || cellValue === "Math") {
      sheet.getRange(i + 1, 8).setValue("Maths");
    } else if (cellValue === "Supplementary English" || cellValue === "Supplemental English") {
      sheet.getRange(i + 1, 8).setValue("Supplementary English");
    } else if (cellValue === "Supplementary Maths" || cellValue === "Supplemental Maths") {
      sheet.getRange(i + 1, 8).setValue("Supplementary Maths");
    } else if (cellValue.includes("Preparatory English")) {
      sheet.getRange(i + 1, 8).setValue("Preparatory English");
    } else if (cellValue.includes("Preparatory Maths")) {
      sheet.getRange(i + 1, 8).setValue("Preparatory Maths");
    } else if (cellValue !== "Social Studies and Science" && cellValue.includes("Social Studies")) {
      sheet.getRange(i + 1, 8).setValue("Social Studies");
    } else if (
      cellValue.includes("Day") ||
      cellValue.includes("day") ||
      cellValue.includes(" Day") ||
      cellValue.includes(" Day ") ||
      cellValue.includes(" holiday") ||
      cellValue.includes(" holiday ")
    ) {
      sheet.getRange(i + 1, 8).setValue("Holiday Lesson(s)");
    }
  }
}
