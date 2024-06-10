function importSpecificProgramErrors() {
  /**
   * Imports global data from all active programs
   * Clears content on the Error Database [Aggregate] tab & replaces it with imported global data
   * Adds a new column at beginning for program name
   */

  /**
      @function importSpecificProgramErrors
      @returns {void}
      @description
        * Imports global data from various active program spreadsheets into a central "Error Database [Aggregate]" sheet.
        * The function clears the existing content on the "Error Database [Aggregate]" tab and replaces it with the imported global data. It also adds a new column at the beginning of each row for the program name.
      @remarks
        * The function assumes that each source spreadsheet has a sheet named "Error Tracker" and that the data starts from the second row (A2) in column A. It dynamically determines the range of data based on the last non-empty row in column A.
      
      @param {Array} spreadsheets - Array of objects containing the IDs of the source spreadsheets and their corresponding program names.
        * Each object in the array has the following structure:
          {
            id: string,         // The ID of the source spreadsheet
            program: string     // The name of the program
          }
    
      @param {Sheet} destinationSheet - The sheet in the active spreadsheet where the aggregated data will be stored. It is assumed to be named "Error Database [Aggregate]".
      @param {Range} dataRange - The range of data to be copied from each source spreadsheet. It is determined dynamically based on the last non-empty row in column A.
      @param {Array} values - The values retrieved from the determined range in the source spreadsheet.
      @param {Array} newData - The data prepared to be appended to the destination sheet. It includes the program name as the first column for each row.
      @param {number} startRow - The starting row in the destination sheet where the data will be appended. It is updated after each dataset is appended.
     */

  // Array of objects containing spreadsheet IDs and program names
  let spreadsheets = [
    { id: "1JoWDTReHP4-BALEEcfWxrsY9kbbGVsZCBccP11mkC3U", program: "BayelsaPRIME" },
    { id: "1SN6oPBIQbaNIo0T1CD73WxSQLZLqK8SQ-UbUEfRbRSc", program: "Bridge Andhra Pradesh" },
    { id: "1Z49FAwkq8c0bSFPpG2Ki3Gh_t9XjiqvWov2D_1-rNgk", program: "Bridge Kenya" },
    { id: "1Gu_weq_OpTlQxfQHMvvTBSdXJnFCfwbd87_inX_m5As", program: "Bridge Liberia" },
    { id: "1qwq_D-aajIVYmX3BO0qMuCdbvTdEN-GaT5mfwFEhjKs", program: "Bridge Nigeria" },
    { id: "1AgntRauSd70NYGNcU_tgNxAKo-hfuub8NKaptsUbMTM", program: "Bridge Uganda" },
    { id: "1toB7ZluF72jxN9ku5UD6yq58Ci3OsMb59uxpJ0hJbOQ", program: "EdoBEST" },
    { id: "1uDu-qvOjGibyPqdmjclu3zn4A2wr-jQStE7M_sqK5DI", program: "EKOEXCEL" },
    { id: "1YFCb7aMA2MAuuyHZSwKEEtQ1CdZ5BF6uPJ3CbgzFOPo", program: "KwaraLEARN" },
    { id: "1upw1NRStBTYzNkJPViCBhwwj2PnJr6VwMcV8mDncibQ", program: "RwandaEQUIP" },
    { id: "1hoQl1qeK7C0C7hx1IDRMgMyiUvsreN7ji5wBVxDqDrM", program: "STAR Education" },
  ];

  // { id: '', program: 'ESPOIR République Centrafricaine' },
  // { id: '', program: 'LIFT Mizoram' },
  // { id: '', program: 'MeghalayaFIRST' },
  // { id: '', program: 'BéninINNOVE' },
  // { id: '', program: 'OyoSHINE' },
  // { id: '', program: 'Akwa Ibom RISE' },
  // { id: '', program: 'EnuguMEGA' },

  // Destination sheet to aggregate data into
  const destinationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Error Database [Aggregate]");

  // Clear existing content in columns A2:V in the destination sheet
  destinationSheet.getRange("A2:S").clearContent();

  // Initialize the starting row in the destination sheet
  let startRow = 2;

  // Loop through each spreadsheet object and copy data to the destination sheet
  for (var i = 0; i < spreadsheets.length; i++) {
    var spreadsheetId = spreadsheets[i].id;
    var programName = spreadsheets[i].program;

    // Open the source spreadsheet
    var sourceSheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName("Error Tracker");

    // Find the last row with data in column A
    var lastRow = sourceSheet.getRange("A:A").getValues().filter(String).length;

    // Determine the range dynamically based on the last row with data in column A
    var dataRange = sourceSheet.getRange("A2:Q" + lastRow); // Assuming the data starts from row 2 in column A

    // Get the values from the determined range
    var values = dataRange.getValues();

    // Prepare data to append to the destination sheet
    var newData = values.map(function (row) {
      return [programName].concat(row); // Insert program name in the beginning of each row
    });

    // Append newData to the destination sheet starting from the defined startRow
    destinationSheet.getRange(startRow, 1, newData.length, newData[0].length).setValues(newData);

    // Update the startRow for the next dataset to be appended
    startRow += newData.length;
  }
}
