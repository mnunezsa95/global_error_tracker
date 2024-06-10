var updateEntireSpreadsheet = () => {
  /**
   * Executes three functions: importSpecificProgramErrors, updateGradeNames, and condenseSubjectNames,
   * followed by the classifyCore function.
   * Updates the date on the README Tab.
   * Logs the execution time for the entire update process.
   * Sends out an email to a list of users to notify them whether the update was successful or not.
   *
   * @function updateEntireSpreadsheet
   * @returns {void}
   */

  // Define the frequency of update and recipient list for email notification
  const frequencyOfUpdate = 2;
  const recipientList = ["defaultUser@newglobe.education", "user2@newglobe.education"];

  // Get the start time of the execution
  const startTime = new Date().getTime();

  // Access the README tab of the active spreadsheet
  const readmeTab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("README");

  // Get the current date and time in a readable format
  const currentDate = new Date().toLocaleString("en-US", {
    weekday: "short",
    year: "numeric",
    month: "long",
    day: "numeric",
    timeZoneName: "short",
    hour12: true,
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  });

  // Clear previous update information in the README tab
  readmeTab.getRange("L2:N2").clearContent();

  let executionTime;

  try {
    // Execute the functions to update the spreadsheet
    importSpecificProgramErrors();
    updateGradeNames();
    condenseSubjectNames();
    classifyCore();

    // Set the "Last Update" timestamp in the README tab
    readmeTab.getRange("K2:K2").setValue("Last Update");

    // Calculate execution time after the try block
    executionTime = new Date().getTime() - startTime;

    // Prepare success message for email notification
    const successMessage = `Successful Update of Global Academic Error Dashboard.\n\nUpdate Time: ${currentDate}\n\nExecution Time: ${executionTime} milliseconds`;

    // Log success message
    console.log(successMessage);

    // Send email notification for successful update
    sendEmail(recipientList, "Successful AET Update", successMessage);
  } catch (err) {
    // Set "Last Failed Updated" timestamp in the README tab
    readmeTab.getRange("K2:K2").setValue("Last Failed Updated");

    // Prepare failure message for email notification
    const failureMessage = `Unsuccessful Update of Global Academic Error Dashboard.\n\nError Message: ${err}\n\nWill try again in ${frequencyOfUpdate} Hour(s).`;

    // Log error message
    console.log(err);
    console.log(failureMessage);

    // Send email notification for unsuccessful update
    sendEmail(recipientList, "Unsuccessful AET Update", failureMessage);
  } finally {
    const endTime = new Date().getTime(); // Get the end time of the execution
    executionTime = endTime - startTime; // Calculate total execution time if it wasn't set in the try block
    readmeTab.getRange("N2:N2").setValue(currentDate); // Set the current date in the README tab
  }
};
