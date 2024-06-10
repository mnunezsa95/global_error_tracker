const sendEmail = (recipientList = "defaultUser@newglobe.education", subjectLine, message) => {
  const recipients = recipientList.length > 1 ? recipientList.join(",") : "defaultUser@newglobe.education";
  const subject = subjectLine;
  const body = message;

  GmailApp.sendEmail(recipients, subject, body);
};
