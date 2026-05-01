function sendCustomEmails() {
  // === CONFIGURATION ===
  const sheetName = "Sheet1"; // your sheet name
  const htmlFileName = "email-template"; // your HTML file name
  const statusColumn = 6; // Column F = Status (after removing Date column)

  // === LOAD SHEET ===
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();

  // === LOAD HTML TEMPLATE ===
  const htmlTemplate = HtmlService.createHtmlOutputFromFile(htmlFileName).getContent();

  // === FORMAT TODAY'S DATE ===
  const today = new Date();
  const formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "MMMM d, yyyy");

  // === LOOP THROUGH SHEET ROWS ===
  for (let i = 1; i < data.length; i++) {
    const [recipientName, recipientTitle, recipientAddress, salutation, email, status] = data[i];

    // Skip if already sent or missing email
    if (!email || status === "Sent") continue;

    // === REPLACE PLACEHOLDERS IN TEMPLATE ===
    let personalizedHtml = htmlTemplate
      .replace(/{{Date}}/g, formattedDate)
      .replace(/{{RecipientName}}/g, recipientName || "")
      .replace(/{{RecipientTitle}}/g, recipientTitle || "")
      .replace(/{{RecipientAddress}}/g, recipientAddress || "")
      .replace(/{{Salutation}}/g, salutation || "");

    // === SUBJECT LINE ===
    const subject = `Sponsorship Invitation: HAYAG - GANAP Performing Arts`;

    // === SEND EMAIL ===
    GmailApp.sendEmail(email, subject, "", {
      htmlBody: personalizedHtml,
      from: "donate@ganap.org.ph",
      name: "Donations",
      bcc: "gabo@ganap.org.ph"
    });

    // === UPDATE STATUS TO SENT ===
    sheet.getRange(i + 1, statusColumn).setValue("Sent");
    SpreadsheetApp.flush();
  }

  Logger.log("All emails sent successfully.");
}
