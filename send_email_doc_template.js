function sendEmail(tableId, rowId, docId, subject) {
  console.log("params:", tableId, rowId, docId, subject);

  // Sanity check
  if (!(docId && tableId && rowId)) {
    console.error("Missing or invalid doc ID, table ID, or row ID");
    return;
  }

  // Get the table row data
  const rowName = "tables/" + tableId + "/rows/" + rowId;
  const row = Area120Tables.Tables.Rows.get(rowName);
  const data = row.values;

  if (data["Status"] === "Lead") {
    // Get column data
    const firstName = data["First Name"];
    const recipient = data["Client"].toString();
    const body = createEmailBody(firstName, docId);
    console.log(firstName, recipient);
    try {
      // Generate email body from doc template
      MailApp.sendEmail({
        to: recipient,
        subject: subject,
        htmlBody: body,
      });
      row.values["Status"] = "Questionnaire Sent";
      Area120Tables.Tables.Rows.patch(row, rowName);
    } catch (err) {
      Logger.log("Failed with error %s", err.message);
    }
  }
}

function createEmailBody(firstName, docId) {
  // Make sure to update the emailTemplateDocId at the top.
  let emailBody = docToHtml(docId);
  emailBody = emailBody.replace(/{{FIRST NAME}}/g, name);
  return emailBody;
}

function docToHtml(docId) {
  let url =
    "https://docs.google.com/feeds/download/documents/export/Export?id=" +
    docId +
    "&exportFormat=html";
  let param = {
    method: "get",
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  };
  return UrlFetchApp.fetch(url, param).getContentText();
}