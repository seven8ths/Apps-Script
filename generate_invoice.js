function generateInvoice(
  tableId,
  rowId,
  clientTableId,
  spreadsheetId,
  templateSheetId,
  templateSheetName,
  outputFolderId
) {
  console.log("params:", tableId, rowId, clientTableId);

  // Sanity check
  if (!(tableId && rowId)) {
    console.error("Missing or invalid table ID or row ID");
    return;
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const templateSheet = ss.getSheetByName(templateSheetName);

  // Get the table row data to generate the invoice
  const rowName = "tables/" + tableId + "/rows/" + rowId;
  const row = Area120Tables.Tables.Rows.get(rowName);
  const invoiceData = row.values;
  const clientData = getClientData(invoiceData["Client"], clientTableId);

  // Set template dates
  const todaysDate = new Date().toDateString();

  // Set column value variables
  const addOns = invoiceData["Add-ons"];

  // Clears existing data from the template.
  clearTemplateSheet(templateSheet);

  // Set template sheet values
  templateSheet.getRange("B9").setValue("Submitted on " + todaysDate);
  templateSheet.getRange("G15").setValue(calculateDueDate(invoiceData["Date"]));
  templateSheet.getRange("B12").setValue(clientData["Name"]);
  templateSheet.getRange("B13").setValue(invoiceData["Client"]);
  templateSheet.getRange("B14").setValue(clientData["Address"].address);
  templateSheet.getRange("G12").setValue(invoiceData["Invoice Number"]);
  templateSheet.getRange("G25").setValue(clientData["Adjustment"]);

  if (invoiceData["Package"] !== "None") {
    templateSheet.getRange("B18").setValue(invoiceData["Package"]);
    templateSheet.getRange("F18").setValue(invoiceData["Quantity"]);
    if (invoiceData["Session"] !== "None") {
      templateSheet.getRange("B19").setValue(invoiceData["Session"]);
      templateSheet.getRange("F19").setValue(invoiceData["Quantity"]);
      templateSheet.getRange("B20").setValue(invoiceData["Location"]);
      templateSheet.getRange("F20").setValue(invoiceData["Quantity"]);
      if (addOns) {
        templateSheet.getRange("B21").setValue(addOns[0]);
        templateSheet.getRange("B22").setValue(addOns[1]);
        templateSheet.getRange("B23").setValue(addOns[2]);
      }
    } else {
      templateSheet.getRange("B19").setValue(invoiceData["Location"]);
      templateSheet.getRange("F19").setValue(invoiceData["Quantity"]);
    }
  } else {
    templateSheet.getRange("B18").setValue(invoiceData["Session"]);
    templateSheet.getRange("F18").setValue(invoiceData["Quantity"]);
    templateSheet.getRange("B19").setValue(invoiceData["Location"]);
    templateSheet.getRange("F19").setValue(invoiceData["Quantity"]);
    if (addOns) {
      templateSheet.getRange("B20").setValue(addOns[0]);
      templateSheet.getRange("B21").setValue(addOns[1]);
      templateSheet.getRange("B22").setValue(addOns[2]);
    }
  }
  const pdf = spreadsheetToPDF(
    `Invoice#${invoiceData["Invoice Number"]}-${clientData["Name"]}`,
    spreadsheetId,
    templateSheetId,
    outputFolderId
  );
  console.log(pdf);
  invoiceData["Invoice Link"] = pdf.getUrl();
  Area120Tables.Tables.Rows.patch(row, rowName);
}

function calculateDueDate(date) {
  year = date.year;
  month = date.month;
  day = date.day - 2;
  dueDate = month + "/" + day + "/" + year;
  return dueDate;
}

function clearTemplateSheet(templateSheet) {
  // Clears existing data from the template.
  const rngClear = templateSheet
    .getRangeList(["B9", "B12:B14", "G12", "G15"])
    .getRanges();
  rngClear.forEach(function (cell) {
    cell.clearContent();
  });
  // This sample only accounts for six rows of data 'B18:G24'. You can extend or make dynamic as necessary.
  templateSheet.getRange(19, 2, 6, 1).clearContent();
  templateSheet.getRange("G25").setValue(0);
}

function spreadsheetToPDF(
  pdfName,
  spreadsheetId,
  templateSheetId,
  outputFolderId
) {
  SpreadsheetApp.flush();
  Utilities.sleep(500);

  //make the pdf from the sheet
  const fr = 0,
    fc = 0,
    lc = 9,
    lr = 28;
  const url =
    "https://docs.google.com/spreadsheets/d/" +
    spreadsheetId +
    "/export" +
    "?format=pdf&" +
    "size=7&" +
    "fzr=true&" +
    "portrait=true&" +
    "fitw=true&" +
    "gridlines=false&" +
    "printtitle=false&" +
    "top_margin=0.5&" +
    "bottom_margin=0.25&" +
    "left_margin=0.5&" +
    "right_margin=0.5&" +
    "sheetnames=false&" +
    "pagenum=UNDEFINED&" +
    "attachment=true&" +
    "gid=" +
    templateSheetId +
    "&" +
    "r1=" +
    fr +
    "&c1=" +
    fc +
    "&r2=" +
    lr +
    "&c2=" +
    lc;

  const params = {
    method: "GET",
    headers: { authorization: "Bearer " + ScriptApp.getOAuthToken() },
  };
  const blob = UrlFetchApp.fetch(url, params)
    .getBlob()
    .setName(pdfName + ".pdf");

  //save the file to folder on Drive
  var folder = DriveApp.getFolderById(outputFolderId);
  const pdfFile = folder.createFile(blob);
  return pdfFile;
}

function getClientData(client, clientTableId) {
  var tableID = clientTableId; // ID for the table
  var pageToken;
  var pageSize = 1000;
  var tableName = "tables/" + tableID;
  var response = Area120Tables.Tables.Rows.list(tableName, {
    filter: `values."Client"="${client}"`,
  });
  if (response) {
    for (var i = 0, rows = response.rows; i < rows.length; i++) {
      if (!rows[i].values) {
        // If blank row, keep going
        Logger.log("Empty row");
        continue;
      }
      Logger.log(rows[i].values);
      return rows[i].values;
    }
  }
}

function sendEmail() {
  ss.toast("Emailing Invoices", APP_TITLE, 1);
  invoices.forEach(function (invoice, index) {
    if (invoice.email_sent != "Yes") {
      ss.toast(`Emailing Invoice for ${invoice.customer}`, APP_TITLE, 1);

      const fileId = invoice.invoice_link.match(/[-\w]{25,}(?!.*[-\w]{25,})/);
      const attachment = DriveApp.getFileById(fileId);

      let recipient = invoice.email;
      if (EMAIL_OVERRIDE) {
        recipient = EMAIL_ADDRESS_OVERRIDE;
      }

      GmailApp.sendEmail(recipient, EMAIL_SUBJECT, EMAIL_BODY, {
        attachments: [attachment.getAs(MimeType.PDF)],
        name: APP_TITLE,
      });
      invoicesSheet.getRange(index + 2, 9).setValue("Yes");
    }
  });
}
